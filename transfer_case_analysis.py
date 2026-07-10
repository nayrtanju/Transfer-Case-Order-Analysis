import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Mapping, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.signal import savgol_filter


# Excel Open XML namespace
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Standard gravitational acceleration
G_TO_MS2 = 9.80665


# =============================================================================
# TRANSFER CASE TARGET DEFINITIONS
# =============================================================================

TRANSFER_CASE_TARGET_RPM = np.array(
    [
        1000,
        1500,
        2000,
        2500,
        3000,
        3500,
        4000,
        4500,
    ],
    dtype=float,
)


TRANSFER_CASE_ORDERS = {
    63.0: {
        "label": "63.00 Order - Gear Mesh",
        "harmonic": "1st",
        "target_rpm": TRANSFER_CASE_TARGET_RPM,
        "target_amp": np.array(
            [
                5.0,
                7.5,
                10.0,
                12.5,
                15.0,
                17.5,
                20.0,
                22.5,
            ],
            dtype=float,
        ),
    },

    85.05: {
        "label": "85.05 Order - Gear Mesh",
        "harmonic": "1st",
        "target_rpm": TRANSFER_CASE_TARGET_RPM,
        "target_amp": np.array(
            [
                2.0,
                3.0,
                4.0,
                5.0,
                6.0,
                7.0,
                8.0,
                9.0,
            ],
            dtype=float,
        ),
    },

    126.0: {
        "label": "126.00 Order - 2nd Harmonic",
        "harmonic": "2nd",
        "target_rpm": None,
        "target_amp": None,
    },

    170.10: {
        "label": "170.10 Order - 2nd Harmonic",
        "harmonic": "2nd",
        "target_rpm": None,
        "target_amp": None,
    },
}


# =============================================================================
# XLSX READER HELPERS
# =============================================================================

def load_shared_strings(
    archive: zipfile.ZipFile
) -> list[str]:
    """
    Read Excel shared strings.

    XLSX files may store text values such as column headers inside
    xl/sharedStrings.xml. This function returns those strings as a list.
    """
    shared_strings: list[str] = []

    shared_strings_path = "xl/sharedStrings.xml"

    if shared_strings_path not in archive.namelist():
        return shared_strings

    with archive.open(shared_strings_path) as stream:
        for _, element in ET.iterparse(
            stream,
            events=("end",),
        ):
            if element.tag == NS + "si":
                text_parts = []

                for text_element in element.iter(NS + "t"):
                    text_parts.append(
                        text_element.text or ""
                    )

                shared_strings.append(
                    "".join(text_parts)
                )

                element.clear()

    return shared_strings


def col_index(
    cell_reference: str
) -> int:
    """
    Convert an Excel column reference to a zero-based index.

    Examples:
        A -> 0
        B -> 1
        E -> 4
        AA -> 26
    """
    match = re.match(
        r"([A-Z]+)",
        cell_reference,
    )

    if match is None:
        raise ValueError(
            f"Invalid Excel cell reference: {cell_reference}"
        )

    letters = match.group(1)

    column_number = 0

    for character in letters:
        column_number = (
            column_number * 26
            + ord(character)
            - 64
        )

    return column_number - 1


def _get_worksheet_row_count(
    archive: zipfile.ZipFile,
    worksheet_path: str
) -> int:
    """
    Estimate worksheet row count from the Excel dimension field.

    The returned value excludes the first header row.
    """
    row_count = 0

    with archive.open(worksheet_path) as stream:
        for _, element in ET.iterparse(
            stream,
            events=("start",),
        ):
            if element.tag == NS + "dimension":
                reference = element.attrib.get(
                    "ref",
                    "",
                )

                match = re.search(
                    r":([A-Z]+)(\d+)",
                    reference,
                )

                if match is not None:
                    last_row = int(
                        match.group(2)
                    )

                    row_count = max(
                        0,
                        last_row - 1,
                    )

                break

    return row_count


def read_xlsx_numeric(
    path: str,
    max_columns: int = 5
) -> Tuple[list, np.ndarray]:
    """
    Read the first five columns from the first worksheet of an XLSX file.

    Expected column structure:

        Column 0 -> Time
        Column 1 -> ChA
        Column 2 -> ChB
        Column 3 -> ChC
        Column 4 -> RPM

    The function uses XML streaming instead of pandas/openpyxl so that
    large measurement files can be read with lower memory usage.
    """
    worksheet_path = "xl/worksheets/sheet1.xml"

    with zipfile.ZipFile(path) as archive:

        if worksheet_path not in archive.namelist():
            raise ValueError(
                "The Excel file does not contain sheet1.xml."
            )

        shared_strings = load_shared_strings(
            archive
        )

        estimated_rows = _get_worksheet_row_count(
            archive,
            worksheet_path,
        )

        if estimated_rows <= 0:
            raise ValueError(
                "The Excel worksheet does not contain data rows."
            )

        # Allocate based on the worksheet dimension.
        # Additional rows are added dynamically if the dimension is inaccurate.
        data = np.full(
            (
                estimated_rows,
                max_columns,
            ),
            np.nan,
            dtype=np.float64,
        )

        headers = [None] * max_columns
        data_row_index = 0

        with archive.open(worksheet_path) as stream:
            for _, row_element in ET.iterparse(
                stream,
                events=("end",),
            ):
                if row_element.tag != NS + "row":
                    continue

                excel_row_number = int(
                    row_element.attrib.get(
                        "r",
                        "0",
                    )
                )

                row_values = (
                    [np.nan] * max_columns
                )

                for cell_element in row_element.findall(
                    NS + "c"
                ):
                    cell_reference = (
                        cell_element.attrib.get(
                            "r",
                            "",
                        )
                    )

                    try:
                        column_index = col_index(
                            cell_reference
                        )
                    except ValueError:
                        continue

                    if column_index >= max_columns:
                        continue

                    cell_type = (
                        cell_element.attrib.get(
                            "t"
                        )
                    )

                    value_element = (
                        cell_element.find(
                            NS + "v"
                        )
                    )

                    if (
                        value_element is None
                        or value_element.text is None
                    ):
                        continue

                    raw_value = value_element.text

                    # First row is treated as the header.
                    if excel_row_number == 1:

                        if cell_type == "s":
                            try:
                                shared_index = int(
                                    raw_value
                                )

                                if (
                                    0
                                    <= shared_index
                                    < len(shared_strings)
                                ):
                                    headers[column_index] = (
                                        shared_strings[
                                            shared_index
                                        ]
                                    )
                                else:
                                    headers[column_index] = (
                                        raw_value
                                    )

                            except ValueError:
                                headers[column_index] = (
                                    raw_value
                                )

                        else:
                            headers[column_index] = (
                                raw_value
                            )

                    else:
                        try:
                            row_values[column_index] = float(
                                raw_value
                            )

                        except (
                            TypeError,
                            ValueError,
                        ):
                            row_values[column_index] = (
                                np.nan
                            )

                if excel_row_number > 1:

                    # Grow the array when Excel dimension metadata
                    # underestimates the actual row count.
                    if data_row_index >= data.shape[0]:
                        extension = np.full(
                            (
                                10000,
                                max_columns,
                            ),
                            np.nan,
                            dtype=np.float64,
                        )

                        data = np.vstack(
                            [
                                data,
                                extension,
                            ]
                        )

                    data[
                        data_row_index,
                        :
                    ] = row_values

                    data_row_index += 1

                row_element.clear()

        if data_row_index == 0:
            raise ValueError(
                "No numeric measurement rows were found "
                "in the Excel worksheet."
            )

        data = data[
            :data_row_index,
            :
        ]

        return headers, data
        # =============================================================================
# TIME VECTOR REPAIR AND ANGULAR RESAMPLING
# =============================================================================

def _validate_raw_vectors(
    time: np.ndarray,
    rpm: np.ndarray,
    signal: np.ndarray
) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Validate and clean raw input vectors.

    Invalid rows are removed only when they contain:
        - NaN
        - infinite values
        - non-positive RPM

    Repeated time values are preserved and repaired later.
    """
    time = np.asarray(
        time,
        dtype=float,
    )

    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    signal = np.asarray(
        signal,
        dtype=float,
    )

    if not (
        len(time)
        == len(rpm)
        == len(signal)
    ):
        raise ValueError(
            "Time, RPM and signal vectors must have the same length."
        )

    valid_mask = (
        np.isfinite(time)
        & np.isfinite(rpm)
        & np.isfinite(signal)
        & (rpm > 0)
    )

    time = time[valid_mask]
    rpm = rpm[valid_mask]
    signal = signal[valid_mask]

    if len(time) < 3:
        raise ValueError(
            "Not enough valid samples are available for analysis."
        )

    if np.any(
        np.diff(time) < 0
    ):
        raise ValueError(
            "Time values must not decrease."
        )

    return time, rpm, signal


def repair_time_vector_without_dropping_rows(
    time: np.ndarray
) -> np.ndarray:
    """
    Create a strictly increasing time vector without deleting samples.

    Pico CSV exports may contain repeated time values because time is
    written with limited decimal precision.

    Example:

        10.7905
        10.7905
        10.7906

    The repeated timestamp is repaired using the median positive time step.
    No measurement row is deleted.
    """
    time = np.asarray(
        time,
        dtype=float,
    )

    if len(time) < 2:
        raise ValueError(
            "At least two time samples are required."
        )

    time_difference = np.diff(
        time
    )

    if np.any(
        time_difference < 0
    ):
        raise ValueError(
            "Time values must not decrease."
        )

    positive_time_steps = time_difference[
        np.isfinite(time_difference)
        & (time_difference > 0)
    ]

    if len(positive_time_steps) == 0:
        raise ValueError(
            "The time vector contains no positive time increments. "
            "Please increase the export time precision."
        )

    median_time_step = float(
        np.median(
            positive_time_steps
        )
    )

    if (
        not np.isfinite(median_time_step)
        or median_time_step <= 0
    ):
        raise ValueError(
            "A valid sample interval could not be estimated."
        )

    repaired_time_steps = np.where(
        np.isfinite(time_difference)
        & (time_difference > 0),
        time_difference,
        median_time_step,
    )

    repaired_time = np.empty_like(
        time,
        dtype=float,
    )

    repaired_time[0] = time[0]

    repaired_time[1:] = (
        time[0]
        + np.cumsum(
            repaired_time_steps
        )
    )

    if not np.all(
        np.diff(repaired_time) > 0
    ):
        raise ValueError(
            "The repaired time vector is not strictly increasing."
        )

    return repaired_time


def angular_resample(
    time: np.ndarray,
    rpm: np.ndarray,
    signal: np.ndarray,
    samples_per_rev: int = 512
) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Convert a time-domain vibration signal to the angular domain.

    Parameters
    ----------
    time:
        Time vector in seconds.

    rpm:
        Rotational speed vector in revolutions per minute.

    signal:
        Vibration signal in m/s².

    samples_per_rev:
        Number of angular samples per revolution.

    Returns
    -------
    theta_uniform:
        Uniform angular-position vector in radians.

    signal_uniform:
        Vibration signal interpolated to the uniform angular grid.

    rpm_uniform:
        RPM interpolated to the same angular grid.
    """
    if samples_per_rev < 2:
        raise ValueError(
            "samples_per_rev must be at least 2."
        )

    time, rpm, signal = _validate_raw_vectors(
        time,
        rpm,
        signal,
    )

    repaired_time = repair_time_vector_without_dropping_rows(
        time
    )

    delta_time = np.diff(
        repaired_time,
        prepend=repaired_time[0],
    )

    positive_delta_time = delta_time[
        np.isfinite(delta_time)
        & (delta_time > 0)
    ]

    if len(positive_delta_time) == 0:
        raise ValueError(
            "A positive time step could not be calculated."
        )

    median_delta_time = float(
        np.median(
            positive_delta_time
        )
    )

    delta_time[0] = median_delta_time

    delta_time = np.where(
        np.isfinite(delta_time)
        & (delta_time > 0),
        delta_time,
        median_delta_time,
    )

    angular_velocity = (
        2.0
        * np.pi
        * rpm
        / 60.0
    )

    angular_increment = (
        angular_velocity
        * delta_time
    )

    theta = np.cumsum(
        angular_increment
    )

    increasing_mask = np.r_[
        True,
        np.diff(theta) > 0
    ]

    theta = theta[
        increasing_mask
    ]

    signal = signal[
        increasing_mask
    ]

    rpm = rpm[
        increasing_mask
    ]

    if len(theta) < 2:
        raise ValueError(
            "Angular position does not contain enough increasing samples."
        )

    if theta[-1] <= theta[0]:
        raise ValueError(
            "The available angular travel is insufficient."
        )

    angular_step = (
        2.0
        * np.pi
        / float(samples_per_rev)
    )

    theta_uniform = np.arange(
        theta[0],
        theta[-1],
        angular_step,
    )

    minimum_required_samples = (
        samples_per_rev * 2
    )

    if len(theta_uniform) < minimum_required_samples:
        available_revolutions = (
            len(theta_uniform)
            / float(samples_per_rev)
        )

        raise ValueError(
            "Insufficient measurement duration for angular resampling. "
            f"Available revolutions: {available_revolutions:.2f}. "
            "At least 2 complete revolutions are required."
        )

    signal_uniform = np.interp(
        theta_uniform,
        theta,
        signal,
    )

    rpm_uniform = np.interp(
        theta_uniform,
        theta,
        rpm,
    )

    if not np.all(
        np.isfinite(signal_uniform)
    ):
        raise ValueError(
            "Angular-resampled signal contains invalid values."
        )

    if not np.all(
        np.isfinite(rpm_uniform)
    ):
        raise ValueError(
            "Angular-resampled RPM contains invalid values."
        )

    return (
        theta_uniform,
        signal_uniform,
        rpm_uniform,
    )
    # =============================================================================
# ORDER MAP
# =============================================================================

def order_map(
    theta_u: np.ndarray,
    signal_u: np.ndarray,
    rpm_u: np.ndarray,
    samples_per_rev: int = 512,
    revs_per_block: int = 20,
    overlap: float = 0.75,
    max_order: float = 200.0,
):
    """
    Transfer Case Order Map

    FFT is calculated in the angular domain.

    Default settings

        Samples / Rev = 512
        Revs / Block = 20

    giving

        ΔOrder = 1 / RevsPerBlock

               = 0.05 order

    which allows accurate extraction of

        63.00

        85.05

        126.00

        170.10
    """

    theta_u = np.asarray(theta_u, dtype=float)
    signal_u = np.asarray(signal_u, dtype=float)
    rpm_u = np.asarray(rpm_u, dtype=float)

    if len(theta_u) != len(signal_u):
        raise ValueError(
            "Theta and signal length mismatch."
        )

    if len(theta_u) != len(rpm_u):
        raise ValueError(
            "Theta and RPM length mismatch."
        )

    if samples_per_rev < 16:
        raise ValueError(
            "Samples / Rev is too low."
        )

    nyquist_order = samples_per_rev / 2.0

    if max_order >= nyquist_order:
        raise ValueError(
            f"Maximum available order is {nyquist_order:.1f}"
        )

    samples_per_block = int(
        samples_per_rev * revs_per_block
    )

    if len(signal_u) < samples_per_block:
        raise ValueError(
            "Measurement is shorter than one FFT block."
        )

    hop = int(
        samples_per_block * (1.0 - overlap)
    )

    hop = max(
        hop,
        1,
    )

    window = np.hanning(
        samples_per_block
    )

    window_sum = np.sum(window)

    fft_orders = np.fft.rfftfreq(
        samples_per_block,
        d=1.0 / samples_per_rev,
    )

    keep = fft_orders <= max_order

    fft_orders = fft_orders[keep]

    spectra = []

    average_rpm = []

    for start in range(
        0,
        len(signal_u) - samples_per_block + 1,
        hop,
    ):

        stop = start + samples_per_block

        block = signal_u[start:stop]

        block_rpm = rpm_u[start:stop]

        if len(block) != samples_per_block:
            continue

        if np.any(~np.isfinite(block)):
            continue

        block = block - np.mean(block)

        spectrum = np.fft.rfft(
            block * window
        )

        amplitude = (
            np.sqrt(2.0)
            * np.abs(spectrum)
            / window_sum
        )

        spectra.append(
            amplitude[keep]
        )

        average_rpm.append(
            np.mean(block_rpm)
        )

    if len(spectra) == 0:
        raise ValueError(
            "No FFT blocks could be created."
        )

    spectra = np.vstack(
        spectra
    )

    average_rpm = np.asarray(
        average_rpm,
        dtype=float,
    )

    return (
        fft_orders,
        average_rpm,
        spectra,
    )
    # =============================================================================
# ORDER EXTRACTION AND RPM RESAMPLING
# =============================================================================

def smooth_curve(
    values: np.ndarray,
    window_length: int = 9,
    polyorder: int = 2
) -> np.ndarray:
    """
    Smooth an amplitude curve using a Savitzky-Golay filter.

    The function automatically adjusts the window length for short datasets.
    """
    values = np.asarray(
        values,
        dtype=float,
    )

    if len(values) < 5:
        return values

    finite_mask = np.isfinite(
        values
    )

    if not np.all(finite_mask):
        finite_indices = np.flatnonzero(
            finite_mask
        )

        if len(finite_indices) < 2:
            return values

        values = np.interp(
            np.arange(len(values)),
            finite_indices,
            values[finite_indices],
        )

    window_length = int(
        window_length
    )

    polyorder = int(
        polyorder
    )

    if window_length % 2 == 0:
        window_length += 1

    if window_length >= len(values):
        window_length = len(values) - 1

    if window_length % 2 == 0:
        window_length -= 1

    minimum_window = polyorder + 2

    if minimum_window % 2 == 0:
        minimum_window += 1

    if window_length < minimum_window:
        return values

    return savgol_filter(
        values,
        window_length=window_length,
        polyorder=polyorder,
    )


def average_duplicate_rpm_values(
    rpm: np.ndarray,
    amplitude: np.ndarray
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Average amplitude values belonging to identical RPM values.

    This prevents interpolation instability when multiple FFT blocks have the
    same mean RPM.
    """
    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    amplitude = np.asarray(
        amplitude,
        dtype=float,
    )

    unique_rpm, inverse_indices = np.unique(
        rpm,
        return_inverse=True,
    )

    if len(unique_rpm) == len(rpm):
        return rpm, amplitude

    amplitude_sum = np.zeros(
        len(unique_rpm),
        dtype=float,
    )

    amplitude_count = np.zeros(
        len(unique_rpm),
        dtype=float,
    )

    np.add.at(
        amplitude_sum,
        inverse_indices,
        amplitude,
    )

    np.add.at(
        amplitude_count,
        inverse_indices,
        1.0,
    )

    averaged_amplitude = (
        amplitude_sum
        / np.maximum(
            amplitude_count,
            1.0,
        )
    )

    return (
        unique_rpm,
        averaged_amplitude,
    )


def resample_to_rpm_step(
    rpm: np.ndarray,
    amplitude: np.ndarray,
    rpm_step: float = 10.0
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Resample an order amplitude curve to a fixed RPM grid.
    """
    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    amplitude = np.asarray(
        amplitude,
        dtype=float,
    )

    if rpm_step <= 0:
        raise ValueError(
            "rpm_step must be greater than zero."
        )

    valid_mask = (
        np.isfinite(rpm)
        & np.isfinite(amplitude)
    )

    rpm = rpm[
        valid_mask
    ]

    amplitude = amplitude[
        valid_mask
    ]

    if len(rpm) < 2:
        raise ValueError(
            "Not enough valid RPM blocks are available for resampling."
        )

    sort_indices = np.argsort(
        rpm,
        kind="stable",
    )

    rpm = rpm[
        sort_indices
    ]

    amplitude = amplitude[
        sort_indices
    ]

    rpm, amplitude = average_duplicate_rpm_values(
        rpm,
        amplitude,
    )

    if len(rpm) < 2:
        raise ValueError(
            "RPM range contains fewer than two unique points."
        )

    rpm_minimum = (
        np.ceil(
            rpm[0] / rpm_step
        )
        * rpm_step
    )

    rpm_maximum = (
        np.floor(
            rpm[-1] / rpm_step
        )
        * rpm_step
    )

    if rpm_maximum <= rpm_minimum:
        raise ValueError(
            "RPM range is too narrow for the selected RPM step."
        )

    rpm_grid = np.arange(
        rpm_minimum,
        rpm_maximum + rpm_step,
        rpm_step,
    )

    amplitude_grid = np.interp(
        rpm_grid,
        rpm,
        amplitude,
    )

    return (
        rpm_grid,
        amplitude_grid,
    )


def interpolate_fractional_order(
    orders: np.ndarray,
    spectrum: np.ndarray,
    target_order: float
) -> np.ndarray:
    """
    Interpolate a fractional order from each FFT block.

    This is used when the requested order does not fall exactly on an FFT bin.
    """
    orders = np.asarray(
        orders,
        dtype=float,
    )

    spectrum = np.asarray(
        spectrum,
        dtype=float,
    )

    if spectrum.ndim != 2:
        raise ValueError(
            "Spectrum must be two-dimensional."
        )

    interpolated_amplitude = np.empty(
        spectrum.shape[0],
        dtype=float,
    )

    for block_index in range(
        spectrum.shape[0]
    ):
        interpolated_amplitude[
            block_index
        ] = np.interp(
            target_order,
            orders,
            spectrum[block_index, :],
        )

    return interpolated_amplitude


def extract_order_vs_rpm(
    orders: np.ndarray,
    rpms: np.ndarray,
    spectrum: np.ndarray,
    target_order: float,
    width: float = 0.15,
    rpm_step: float = 10.0,
    smooth: bool = True
) -> Tuple[np.ndarray, np.ndarray]:
    """
    Extract an order-amplitude curve versus RPM.

    The selected order band is integrated using RSS:

        amplitude = sqrt(sum(band_amplitudes²))

    If no FFT bin falls inside the requested band, the function performs
    fractional-order interpolation.
    """
    orders = np.asarray(
        orders,
        dtype=float,
    )

    rpms = np.asarray(
        rpms,
        dtype=float,
    )

    spectrum = np.asarray(
        spectrum,
        dtype=float,
    )

    if spectrum.ndim != 2:
        raise ValueError(
            "Order spectrum must be a two-dimensional array. "
            f"Received shape: {spectrum.shape}."
        )

    if len(orders) == 0:
        raise ValueError(
            "Order axis is empty."
        )

    if len(rpms) == 0:
        raise ValueError(
            "RPM vector is empty."
        )

    if spectrum.shape[0] != len(rpms):
        raise ValueError(
            "Spectrum row count does not match the RPM vector length."
        )

    if spectrum.shape[1] != len(orders):
        raise ValueError(
            "Spectrum column count does not match the order-axis length."
        )

    if target_order < orders[0]:
        raise ValueError(
            f"Target order {target_order:.2f} is below the "
            f"minimum calculated order {orders[0]:.2f}."
        )

    if target_order > orders[-1]:
        raise ValueError(
            f"Target order {target_order:.2f} is above the "
            f"maximum calculated order {orders[-1]:.2f}. "
            "Increase Max Order."
        )

    if width <= 0:
        raise ValueError(
            "Order width must be greater than zero."
        )

    half_width = width / 2.0

    band_mask = (
        (orders >= target_order - half_width)
        &
        (orders <= target_order + half_width)
    )

    if np.any(
        band_mask
    ):
        amplitude = np.sqrt(
            np.sum(
                spectrum[:, band_mask] ** 2,
                axis=1,
            )
        )

    else:
        amplitude = interpolate_fractional_order(
            orders,
            spectrum,
            target_order,
        )

    valid_mask = (
        np.isfinite(rpms)
        & np.isfinite(amplitude)
    )

    rpm_values = rpms[
        valid_mask
    ]

    amplitude_values = amplitude[
        valid_mask
    ]

    if len(rpm_values) < 2:
        raise ValueError(
            f"Not enough valid RPM blocks are available for "
            f"order {target_order:.2f}."
        )

    sort_indices = np.argsort(
        rpm_values,
        kind="stable",
    )

    rpm_sorted = rpm_values[
        sort_indices
    ]

    amplitude_sorted = amplitude_values[
        sort_indices
    ]

    if smooth:
        amplitude_sorted = smooth_curve(
            amplitude_sorted,
            window_length=9,
            polyorder=2,
        )

    rpm_resampled, amplitude_resampled = resample_to_rpm_step(
        rpm_sorted,
        amplitude_sorted,
        rpm_step=rpm_step,
    )

    return (
        rpm_resampled,
        amplitude_resampled,
    )
    # =============================================================================
# TARGET EVALUATION AND PASS / FAIL LOGIC
# =============================================================================

def integrate_positive_area(
    rpm: np.ndarray,
    difference: np.ndarray
) -> float:
    """
    Integrate only the portion of a curve that is above zero.

    This is used to calculate the total area where the measured order curve
    exceeds the target curve.

    Unit:

        m/s² · RPM
    """
    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    difference = np.asarray(
        difference,
        dtype=float,
    )

    if len(rpm) != len(difference):
        raise ValueError(
            "RPM and difference arrays must have the same length."
        )

    if len(rpm) < 2:
        return 0.0

    valid_mask = (
        np.isfinite(rpm)
        & np.isfinite(difference)
    )

    rpm = rpm[
        valid_mask
    ]

    difference = difference[
        valid_mask
    ]

    if len(rpm) < 2:
        return 0.0

    positive_difference = np.maximum(
        difference,
        0.0,
    )

    if hasattr(
        np,
        "trapezoid",
    ):
        area = np.trapezoid(
            positive_difference,
            rpm,
        )

    else:
        area = np.trapz(
            positive_difference,
            rpm,
        )

    return float(area)


def evaluate_curve_against_target(
    rpm: np.ndarray,
    amplitude: np.ndarray,
    target_rpm: Optional[np.ndarray],
    target_amp: Optional[np.ndarray]
) -> dict:
    """
    Evaluate one order curve against its target.

    Returned metrics:

        Peak RPM
        Peak Amplitude
        Target at Peak RPM
        Max Margin
        Max Margin %
        Exceedance Area
        Status

    Target-free harmonic orders return INFO status.
    """
    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    amplitude = np.asarray(
        amplitude,
        dtype=float,
    )

    if len(rpm) != len(amplitude):
        raise ValueError(
            "RPM and amplitude arrays must have the same length."
        )

    valid_mask = (
        np.isfinite(rpm)
        & np.isfinite(amplitude)
    )

    rpm = rpm[
        valid_mask
    ]

    amplitude = amplitude[
        valid_mask
    ]

    if len(rpm) == 0:
        raise ValueError(
            "No valid order curve data are available for evaluation."
        )

    peak_index = int(
        np.argmax(
            amplitude
        )
    )

    peak_rpm = float(
        rpm[
            peak_index
        ]
    )

    peak_amplitude = float(
        amplitude[
            peak_index
        ]
    )

    has_target = (
        target_rpm is not None
        and target_amp is not None
    )

    if not has_target:
        return {
            "Peak RPM": peak_rpm,
            "Peak Amplitude [m/s²]": peak_amplitude,
            "Target at Peak RPM [m/s²]": np.nan,
            "Max Margin [m/s²]": np.nan,
            "Max Margin [%]": np.nan,
            "Exceedance Area [m/s²·RPM]": np.nan,
            "Status": "INFO",
        }

    target_rpm = np.asarray(
        target_rpm,
        dtype=float,
    )

    target_amp = np.asarray(
        target_amp,
        dtype=float,
    )

    if len(target_rpm) != len(target_amp):
        raise ValueError(
            "Target RPM and target amplitude arrays must have the same length."
        )

    if len(target_rpm) < 2:
        raise ValueError(
            "At least two target points are required."
        )

    target_valid_mask = (
        np.isfinite(target_rpm)
        & np.isfinite(target_amp)
    )

    target_rpm = target_rpm[
        target_valid_mask
    ]

    target_amp = target_amp[
        target_valid_mask
    ]

    if len(target_rpm) < 2:
        raise ValueError(
            "Target curve contains fewer than two valid points."
        )

    target_sort_indices = np.argsort(
        target_rpm,
        kind="stable",
    )

    target_rpm = target_rpm[
        target_sort_indices
    ]

    target_amp = target_amp[
        target_sort_indices
    ]

    target_curve = np.interp(
        rpm,
        target_rpm,
        target_amp,
    )

    target_at_peak_rpm = float(
        np.interp(
            peak_rpm,
            target_rpm,
            target_amp,
        )
    )

    margin_curve = (
        amplitude
        - target_curve
    )

    max_margin_index = int(
        np.argmax(
            margin_curve
        )
    )

    max_margin = float(
        margin_curve[
            max_margin_index
        ]
    )

    target_at_max_margin = float(
        target_curve[
            max_margin_index
        ]
    )

    max_margin_percent = (
        max_margin
        / target_at_max_margin
        * 100.0
        if target_at_max_margin > 0
        else np.nan
    )

    exceedance_area = integrate_positive_area(
        rpm,
        margin_curve,
    )

    # Any positive exceedance area means at least part of the curve
    # is above the target.
    status = (
        "PASS"
        if exceedance_area <= 1e-9
        else "FAIL"
    )

    return {
        "Peak RPM": peak_rpm,
        "Peak Amplitude [m/s²]": peak_amplitude,
        "Target at Peak RPM [m/s²]": target_at_peak_rpm,
        "Max Margin [m/s²]": max_margin,
        "Max Margin [%]": max_margin_percent,
        "Exceedance Area [m/s²·RPM]": exceedance_area,
        "Status": status,
    }


def calculate_overall_status(
    results_by_order: Mapping[
        float,
        pd.DataFrame
    ]
) -> str:
    """
    Calculate overall Transfer Case analysis status.

    INFO rows are excluded from PASS / FAIL logic.

    Rules:

        Any evaluated FAIL row -> FAIL
        All evaluated rows PASS -> PASS
        No evaluated rows -> INFO
    """
    evaluated_statuses = []

    for result_dataframe in results_by_order.values():

        if (
            "Status"
            not in result_dataframe.columns
        ):
            raise ValueError(
                "Result dataframe does not contain a Status column."
            )

        valid_status_rows = result_dataframe[
            result_dataframe[
                "Status"
            ] != "INFO"
        ]

        if len(valid_status_rows) > 0:
            evaluated_statuses.extend(
                valid_status_rows[
                    "Status"
                ].tolist()
            )

    if len(evaluated_statuses) == 0:
        return "INFO"

    if any(
        status == "FAIL"
        for status in evaluated_statuses
    ):
        return "FAIL"

    return "PASS"
      # =============================================================================
# MAIN TRANSFER CASE ANALYSIS API
# =============================================================================

def analyze_transfer_case_orders(
    time: np.ndarray,
    rpm: np.ndarray,
    channels: Mapping[str, np.ndarray],
    order_definitions: Optional[
        Mapping[float, dict]
    ] = None,
    samples_per_rev: int = 512,
    revs_per_block: int = 20,
    overlap: float = 0.75,
    max_order: float = 200.0,
    order_width: float = 0.15,
    rpm_step: float = 10.0,
    calibration_factor: float = 1.0
) -> Tuple[
    Dict[float, dict],
    Dict[float, pd.DataFrame],
    Dict[float, pd.DataFrame],
]:
    """
    Run the complete Transfer Case Gear Mesh order analysis.

    Parameters
    ----------
    time:
        Time vector in seconds.

    rpm:
        RPM vector.

    channels:
        Dictionary containing vibration channels.

        Expected structure:

            {
                "ChA": np.ndarray,
                "ChB": np.ndarray,
                "ChC": np.ndarray,
            }

    order_definitions:
        Order and target definitions.

        If None, TRANSFER_CASE_ORDERS is used.

    samples_per_rev:
        Angular samples per revolution.

    revs_per_block:
        Number of revolutions per FFT block.

        Recommended value for Transfer Case:

            20 revolutions

        This gives:

            Delta Order = 1 / 20 = 0.05 order

    overlap:
        FFT block overlap ratio.

    max_order:
        Maximum calculated order.

    order_width:
        Width of the order extraction band.

    rpm_step:
        Output RPM grid step.

    calibration_factor:
        Final amplitude multiplier.

    Returns
    -------
    curves_by_order:
        Nested dictionary containing order curves.

        Structure:

            {
                63.0: {
                    "ChA": {
                        "rpm": ...,
                        "amp": ...,
                    },
                    "ChB": {...},
                    "ChC": {...},
                },
                ...
            }

    results_by_order:
        Dictionary containing one result dataframe per order.

    raw_curves_by_order:
        Dictionary containing one combined curve dataframe per order.
    """
    if order_definitions is None:
        order_definitions = (
            TRANSFER_CASE_ORDERS
        )

    if not isinstance(
        channels,
        Mapping,
    ):
        raise ValueError(
            "channels must be a mapping of channel names to arrays."
        )

    if len(channels) == 0:
        raise ValueError(
            "At least one vibration channel is required."
        )

    time = np.asarray(
        time,
        dtype=float,
    )

    rpm = np.asarray(
        rpm,
        dtype=float,
    )

    if len(time) != len(rpm):
        raise ValueError(
            "Time and RPM vectors must have the same length."
        )

    if len(time) < 3:
        raise ValueError(
            "The measurement contains too few samples."
        )

    if calibration_factor <= 0:
        raise ValueError(
            "calibration_factor must be greater than zero."
        )

    if max_order <= 0:
        raise ValueError(
            "max_order must be greater than zero."
        )

    if order_width <= 0:
        raise ValueError(
            "order_width must be greater than zero."
        )

    if rpm_step <= 0:
        raise ValueError(
            "rpm_step must be greater than zero."
        )

    normalized_order_definitions = {}

    for raw_order_value, raw_definition in (
        order_definitions.items()
    ):
        order_value = float(
            raw_order_value
        )

        if not isinstance(
            raw_definition,
            Mapping,
        ):
            raise ValueError(
                f"Order definition for {order_value:.2f} "
                "must be a mapping."
            )

        required_keys = {
            "label",
            "harmonic",
            "target_rpm",
            "target_amp",
        }

        missing_keys = (
            required_keys
            - set(raw_definition.keys())
        )

        if missing_keys:
            raise ValueError(
                f"Order definition for {order_value:.2f} "
                f"is missing keys: {sorted(missing_keys)}"
            )

        normalized_order_definitions[
            order_value
        ] = dict(
            raw_definition
        )

    if len(
        normalized_order_definitions
    ) == 0:
        raise ValueError(
            "No Transfer Case orders are defined."
        )

    highest_requested_order = max(
        normalized_order_definitions.keys()
    )

    if max_order < highest_requested_order:
        raise ValueError(
            f"Max Order must be at least "
            f"{highest_requested_order:.2f}. "
            f"Current value: {max_order:.2f}."
        )

    angular_nyquist_order = (
        samples_per_rev / 2.0
    )

    if highest_requested_order > angular_nyquist_order:
        raise ValueError(
            f"The highest requested order "
            f"{highest_requested_order:.2f} exceeds the "
            f"angular Nyquist limit "
            f"{angular_nyquist_order:.2f}. "
            "Increase Samples per Rev."
        )

    curves_by_order: Dict[
        float,
        dict
    ] = {}

    results_by_order: Dict[
        float,
        pd.DataFrame
    ] = {}

    raw_curves_by_order: Dict[
        float,
        pd.DataFrame
    ] = {}

    result_rows_by_order: Dict[
        float,
        list
    ] = {}

    for order_value in (
        normalized_order_definitions
    ):
        curves_by_order[
            order_value
        ] = {}

        result_rows_by_order[
            order_value
        ] = []

    # -------------------------------------------------------------------------
    # Channel analysis
    # -------------------------------------------------------------------------

    for channel_name, raw_signal in (
        channels.items()
    ):
        signal = np.asarray(
            raw_signal,
            dtype=float,
        )

        if len(signal) != len(time):
            raise ValueError(
                f"Channel {channel_name} length does not match "
                "the Time and RPM vectors."
            )

        theta_uniform, signal_uniform, rpm_uniform = angular_resample(
            time=time,
            rpm=rpm,
            signal=signal,
            samples_per_rev=samples_per_rev,
        )

        orders, block_rpms, spectrum = order_map(
            theta_u=theta_uniform,
            signal_u=signal_uniform,
            rpm_u=rpm_uniform,
            samples_per_rev=samples_per_rev,
            revs_per_block=revs_per_block,
            overlap=overlap,
            max_order=max_order,
        )

        if spectrum.ndim != 2:
            raise ValueError(
                f"Calculated spectrum for channel {channel_name} "
                "is not two-dimensional."
            )

        if spectrum.shape[0] != len(
            block_rpms
        ):
            raise ValueError(
                f"Spectrum row count for channel {channel_name} "
                "does not match the RPM block count."
            )

        if spectrum.shape[1] != len(
            orders
        ):
            raise ValueError(
                f"Spectrum column count for channel {channel_name} "
                "does not match the order axis."
            )

        for order_value, definition in (
            normalized_order_definitions.items()
        ):
            rpm_curve, amplitude_curve = extract_order_vs_rpm(
                orders=orders,
                rpms=block_rpms,
                spectrum=spectrum,
                target_order=order_value,
                width=order_width,
                rpm_step=rpm_step,
                smooth=True,
            )

            amplitude_curve = (
                np.asarray(
                    amplitude_curve,
                    dtype=float,
                )
                * calibration_factor
            )

            if len(rpm_curve) == 0:
                raise ValueError(
                    f"No RPM curve was generated for "
                    f"{order_value:.2f} order, "
                    f"channel {channel_name}."
                )

            if len(rpm_curve) != len(
                amplitude_curve
            ):
                raise ValueError(
                    f"RPM and amplitude length mismatch for "
                    f"{order_value:.2f} order, "
                    f"channel {channel_name}."
                )

            curves_by_order[
                order_value
            ][channel_name] = {
                "rpm": np.asarray(
                    rpm_curve,
                    dtype=float,
                ),
                "amp": amplitude_curve,
            }

            evaluation_result = evaluate_curve_against_target(
                rpm=rpm_curve,
                amplitude=amplitude_curve,
                target_rpm=definition[
                    "target_rpm"
                ],
                target_amp=definition[
                    "target_amp"
                ],
            )

            result_row = {
                "Order": order_value,
                "Order Label": definition[
                    "label"
                ],
                "Harmonic": definition[
                    "harmonic"
                ],
                "Channel": channel_name,
            }

            result_row.update(
                evaluation_result
            )

            result_rows_by_order[
                order_value
            ].append(
                result_row
            )

    # -------------------------------------------------------------------------
    # Build result tables and combined raw curves
    # -------------------------------------------------------------------------

    for order_value, definition in (
        normalized_order_definitions.items()
    ):
        order_result_rows = (
            result_rows_by_order[
                order_value
            ]
        )

        if len(order_result_rows) == 0:
            raise ValueError(
                f"No result rows were created for "
                f"{order_value:.2f} order."
            )

        result_dataframe = pd.DataFrame(
            order_result_rows
        )

        results_by_order[
            order_value
        ] = result_dataframe

        channel_curves = (
            curves_by_order[
                order_value
            ]
        )

        if len(channel_curves) == 0:
            raise ValueError(
                f"No channel curves were generated for "
                f"{order_value:.2f} order."
            )

        first_channel_name = next(
            iter(
                channel_curves
            )
        )

        base_rpm = np.asarray(
            channel_curves[
                first_channel_name
            ]["rpm"],
            dtype=float,
        )

        if len(base_rpm) == 0:
            raise ValueError(
                f"The base RPM curve is empty for "
                f"{order_value:.2f} order."
            )

        curve_dataframe = pd.DataFrame(
            {
                "RPM": base_rpm
            }
        )

        for channel_name, channel_curve in (
            channel_curves.items()
        ):
            channel_rpm = np.asarray(
                channel_curve["rpm"],
                dtype=float,
            )

            channel_amplitude = np.asarray(
                channel_curve["amp"],
                dtype=float,
            )

            if len(channel_rpm) != len(
                channel_amplitude
            ):
                raise ValueError(
                    f"RPM and amplitude length mismatch for "
                    f"{order_value:.2f} order, "
                    f"channel {channel_name}."
                )

            curve_dataframe[
                channel_name
            ] = np.interp(
                base_rpm,
                channel_rpm,
                channel_amplitude,
            )

        target_rpm = definition[
            "target_rpm"
        ]

        target_amp = definition[
            "target_amp"
        ]

        if (
            target_rpm is not None
            and target_amp is not None
        ):
            curve_dataframe[
                "Target"
            ] = np.interp(
                base_rpm,
                np.asarray(
                    target_rpm,
                    dtype=float,
                ),
                np.asarray(
                    target_amp,
                    dtype=float,
                ),
            )

        raw_curves_by_order[
            order_value
        ] = curve_dataframe

    return (
        curves_by_order,
        results_by_order,
        raw_curves_by_order,
    )  
# =============================================================================
# ORDER MAP PLOTTING AND MODULE VALIDATION
# =============================================================================

def plot_order_map(
    orders: np.ndarray,
    rpms: np.ndarray,
    spectrum: np.ndarray,
    channel_name: str = "Channel",
    db_reference: float = 1.0,
    calibration_factor: float = 1.0
):
    """
    Create the Transfer Case order map / waterfall plot.

    Parameters
    ----------
    orders:
        Calculated order axis.

    rpms:
        Mean RPM value for every FFT block.

    spectrum:
        Two-dimensional order spectrum.

        Shape:

            number_of_blocks × number_of_orders

    channel_name:
        Name of the plotted vibration channel.

    db_reference:
        Reference amplitude used for the dB calculation.

    calibration_factor:
        Final amplitude multiplier.

    Returns
    -------
    matplotlib.figure.Figure
        Order map figure.
    """
    orders = np.asarray(
        orders,
        dtype=float,
    )

    rpms = np.asarray(
        rpms,
        dtype=float,
    )

    spectrum = np.asarray(
        spectrum,
        dtype=float,
    )

    if len(orders) == 0:
        raise ValueError(
            "Order axis is empty."
        )

    if len(rpms) == 0:
        raise ValueError(
            "RPM vector is empty."
        )

    if spectrum.ndim != 2:
        raise ValueError(
            "Order spectrum must be a two-dimensional array. "
            f"Received shape: {spectrum.shape}."
        )

    if spectrum.shape[0] != len(rpms):
        raise ValueError(
            "Spectrum row count does not match the RPM vector length."
        )

    if spectrum.shape[1] != len(orders):
        raise ValueError(
            "Spectrum column count does not match the order-axis length."
        )

    if db_reference <= 0:
        raise ValueError(
            "db_reference must be greater than zero."
        )

    if calibration_factor <= 0:
        raise ValueError(
            "calibration_factor must be greater than zero."
        )

    valid_rpm_mask = np.isfinite(
        rpms
    )

    if not np.any(
        valid_rpm_mask
    ):
        raise ValueError(
            "RPM vector does not contain valid values."
        )

    rpms = rpms[
        valid_rpm_mask
    ]

    spectrum = spectrum[
        valid_rpm_mask,
        :
    ]

    if not np.all(
        np.isfinite(spectrum)
    ):
        spectrum = np.where(
            np.isfinite(spectrum),
            spectrum,
            0.0,
        )

    sort_indices = np.argsort(
        rpms,
        kind="stable",
    )

    sorted_rpms = rpms[
        sort_indices
    ]

    sorted_spectrum = spectrum[
        sort_indices,
        :
    ]

    scaled_spectrum = (
        sorted_spectrum
        * calibration_factor
    )

    spectrum_db = (
        20.0
        * np.log10(
            np.maximum(
                scaled_spectrum,
                1e-12,
            )
            / db_reference
        )
    )

    figure, axis = plt.subplots(
        figsize=(12, 7)
    )

    image = axis.imshow(
        spectrum_db,
        aspect="auto",
        origin="lower",
        extent=[
            float(orders[0]),
            float(orders[-1]),
            float(sorted_rpms[0]),
            float(sorted_rpms[-1]),
        ],
        interpolation="nearest",
        cmap="jet",
    )

    colorbar = figure.colorbar(
        image,
        ax=axis,
    )

    colorbar.set_label(
        "Amplitude [dB re 1 m/s²]"
    )

    axis.set_xlabel(
        "Order"
    )

    axis.set_ylabel(
        "RPM"
    )

    axis.set_title(
        f"Transfer Case Order Map / Waterfall - {channel_name}"
    )

    axis.grid(
        False
    )

    figure.tight_layout()

    return figure


def validate_transfer_case_module() -> dict:
    """
    Validate the internal Transfer Case module configuration.

    This function does not run an actual vehicle analysis. It checks:

        - Required order definitions
        - Target array lengths
        - Highest requested order
        - Recommended FFT settings
        - Required public functions

    Returns
    -------
    dict
        Validation summary.
    """
    required_orders = {
        63.0,
        85.05,
        126.0,
        170.10,
    }

    configured_orders = set(
        float(order_value)
        for order_value in TRANSFER_CASE_ORDERS.keys()
    )

    missing_orders = (
        required_orders
        - configured_orders
    )

    if missing_orders:
        raise ValueError(
            "Transfer Case order definitions are incomplete. "
            f"Missing orders: {sorted(missing_orders)}"
        )

    for order_value, definition in (
        TRANSFER_CASE_ORDERS.items()
    ):
        required_keys = {
            "label",
            "harmonic",
            "target_rpm",
            "target_amp",
        }

        missing_keys = (
            required_keys
            - set(definition.keys())
        )

        if missing_keys:
            raise ValueError(
                f"Order {order_value:.2f} definition is missing: "
                f"{sorted(missing_keys)}"
            )

        target_rpm = definition[
            "target_rpm"
        ]

        target_amp = definition[
            "target_amp"
        ]

        if (
            target_rpm is None
            and target_amp is None
        ):
            continue

        if (
            target_rpm is None
            or target_amp is None
        ):
            raise ValueError(
                f"Order {order_value:.2f} has an incomplete target definition."
            )

        target_rpm = np.asarray(
            target_rpm,
            dtype=float,
        )

        target_amp = np.asarray(
            target_amp,
            dtype=float,
        )

        if len(target_rpm) != len(
            target_amp
        ):
            raise ValueError(
                f"Order {order_value:.2f} target RPM and amplitude "
                "lengths do not match."
            )

        if len(target_rpm) < 2:
            raise ValueError(
                f"Order {order_value:.2f} target curve "
                "contains fewer than two points."
            )

        if not np.all(
            np.diff(target_rpm) > 0
        ):
            raise ValueError(
                f"Order {order_value:.2f} target RPM values "
                "must be strictly increasing."
            )

        if not np.all(
            np.isfinite(target_amp)
        ):
            raise ValueError(
                f"Order {order_value:.2f} target amplitude "
                "contains invalid values."
            )

    required_public_functions = {
        "read_xlsx_numeric": read_xlsx_numeric,
        "angular_resample": angular_resample,
        "order_map": order_map,
        "extract_order_vs_rpm": extract_order_vs_rpm,
        "evaluate_curve_against_target": evaluate_curve_against_target,
        "calculate_overall_status": calculate_overall_status,
        "analyze_transfer_case_orders": analyze_transfer_case_orders,
        "plot_order_map": plot_order_map,
    }

    for function_name, function_object in (
        required_public_functions.items()
    ):
        if not callable(
            function_object
        ):
            raise ValueError(
                f"{function_name} is not callable."
            )

    highest_order = max(
        configured_orders
    )

    recommended_samples_per_rev = 512
    recommended_revs_per_block = 20
    recommended_max_order = 200.0

    angular_nyquist_order = (
        recommended_samples_per_rev
        / 2.0
    )

    if highest_order > angular_nyquist_order:
        raise ValueError(
            "The highest Transfer Case order exceeds "
            "the recommended angular Nyquist limit."
        )

    order_resolution = (
        1.0
        / recommended_revs_per_block
    )

    return {
        "Module Status": "READY",
        "Configured Orders": sorted(
            configured_orders
        ),
        "Highest Order": highest_order,
        "Recommended Samples per Rev": recommended_samples_per_rev,
        "Recommended Revs per Block": recommended_revs_per_block,
        "Recommended Max Order": recommended_max_order,
        "Angular Nyquist Order": angular_nyquist_order,
        "Order Resolution": order_resolution,
    }


# =============================================================================
# PUBLIC MODULE INTERFACE
# =============================================================================

__all__ = [
    "G_TO_MS2",
    "TRANSFER_CASE_TARGET_RPM",
    "TRANSFER_CASE_ORDERS",
    "read_xlsx_numeric",
    "repair_time_vector_without_dropping_rows",
    "angular_resample",
    "order_map",
    "smooth_curve",
    "average_duplicate_rpm_values",
    "resample_to_rpm_step",
    "interpolate_fractional_order",
    "extract_order_vs_rpm",
    "integrate_positive_area",
    "evaluate_curve_against_target",
    "calculate_overall_status",
    "analyze_transfer_case_orders",
    "plot_order_map",
    "validate_transfer_case_module",
]


# =============================================================================
# OPTIONAL STANDALONE MODULE CHECK
# =============================================================================

if __name__ == "__main__":
    validation_summary = validate_transfer_case_module()

    print(
        "Transfer Case analysis module validation:"
    )

    for key, value in validation_summary.items():
        print(
            f"{key}: {value}"
        )
