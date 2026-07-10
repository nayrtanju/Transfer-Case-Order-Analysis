import zipfile
import xml.etree.ElementTree as ET
import re

import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import savgol_filter


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def load_shared_strings(z):
    strings = []

    if "xl/sharedStrings.xml" not in z.namelist():
        return strings

    with z.open("xl/sharedStrings.xml") as stream:
        for _, elem in ET.iterparse(stream, events=("end",)):
            if elem.tag == NS + "si":
                texts = []

                for text_element in elem.iter(NS + "t"):
                    texts.append(text_element.text or "")

                strings.append("".join(texts))
                elem.clear()

    return strings


def col_index(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref)

    if match is None:
        raise ValueError(
            f"Invalid Excel cell reference: {cell_ref}"
        )

    letters = match.group(1)

    index = 0

    for character in letters:
        index = index * 26 + ord(character) - 64

    return index - 1


def read_xlsx_numeric(path):
    with zipfile.ZipFile(path) as z:
        shared = load_shared_strings(z)

        sheet_path = "xl/worksheets/sheet1.xml"

        if sheet_path not in z.namelist():
            raise ValueError(
                "The Excel file does not contain sheet1.xml."
            )

        nrows = 0

        with z.open(sheet_path) as stream:
            for _, elem in ET.iterparse(
                stream,
                events=("start",)
            ):
                if elem.tag == NS + "dimension":
                    reference = elem.attrib.get("ref", "")

                    match = re.search(
                        r":([A-Z]+)(\d+)",
                        reference
                    )

                    if match:
                        nrows = max(
                            0,
                            int(match.group(2)) - 1
                        )

                    break

        if nrows <= 0:
            raise ValueError(
                "The Excel worksheet does not contain numeric data rows."
            )

        data = np.empty(
            (nrows, 5),
            dtype=np.float64
        )

        data[:] = np.nan

        headers = [None] * 5
        row_index = -1

        with z.open(sheet_path) as stream:
            for _, row in ET.iterparse(
                stream,
                events=("end",)
            ):
                if row.tag != NS + "row":
                    continue

                row_number = int(
                    row.attrib.get("r", "0")
                )

                values = [np.nan] * 5

                for cell in row.findall(NS + "c"):
                    reference = cell.attrib.get("r", "")

                    try:
                        column_index = col_index(reference)
                    except ValueError:
                        continue

                    if column_index >= 5:
                        continue

                    cell_type = cell.attrib.get("t")
                    value_element = cell.find(NS + "v")

                    if value_element is None:
                        continue

                    text_value = value_element.text

                    if text_value is None:
                        continue

                    if row_number == 1:
                        if cell_type == "s":
                            shared_index = int(text_value)

                            if 0 <= shared_index < len(shared):
                                headers[column_index] = shared[shared_index]
                            else:
                                headers[column_index] = text_value
                        else:
                            headers[column_index] = text_value
                    else:
                        try:
                            values[column_index] = float(text_value)
                        except (TypeError, ValueError):
                            values[column_index] = np.nan

                if row_number > 1:
                    row_index += 1

                    if row_index >= data.shape[0]:
                        data = np.vstack(
                            [
                                data,
                                np.full(
                                    (10000, 5),
                                    np.nan,
                                    dtype=np.float64
                                )
                            ]
                        )

                    data[row_index, :] = values

                row.clear()

        if row_index < 0:
            raise ValueError(
                "No numeric data rows were found in the Excel worksheet."
            )

        return headers, data[:row_index + 1]


def angular_resample(
    time,
    rpm,
    signal,
    samples_per_rev=512
):
    time = np.asarray(
        time,
        dtype=float
    )

    rpm = np.asarray(
        rpm,
        dtype=float
    )

    signal = np.asarray(
        signal,
        dtype=float
    )

    if not (
        len(time) == len(rpm) == len(signal)
    ):
        raise ValueError(
            "Time, RPM and signal arrays must have the same length."
        )

    if samples_per_rev <= 0:
        raise ValueError(
            "samples_per_rev must be greater than zero."
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
            "Not enough valid samples are available for angular resampling."
        )

    # Stable sorting preserves the original order of equal-time samples.
    sort_indices = np.argsort(
        time,
        kind="stable"
    )

    time = time[sort_indices]
    rpm = rpm[sort_indices]
    signal = signal[sort_indices]

    dt = np.diff(
        time,
        prepend=time[0]
    )

    positive_dt = dt[
        np.isfinite(dt) & (dt > 0)
    ]

    if len(positive_dt) == 0:
        raise ValueError(
            "Time vector does not contain positive time steps. "
            "Please check the time column precision."
        )

    median_dt = float(
        np.median(positive_dt)
    )

    if not np.isfinite(median_dt) or median_dt <= 0:
        raise ValueError(
            "A valid median time step could not be calculated."
        )

    # Pico CSV files may contain repeated time values because of
    # decimal precision. No data rows are deleted. Zero or negative
    # time increments are replaced with the median positive time step.
    dt = np.where(
        np.isfinite(dt) & (dt > 0),
        dt,
        median_dt
    )

    dt[0] = median_dt

    angular_velocity = (
        2.0
        * np.pi
        * rpm
        / 60.0
    )

    theta = np.cumsum(
        angular_velocity * dt
    )

    if len(theta) < 2:
        raise ValueError(
            "Angular position could not be calculated."
        )

    increasing_mask = np.r_[
        True,
        np.diff(theta) > 0
    ]

    theta = theta[increasing_mask]
    signal = signal[increasing_mask]
    rpm = rpm[increasing_mask]

    if len(theta) < 2:
        raise ValueError(
            "Angular position does not contain enough increasing samples."
        )

    angular_step = (
        2.0
        * np.pi
        / samples_per_rev
    )

    if theta[-1] <= theta[0]:
        raise ValueError(
            "The available angular range is insufficient for resampling."
        )

    theta_uniform = np.arange(
        theta[0],
        theta[-1],
        angular_step
    )

    if len(theta_uniform) < 2:
        raise ValueError(
            "Angular resampling produced too few samples."
        )

    signal_uniform = np.interp(
        theta_uniform,
        theta,
        signal
    )

    rpm_uniform = np.interp(
        theta_uniform,
        theta,
        rpm
    )

    return (
        theta_uniform,
        signal_uniform,
        rpm_uniform
    )


def order_map(
    theta_u,
    x_u,
    rpm_u,
    samples_per_rev=512,
    revs_per_block=8,
    overlap=0.75,
    max_order=30
):
    theta_u = np.asarray(
        theta_u,
        dtype=float
    )

    x_u = np.asarray(
        x_u,
        dtype=float
    )

    rpm_u = np.asarray(
        rpm_u,
        dtype=float
    )

    if not (
        len(theta_u)
        == len(x_u)
        == len(rpm_u)
    ):
        raise ValueError(
            "theta_u, x_u and rpm_u must have the same length."
        )

    if samples_per_rev <= 0:
        raise ValueError(
            "samples_per_rev must be greater than zero."
        )

    if revs_per_block <= 0:
        raise ValueError(
            "revs_per_block must be greater than zero."
        )

    if not 0 <= overlap < 1:
        raise ValueError(
            "overlap must be greater than or equal to 0 "
            "and smaller than 1."
        )

    if max_order <= 0:
        raise ValueError(
            "max_order must be greater than zero."
        )

    if len(x_u) == 0:
        raise ValueError(
            "Angular resampling produced no samples."
        )

    available_revolutions = (
        len(x_u)
        / float(samples_per_rev)
    )

    available_complete_revolutions = int(
        np.floor(available_revolutions)
    )

    effective_revs_per_block = min(
        int(revs_per_block),
        available_complete_revolutions
    )

    if effective_revs_per_block < 2:
        raise ValueError(
            "Insufficient measurement duration for order analysis. "
            f"Available revolutions: {available_revolutions:.2f}. "
            "At least 2 complete revolutions are required."
        )

    samples_per_block = int(
        samples_per_rev
        * effective_revs_per_block
    )

    hop_size = max(
        1,
        int(
            round(
                samples_per_block
                * (1.0 - overlap)
            )
        )
    )

    if len(x_u) < samples_per_block:
        raise ValueError(
            "Angular-resampled signal is shorter than one FFT block. "
            f"Signal samples: {len(x_u)}, "
            f"required samples: {samples_per_block}."
        )

    window = np.hanning(
        samples_per_block
    )

    window_sum = float(
        np.sum(window)
    )

    if not np.isfinite(window_sum) or window_sum <= 0:
        raise ValueError(
            "Invalid FFT window scaling."
        )

    complete_order_axis = np.fft.rfftfreq(
        samples_per_block,
        d=1.0 / samples_per_rev
    )

    order_mask = (
        complete_order_axis <= max_order
    )

    orders = complete_order_axis[
        order_mask
    ]

    if len(orders) == 0:
        raise ValueError(
            "No order bins are available for the selected max_order."
        )

    spectra = []
    block_rpms = []

    last_start = (
        len(x_u)
        - samples_per_block
    )

    for start in range(
        0,
        last_start + 1,
        hop_size
    ):
        stop = start + samples_per_block

        signal_block = x_u[start:stop]
        rpm_block = rpm_u[start:stop]

        if len(signal_block) != samples_per_block:
            continue

        if len(rpm_block) != samples_per_block:
            continue

        if not np.all(
            np.isfinite(signal_block)
        ):
            continue

        if not np.all(
            np.isfinite(rpm_block)
        ):
            continue

        signal_block = (
            signal_block
            - np.mean(signal_block)
        )

        fft_values = np.fft.rfft(
            signal_block * window
        )

        # Artemis-compatible RMS amplitude scaling.
        amplitude = (
            np.sqrt(2.0)
            * np.abs(fft_values)
            / window_sum
        )

        selected_amplitude = amplitude[
            order_mask
        ]

        spectra.append(
            selected_amplitude
        )

        block_rpms.append(
            float(
                np.mean(rpm_block)
            )
        )

    if len(spectra) == 0:
        raise ValueError(
            "No valid FFT blocks could be generated. "
            f"Angular samples: {len(x_u)}, "
            f"block size: {samples_per_block}, "
            f"available revolutions: {available_revolutions:.2f}."
        )

    # vstack guarantees that the spectrum is always two-dimensional.
    spectrum_array = np.vstack(
        spectra
    ).astype(
        np.float64,
        copy=False
    )

    rpm_array = np.asarray(
        block_rpms,
        dtype=np.float64
    )

    if spectrum_array.ndim != 2:
        raise ValueError(
            "Calculated order spectrum is not two-dimensional."
        )

    if spectrum_array.shape[0] != len(rpm_array):
        raise ValueError(
            "The number of FFT blocks does not match the RPM vector."
        )

    return (
        orders,
        rpm_array,
        spectrum_array
    )


def smooth_curve(
    y,
    window_length=9,
    polyorder=2
):
    y = np.asarray(
        y,
        dtype=float
    )

    if len(y) < 5:
        return y

    valid_mask = np.isfinite(y)

    if not np.all(valid_mask):
        valid_indices = np.flatnonzero(
            valid_mask
        )

        if len(valid_indices) < 2:
            return y

        y = np.interp(
            np.arange(len(y)),
            valid_indices,
            y[valid_indices]
        )

    window_length = int(
        window_length
    )

    polyorder = int(
        polyorder
    )

    if window_length % 2 == 0:
        window_length += 1

    if window_length >= len(y):
        window_length = len(y) - 1

    if window_length % 2 == 0:
        window_length -= 1

    minimum_window = polyorder + 2

    if minimum_window % 2 == 0:
        minimum_window += 1

    if window_length < minimum_window:
        return y

    return savgol_filter(
        y,
        window_length=window_length,
        polyorder=polyorder
    )


def resample_to_rpm_step(
    rpm,
    amplitude,
    rpm_step=10
):
    rpm = np.asarray(
        rpm,
        dtype=float
    )

    amplitude = np.asarray(
        amplitude,
        dtype=float
    )

    if rpm_step <= 0:
        raise ValueError(
            "rpm_step must be greater than zero."
        )

    valid_mask = (
        np.isfinite(rpm)
        & np.isfinite(amplitude)
    )

    rpm = rpm[valid_mask]
    amplitude = amplitude[valid_mask]

    if len(rpm) < 2:
        return rpm, amplitude

    sort_indices = np.argsort(
        rpm,
        kind="stable"
    )

    rpm = rpm[sort_indices]
    amplitude = amplitude[sort_indices]

    # Average amplitude values when multiple FFT blocks have
    # exactly the same RPM value.
    unique_rpm, inverse_indices = np.unique(
        rpm,
        return_inverse=True
    )

    if len(unique_rpm) != len(rpm):
        amplitude_sum = np.zeros(
            len(unique_rpm),
            dtype=float
        )

        amplitude_count = np.zeros(
            len(unique_rpm),
            dtype=float
        )

        np.add.at(
            amplitude_sum,
            inverse_indices,
            amplitude
        )

        np.add.at(
            amplitude_count,
            inverse_indices,
            1.0
        )

        rpm = unique_rpm

        amplitude = (
            amplitude_sum
            / np.maximum(
                amplitude_count,
                1.0
            )
        )

    if len(rpm) < 2:
        return rpm, amplitude

    rpm_minimum = (
        np.ceil(rpm[0] / rpm_step)
        * rpm_step
    )

    rpm_maximum = (
        np.floor(rpm[-1] / rpm_step)
        * rpm_step
    )

    if rpm_maximum <= rpm_minimum:
        return rpm, amplitude

    rpm_grid = np.arange(
        rpm_minimum,
        rpm_maximum + rpm_step,
        rpm_step
    )

    amplitude_grid = np.interp(
        rpm_grid,
        rpm,
        amplitude
    )

    return rpm_grid, amplitude_grid


def plot_order_map(
    orders,
    rpms,
    spec,
    channel_name="Channel",
    db_reference=1.0
):
    orders = np.asarray(
        orders,
        dtype=float
    )

    rpms = np.asarray(
        rpms,
        dtype=float
    )

    spec = np.asarray(
        spec,
        dtype=float
    )

    if spec.ndim != 2:
        raise ValueError(
            "Order spectrum must be a two-dimensional array. "
            f"Received shape: {spec.shape}."
        )

    if spec.shape[0] == 0 or spec.shape[1] == 0:
        raise ValueError(
            "Order spectrum is empty."
        )

    if spec.shape[0] != len(rpms):
        raise ValueError(
            "RPM vector length does not match the number of spectrum blocks."
        )

    if spec.shape[1] != len(orders):
        raise ValueError(
            "Order axis length does not match the spectrum columns."
        )

    if db_reference <= 0:
        raise ValueError(
            "db_reference must be greater than zero."
        )

    sort_indices = np.argsort(
        rpms
    )

    sorted_rpms = rpms[
        sort_indices
    ]

    sorted_spectrum = spec[
        sort_indices,
        :
    ]

    spectrum_db = (
        20.0
        * np.log10(
            np.maximum(
                sorted_spectrum,
                1e-12
            )
            / db_reference
        )
    )

    figure, axis = plt.subplots(
        figsize=(11, 7)
    )

    image = axis.imshow(
        spectrum_db,
        aspect="auto",
        origin="lower",
        extent=[
            orders[0],
            orders[-1],
            sorted_rpms[0],
            sorted_rpms[-1]
        ],
        interpolation="nearest",
        cmap="jet"
    )

    figure.colorbar(
        image,
        ax=axis,
        label="Amplitude [dB re 1 m/s²]"
    )

    axis.set_xlabel("Order")
    axis.set_ylabel("RPM")
    axis.set_title(
        f"Order Map - {channel_name}"
    )

    return figure


def extract_order_vs_rpm(
    orders,
    rpms,
    spec,
    target_order=10.0,
    width=0.15,
    rpm_step=10,
    smooth=True
):
    orders = np.asarray(
        orders,
        dtype=float
    )

    rpms = np.asarray(
        rpms,
        dtype=float
    )

    spec = np.asarray(
        spec,
        dtype=float
    )

    if spec.ndim != 2:
        raise ValueError(
            "Order spectrum must be a two-dimensional array. "
            f"Received shape: {spec.shape}."
        )

    if spec.shape[0] == 0 or spec.shape[1] == 0:
        raise ValueError(
            "Order spectrum is empty."
        )

    if len(orders) == 0:
        raise ValueError(
            "Order axis is empty."
        )

    if len(rpms) == 0:
        raise ValueError(
            "RPM vector is empty."
        )

    if spec.shape[0] != len(rpms):
        raise ValueError(
            "RPM vector length does not match the number of spectrum blocks."
        )

    if spec.shape[1] != len(orders):
        raise ValueError(
            "Order axis length does not match the spectrum columns."
        )

    if target_order < orders[0]:
        raise ValueError(
            f"Target order {target_order:.2f} is below the calculated "
            f"minimum order {orders[0]:.2f}."
        )

    if target_order > orders[-1]:
        raise ValueError(
            f"Target order {target_order:.2f} is above the calculated "
            f"maximum order {orders[-1]:.2f}. Increase Max order."
        )

    if width <= 0:
        raise ValueError(
            "Order width must be greater than zero."
        )

    band_mask = (
        (
            orders
            >= target_order - width / 2.0
        )
        &
        (
            orders
            <= target_order + width / 2.0
        )
    )

    if not np.any(band_mask):
        nearest_order_index = int(
            np.argmin(
                np.abs(
                    orders - target_order
                )
            )
        )

        amplitude = spec[
            :,
            nearest_order_index
        ]
    else:
        # RSS integration across the selected order band.
        amplitude = np.sqrt(
            np.sum(
                spec[:, band_mask] ** 2,
                axis=1
            )
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
        kind="stable"
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
            polyorder=2
        )

    rpm_resampled, amplitude_resampled = resample_to_rpm_step(
        rpm_sorted,
        amplitude_sorted,
        rpm_step=rpm_step
    )

    if len(rpm_resampled) == 0:
        raise ValueError(
            f"RPM resampling produced no data for order "
            f"{target_order:.2f}."
        )

    return (
        rpm_resampled,
        amplitude_resampled
    )
