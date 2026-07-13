import os
import re
import tempfile
import traceback
from io import BytesIO
from typing import Dict, Mapping, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill


# =============================================================================
# STREAMLIT PAGE CONFIGURATION
# =============================================================================
st.set_page_config(
    page_title="NVH Analysis Suite",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =============================================================================
# CORPORATE UI THEME
# =============================================================================

st.markdown(
    """
    <style>
    /* ---------------------------------------------------------
       GLOBAL PAGE
    --------------------------------------------------------- */

    .stApp {
        background-color: #f4f6f8;
    }

    .block-container {
        max-width: 1450px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }

    html,
    body,
    [class*="css"] {
        font-family:
            Inter,
            "Segoe UI",
            Arial,
            sans-serif;
    }


    /* ---------------------------------------------------------
       HEADER
    --------------------------------------------------------- */

    .corporate-header {
        background:
            linear-gradient(
                135deg,
                #0b1f33 0%,
                #123a63 55%,
                #1768a6 100%
            );

        border-radius: 16px;
        padding: 28px 34px;
        margin-bottom: 24px;

        box-shadow:
            0 10px 30px rgba(11, 31, 51, 0.18);

        color: white;
    }

    .corporate-header-title {
        font-size: 2.05rem;
        font-weight: 700;
        line-height: 1.15;
        margin: 0;
        letter-spacing: -0.02em;
    }

    .corporate-header-subtitle {
        font-size: 1rem;
        margin-top: 8px;
        color: rgba(255, 255, 255, 0.82);
    }

    .corporate-header-badge {
        display: inline-block;
        margin-top: 14px;
        padding: 6px 12px;

        border-radius: 999px;

        background-color:
            rgba(255, 255, 255, 0.14);

        border:
            1px solid
            rgba(255, 255, 255, 0.20);

        font-size: 0.82rem;
        font-weight: 600;
        letter-spacing: 0.02em;
    }


    /* ---------------------------------------------------------
       SECTION TITLES
    --------------------------------------------------------- */

    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #18324a;

        margin-top: 1.6rem;
        margin-bottom: 0.7rem;

        padding-left: 12px;
        border-left: 4px solid #1768a6;
    }


    /* ---------------------------------------------------------
       CARD DESIGN
    --------------------------------------------------------- */

    .engineering-card {
        background-color: #ffffff;

        border:
            1px solid
            #dfe5ea;

        border-radius: 14px;

        padding: 18px 20px;

        margin-bottom: 16px;

        box-shadow:
            0 4px 16px
            rgba(17, 45, 72, 0.06);
    }


    /* ---------------------------------------------------------
       INPUTS
    --------------------------------------------------------- */

    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div {
        border-radius: 10px !important;
        border-color: #cdd7df !important;
        background-color: white !important;
    }

    div[data-baseweb="input"] > div:focus-within,
    div[data-baseweb="select"] > div:focus-within {
        border-color: #1768a6 !important;

        box-shadow:
            0 0 0 2px
            rgba(23, 104, 166, 0.12) !important;
    }


    /* ---------------------------------------------------------
       FILE UPLOADER
    --------------------------------------------------------- */

    section[data-testid="stFileUploaderDropzone"] {
        background-color: #ffffff;

        border:
            1.5px dashed
            #9eb5c7;

        border-radius: 14px;

        padding: 18px;

        transition:
            border-color 0.2s ease,
            background-color 0.2s ease;
    }

    section[data-testid="stFileUploaderDropzone"]:hover {
        border-color: #1768a6;
        background-color: #f7fbff;
    }


    /* ---------------------------------------------------------
       BUTTONS
    --------------------------------------------------------- */

    .stButton > button {
        min-height: 48px;

        border-radius: 10px;

        border: none;

        font-weight: 700;
        letter-spacing: 0.01em;

        background:
            linear-gradient(
                135deg,
                #155d95,
                #1d78bd
            );

        color: white;

        box-shadow:
            0 6px 16px
            rgba(23, 104, 166, 0.24);

        transition:
            transform 0.15s ease,
            box-shadow 0.15s ease;
    }

    .stButton > button:hover {
        transform: translateY(-1px);

        box-shadow:
            0 9px 22px
            rgba(23, 104, 166, 0.28);

        color: white;
    }


    /* ---------------------------------------------------------
       DOWNLOAD BUTTON
    --------------------------------------------------------- */

    .stDownloadButton > button {
        min-height: 46px;

        border-radius: 10px;

        border:
            1px solid
            #1768a6;

        background-color: white;

        color: #1768a6;

        font-weight: 700;
    }

    .stDownloadButton > button:hover {
        background-color: #1768a6;
        color: white;
    }


    /* ---------------------------------------------------------
       METRICS
    --------------------------------------------------------- */

    div[data-testid="metric-container"] {
        background-color: white;

        border:
            1px solid
            #dfe5ea;

        border-radius: 12px;

        padding: 14px 16px;

        box-shadow:
            0 3px 12px
            rgba(17, 45, 72, 0.05);
    }

    div[data-testid="metric-container"]
    label {
        color: #5c6f7f !important;
        font-weight: 600 !important;
    }

    div[data-testid="metric-container"]
    [data-testid="stMetricValue"] {
        color: #17324d;
        font-weight: 700;
    }


    /* ---------------------------------------------------------
       TABS
    --------------------------------------------------------- */

    button[data-baseweb="tab"] {
        font-weight: 600;
        padding-left: 18px;
        padding-right: 18px;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #1768a6 !important;
    }


    /* ---------------------------------------------------------
       DATAFRAMES
    --------------------------------------------------------- */

    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;

        border:
            1px solid
            #dfe5ea;
    }


    /* ---------------------------------------------------------
       STATUS MESSAGES
    --------------------------------------------------------- */

    div[data-testid="stAlert"] {
        border-radius: 12px;
    }


    /* ---------------------------------------------------------
       HIDE DEFAULT STREAMLIT CHROME
    --------------------------------------------------------- */

    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header[data-testid="stHeader"] {
        background-color: transparent;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
# =============================================================================
# CORPORATE HEADER
# =============================================================================

st.markdown(
    """
<div class="corporate-header">
    <div class="corporate-header-title">NVH Analysis Suite</div>
    <div class="corporate-header-subtitle">
        Axle Whine and Transfer Case Gear Mesh Analysis Platform
    </div>
    <div class="corporate-header-badge">
        Engineering Validation Environment
    </div>
</div>
    """,
    unsafe_allow_html=True,
)
# =============================================================================
# ANALYSIS TYPE CONSTANTS
# =============================================================================

ANALYSIS_AXLE = "Axle Whine Order Analysis"

ANALYSIS_TRANSFER_CASE = (
    "Transfer Case Gear Mesh Analysis"
)

CHANNEL_NAMES = [
    "ChA",
    "ChB",
    "ChC",
]


# =============================================================================
# GENERAL APPLICATION CONSTANTS
# =============================================================================

MAX_FILE_SIZE_MB = 500

MAX_ROWS = 3_000_000

G_TO_MS2 = 9.80665


# =============================================================================
# AXLE WHINE ANALYSIS ENGINE IMPORT
# =============================================================================

try:
    from order_analysis import (
        read_xlsx_numeric as axle_read_xlsx_numeric,
        angular_resample as axle_angular_resample,
        order_map as axle_order_map,
        extract_order_vs_rpm as axle_extract_order_vs_rpm,
    )

except Exception:
    st.error(
        "order_analysis.py could not be loaded."
    )

    st.code(
        traceback.format_exc()
    )

    st.stop()


# =============================================================================
# TRANSFER CASE ANALYSIS ENGINE IMPORT
# =============================================================================

try:
    from transfer_case_analysis import (
        read_xlsx_numeric as tc_read_xlsx_numeric,
        angular_resample as tc_angular_resample,
        order_map as tc_order_map,
        analyze_transfer_case_orders,
        TRANSFER_CASE_ORDERS,
        validate_transfer_case_module,
    )

except Exception:
    st.error(
        "transfer_case_analysis.py could not be loaded."
    )

    st.code(
        traceback.format_exc()
    )

    st.stop()


# =============================================================================
# TRANSFER CASE MODULE VALIDATION
# =============================================================================

try:
    transfer_case_validation = (
        validate_transfer_case_module()
    )

except Exception:
    st.error(
        "Transfer Case analysis module validation failed."
    )

    st.code(
        traceback.format_exc()
    )

    st.stop()
    # =============================================================================
# AXLE WHINE TARGET DEFINITIONS
# =============================================================================

AXLE_TARGETS = {
    "Diesel": {
        "Front Axle": {
            "rpm": np.array(
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
            ),
            "amp": np.array(
                [
                    2.5,
                    2.5,
                    2.5,
                    7.5,
                    7.5,
                    7.5,
                    7.5,
                    7.5,
                ],
                dtype=float,
            ),
        },

        "Rear Axle": {
            "rpm": np.array(
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
            ),
            "amp": np.array(
                [
                    2.5,
                    2.5,
                    2.5,
                    7.5,
                    7.5,
                    7.5,
                    7.5,
                    7.5,
                ],
                dtype=float,
            ),
        },
    },

    "Gasoline": {
        "Front Axle": {
            "rpm": np.array(
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
            ),
            "amp": np.array(
                [
                    2.5,
                    2.5,
                    2.5,
                    6.25,
                    10.0,
                    10.0,
                    10.0,
                    10.0,
                ],
                dtype=float,
            ),
        },

        "Rear Axle": {
            "rpm": np.array(
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
            ),
            "amp": np.array(
                [
                    5.0,
                    5.0,
                    5.0,
                    10.0,
                    12.5,
                    12.5,
                    12.5,
                    12.5,
                ],
                dtype=float,
            ),
        },
    },
}


# =============================================================================
# AXLE WHINE ORDER DEFINITIONS
# =============================================================================

def build_axle_order_definitions(
    fuel_type: str,
    axle_type: str
) -> Dict[float, dict]:
    """
    Create the 10th and 20th order definitions for Axle Whine analysis.
    """
    if fuel_type not in AXLE_TARGETS:
        raise ValueError(
            f"Unsupported fuel type: {fuel_type}"
        )

    if axle_type not in AXLE_TARGETS[fuel_type]:
        raise ValueError(
            f"Unsupported axle type: {axle_type}"
        )

    target_definition = (
        AXLE_TARGETS[
            fuel_type
        ][
            axle_type
        ]
    )

    return {
        10.0: {
            "label": "10th Order",
            "harmonic": "Base",
            "target_rpm": target_definition[
                "rpm"
            ],
            "target_amp": target_definition[
                "amp"
            ],
        },

        20.0: {
            "label": "20th Order",
            "harmonic": "2nd",
            "target_rpm": target_definition[
                "rpm"
            ],
            "target_amp": target_definition[
                "amp"
            ],
        },
    }


# =============================================================================
# UNIT CONVERSION
# =============================================================================

def convert_csv_g_to_ms2_if_needed(
    headers: list,
    data: np.ndarray
) -> Tuple[np.ndarray, list]:
    """
    Convert ChA, ChB and ChC from g to m/s² when the CSV headers indicate g.

    Supported examples:

        ChA (g)
        ChB [g]
        ChC g
    """
    converted_channels = []

    if len(headers) < 4:
        return data, converted_channels

    for column_index in [
        1,
        2,
        3,
    ]:
        header = str(
            headers[
                column_index
            ]
        ).strip().lower()

        is_g_unit = (
            "(g)" in header
            or "[g]" in header
            or header.endswith(" g")
        )

        if is_g_unit:
            data[
                :,
                column_index
            ] = (
                data[
                    :,
                    column_index
                ]
                * G_TO_MS2
            )

            converted_channels.append(
                str(
                    headers[
                        column_index
                    ]
                )
            )

    return (
        data,
        converted_channels,
    )


# =============================================================================
# MEASUREMENT FILE LOADING
# =============================================================================

def load_measurement_file(
    uploaded_file,
    analysis_type: str
) -> Tuple[list, np.ndarray]:
    """
    Load and validate XLSX or CSV measurement data.

    Expected first five columns:

        Time
        ChA
        ChB
        ChC
        RPM
    """
    if uploaded_file is None:
        raise ValueError(
            "No measurement file was uploaded."
        )

    uploaded_file_name = str(
        uploaded_file.name
    )

    if "." not in uploaded_file_name:
        raise ValueError(
            "Uploaded file has no extension."
        )

    file_extension = (
        uploaded_file_name
        .rsplit(
            ".",
            1,
        )[-1]
        .lower()
    )

    if (
        uploaded_file.size
        > MAX_FILE_SIZE_MB
        * 1024
        * 1024
    ):
        raise ValueError(
            f"File exceeds the maximum allowed size "
            f"of {MAX_FILE_SIZE_MB} MB."
        )

    converted_channels = []

    # -------------------------------------------------------------------------
    # XLSX
    # -------------------------------------------------------------------------

    if file_extension == "xlsx":
        temporary_path = None

        try:
            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsx",
            ) as temporary_file:
                temporary_file.write(
                    uploaded_file.getvalue()
                )

                temporary_path = (
                    temporary_file.name
                )

            if (
                analysis_type
                == ANALYSIS_TRANSFER_CASE
            ):
                headers, data = (
                    tc_read_xlsx_numeric(
                        temporary_path
                    )
                )

            else:
                headers, data = (
                    axle_read_xlsx_numeric(
                        temporary_path
                    )
                )

        finally:
            if (
                temporary_path
                and os.path.exists(
                    temporary_path
                )
            ):
                try:
                    os.remove(
                        temporary_path
                    )
                except OSError:
                    pass

    # -------------------------------------------------------------------------
    # CSV
    # -------------------------------------------------------------------------

    elif file_extension == "csv":
        try:
            uploaded_file.seek(
                0
            )

            dataframe = pd.read_csv(
                uploaded_file,
                sep=None,
                engine="python",
            )

            headers = list(
                dataframe.columns
            )

            data = dataframe.to_numpy(
                dtype=float
            )

        except Exception as error:
            raise ValueError(
                "CSV file could not be read. "
                "Please check the delimiter and numeric data format."
            ) from error

        data, converted_channels = (
            convert_csv_g_to_ms2_if_needed(
                headers,
                data,
            )
        )

    else:
        raise ValueError(
            "Unsupported file format. "
            "Please upload an XLSX or CSV file."
        )

    data = np.asarray(
        data,
        dtype=float,
    )

    # -------------------------------------------------------------------------
    # Shape validation
    # -------------------------------------------------------------------------

    if (
        data.ndim != 2
        or data.shape[1] < 5
    ):
        raise ValueError(
            "Measurement file must contain at least five columns "
            "in this order: Time, ChA, ChB, ChC, RPM."
        )

    if data.shape[0] > MAX_ROWS:
        raise ValueError(
            f"Dataset exceeds the maximum row limit "
            f"of {MAX_ROWS:,} rows."
        )

    if data.shape[0] < 10:
        raise ValueError(
            "Dataset is too short for order analysis."
        )

    first_five_columns = data[
        :,
        :5
    ]

    if not np.all(
        np.isfinite(
            first_five_columns
        )
    ):
        raise ValueError(
            "The first five columns contain NaN, infinite "
            "or non-numeric values."
        )

    # -------------------------------------------------------------------------
    # Time and RPM validation
    # -------------------------------------------------------------------------

    time = data[
        :,
        0
    ]

    rpm = data[
        :,
        4
    ]

    # Repeated time values are allowed.
    # Only decreasing time values are rejected.
    if np.any(
        np.diff(time) < 0
    ):
        raise ValueError(
            "Time column contains decreasing values."
        )

    if np.any(
        rpm <= 0
    ):
        raise ValueError(
            "RPM column must contain only positive values."
        )

    if converted_channels:
        st.info(
            "CSV vibration channels converted from g to m/s²: "
            + ", ".join(
                converted_channels
            )
        )

    return (
        headers,
        data,
    )
    # =============================================================================
# COMMON TARGET EVALUATION HELPERS
# =============================================================================

def integrate_positive_area(
    rpm: np.ndarray,
    difference: np.ndarray
) -> float:
    """
    Integrate only the part of a curve that is above zero.

    Used for calculating total target exceedance area.

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
    Evaluate one order curve against a target curve.

    Returns:
        Peak RPM
        Peak Amplitude
        Target at Peak RPM
        Max Margin
        Max Margin %
        Exceedance Area
        Status
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


# =============================================================================
# AXLE WHINE ANALYSIS WRAPPER
# =============================================================================

def analyze_axle_orders(
    time: np.ndarray,
    rpm: np.ndarray,
    channels: Mapping[str, np.ndarray],
    order_definitions: Mapping[float, dict],
    samples_per_rev: int,
    revs_per_block: int,
    overlap: float,
    max_order: float,
    order_width: float,
    rpm_step: float,
    calibration_factor: float
) -> Tuple[
    Dict[float, dict],
    Dict[float, pd.DataFrame],
    Dict[float, pd.DataFrame],
]:
    """
    Run the Axle Whine analysis for all configured orders and channels.
    """
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

    highest_requested_order = max(
        float(order_value)
        for order_value in order_definitions.keys()
    )

    if max_order < highest_requested_order:
        raise ValueError(
            f"Max Order must be at least "
            f"{highest_requested_order:.2f}."
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

    for order_value, definition in (
        order_definitions.items()
    ):
        order_value = float(
            order_value
        )

        channel_curves = {}
        result_rows = []

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

            theta_uniform, signal_uniform, rpm_uniform = (
                axle_angular_resample(
                    time,
                    rpm,
                    signal,
                    samples_per_rev=samples_per_rev,
                )
            )

            orders, block_rpms, spectrum = (
                axle_order_map(
                    theta_uniform,
                    signal_uniform,
                    rpm_uniform,
                    samples_per_rev=samples_per_rev,
                    revs_per_block=revs_per_block,
                    overlap=overlap,
                    max_order=max_order,
                )
            )

            spectrum = np.asarray(
                spectrum,
                dtype=float,
            )

            if spectrum.ndim != 2:
                raise ValueError(
                    f"Calculated spectrum for channel {channel_name} "
                    "is not two-dimensional."
                )

            rpm_curve, amplitude_curve = (
                axle_extract_order_vs_rpm(
                    orders,
                    block_rpms,
                    spectrum,
                    target_order=order_value,
                    width=order_width,
                    rpm_step=rpm_step,
                    smooth=True,
                )
            )

            rpm_curve = np.asarray(
                rpm_curve,
                dtype=float,
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

            channel_curves[
                channel_name
            ] = {
                "rpm": rpm_curve,
                "amp": amplitude_curve,
            }

            evaluation_result = (
                evaluate_curve_against_target(
                    rpm=rpm_curve,
                    amplitude=amplitude_curve,
                    target_rpm=definition[
                        "target_rpm"
                    ],
                    target_amp=definition[
                        "target_amp"
                    ],
                )
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

            result_rows.append(
                result_row
            )

        if len(result_rows) == 0:
            raise ValueError(
                f"No results were generated for "
                f"{order_value:.2f} order."
            )

        result_dataframe = pd.DataFrame(
            result_rows
        )

        first_channel_name = next(
            iter(
                channel_curves
            )
        )

        common_rpm = np.asarray(
            channel_curves[
                first_channel_name
            ]["rpm"],
            dtype=float,
        )

        curve_dataframe = pd.DataFrame(
            {
                "RPM": common_rpm
            }
        )

        for channel_name, channel_curve in (
            channel_curves.items()
        ):
            curve_dataframe[
                channel_name
            ] = np.interp(
                common_rpm,
                np.asarray(
                    channel_curve[
                        "rpm"
                    ],
                    dtype=float,
                ),
                np.asarray(
                    channel_curve[
                        "amp"
                    ],
                    dtype=float,
                ),
            )

        curve_dataframe[
            "Target"
        ] = np.interp(
            common_rpm,
            np.asarray(
                definition[
                    "target_rpm"
                ],
                dtype=float,
            ),
            np.asarray(
                definition[
                    "target_amp"
                ],
                dtype=float,
            ),
        )

        curves_by_order[
            order_value
        ] = channel_curves

        results_by_order[
            order_value
        ] = result_dataframe

        raw_curves_by_order[
            order_value
        ] = curve_dataframe

    return (
        curves_by_order,
        results_by_order,
        raw_curves_by_order,
    )
    # =============================================================================
# RESULT PLOTTING
# =============================================================================

def plot_order_comparison(
    order_label: str,
    channel_curves: Mapping[str, dict],
    target_rpm: Optional[np.ndarray],
    target_amp: Optional[np.ndarray],
    vin_number: str,
    analysis_type: str,
    vehicle_configuration: str
):
    """
    Plot ChA, ChB and ChC order curves versus RPM.

    The target curve is plotted only when it is available.
    """
    figure, axis = plt.subplots(
        figsize=(12, 7)
    )

    for channel_name, curve in channel_curves.items():
        axis.plot(
            curve["rpm"],
            curve["amp"],
            label=channel_name,
            linewidth=2,
        )

    if (
        target_rpm is not None
        and target_amp is not None
    ):
        axis.plot(
            target_rpm,
            target_amp,
            label="Target Curve",
            linewidth=4,
        )

    axis.set_xlabel(
        "RPM"
    )

    axis.set_ylabel(
        "Order Amplitude [m/s²]"
    )

    axis.set_title(
        f"{order_label} vs RPM | "
        f"VIN: {vin_number} | "
        f"{analysis_type} | "
        f"{vehicle_configuration}"
    )

    axis.grid(
        True,
        alpha=0.3,
    )

    axis.legend()

    figure.tight_layout()

    return figure


def create_order_map_figure(
    time: np.ndarray,
    rpm: np.ndarray,
    signal: np.ndarray,
    selected_channel: str,
    analysis_type: str,
    vin_number: str,
    samples_per_rev: int,
    revs_per_block: int,
    overlap: float,
    max_order: float,
    calibration_factor: float
):
    """
    Create an Order Map / Waterfall using the selected analysis engine.

    Axle Whine:
        order_analysis.py

    Transfer Case:
        transfer_case_analysis.py
    """
    if (
        analysis_type
        == ANALYSIS_TRANSFER_CASE
    ):
        (
            theta_uniform,
            signal_uniform,
            rpm_uniform,
        ) = tc_angular_resample(
            time,
            rpm,
            signal,
            samples_per_rev=samples_per_rev,
        )

        (
            orders,
            block_rpms,
            spectrum,
        ) = tc_order_map(
            theta_uniform,
            signal_uniform,
            rpm_uniform,
            samples_per_rev=samples_per_rev,
            revs_per_block=revs_per_block,
            overlap=overlap,
            max_order=max_order,
        )

    else:
        (
            theta_uniform,
            signal_uniform,
            rpm_uniform,
        ) = axle_angular_resample(
            time,
            rpm,
            signal,
            samples_per_rev=samples_per_rev,
        )

        (
            orders,
            block_rpms,
            spectrum,
        ) = axle_order_map(
            theta_uniform,
            signal_uniform,
            rpm_uniform,
            samples_per_rev=samples_per_rev,
            revs_per_block=revs_per_block,
            overlap=overlap,
            max_order=max_order,
        )

    orders = np.asarray(
        orders,
        dtype=float,
    )

    block_rpms = np.asarray(
        block_rpms,
        dtype=float,
    )

    spectrum = np.asarray(
        spectrum,
        dtype=float,
    )

    if spectrum.ndim != 2:
        raise ValueError(
            "Order spectrum must be two-dimensional."
        )

    if spectrum.shape[0] != len(block_rpms):
        raise ValueError(
            "Order spectrum row count does not match the RPM vector."
        )

    if spectrum.shape[1] != len(orders):
        raise ValueError(
            "Order spectrum column count does not match the order axis."
        )

    sort_indices = np.argsort(
        block_rpms,
        kind="stable",
    )

    sorted_rpms = block_rpms[
        sort_indices
    ]

    sorted_spectrum = spectrum[
        sort_indices,
        :
    ]

    decibel_spectrum = (
        20.0
        * np.log10(
            np.maximum(
                sorted_spectrum
                * calibration_factor,
                1e-12,
            )
        )
    )

    figure, axis = plt.subplots(
        figsize=(12, 7)
    )

    image = axis.imshow(
        decibel_spectrum,
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

    figure.colorbar(
        image,
        ax=axis,
        label="Amplitude [dB re 1 m/s²]",
    )

    axis.set_xlabel(
        "Order"
    )

    axis.set_ylabel(
        "RPM"
    )

    axis.set_title(
        f"Order Map / Waterfall - "
        f"{selected_channel} | "
        f"VIN: {vin_number} | "
        f"{analysis_type}"
    )

    figure.tight_layout()

    return figure


# =============================================================================
# EXCEL SHEET FORMATTING
# =============================================================================

def format_comparison_sheet(
    writer,
    sheet_name: str
) -> None:
    """
    Format a comparison sheet and highlight PASS, FAIL and INFO rows.
    """
    worksheet = writer.book[
        sheet_name
    ]

    header_fill = PatternFill(
        start_color="D9EAF7",
        end_color="D9EAF7",
        fill_type="solid",
    )

    pass_fill = PatternFill(
        start_color="C6EFCE",
        end_color="C6EFCE",
        fill_type="solid",
    )

    fail_fill = PatternFill(
        start_color="FFC7CE",
        end_color="FFC7CE",
        fill_type="solid",
    )

    info_fill = PatternFill(
        start_color="D9EAF7",
        end_color="D9EAF7",
        fill_type="solid",
    )

    for cell in worksheet[1]:
        cell.font = Font(
            bold=True
        )

        cell.fill = header_fill

    status_column = None

    for cell in worksheet[1]:
        if cell.value == "Status":
            status_column = cell.column
            break

    if status_column is not None:
        for row_index in range(
            2,
            worksheet.max_row + 1,
        ):
            status_value = worksheet.cell(
                row=row_index,
                column=status_column,
            ).value

            if status_value == "PASS":
                row_fill = pass_fill

            elif status_value == "FAIL":
                row_fill = fail_fill

            elif status_value == "INFO":
                row_fill = info_fill

            else:
                row_fill = None

            if row_fill is not None:
                for column_index in range(
                    1,
                    worksheet.max_column + 1,
                ):
                    worksheet.cell(
                        row=row_index,
                        column=column_index,
                    ).fill = row_fill

    for column_cells in worksheet.columns:
        column_letter = (
            column_cells[0]
            .column_letter
        )

        worksheet.column_dimensions[
            column_letter
        ].width = 22


def format_curve_sheet(
    writer,
    sheet_name: str
) -> None:
    """
    Format a raw curve sheet.
    """
    worksheet = writer.book[
        sheet_name
    ]

    header_fill = PatternFill(
        start_color="D9EAF7",
        end_color="D9EAF7",
        fill_type="solid",
    )

    for cell in worksheet[1]:
        cell.font = Font(
            bold=True
        )

        cell.fill = header_fill

    for column_cells in worksheet.columns:
        column_letter = (
            column_cells[0]
            .column_letter
        )

        worksheet.column_dimensions[
            column_letter
        ].width = 16


# =============================================================================
# EXCEL CURVE PLOT
# =============================================================================

def create_curve_plot_png(
    curve_dataframe: pd.DataFrame,
    order_label: str,
    vin_number: str,
    analysis_type: str,
    vehicle_configuration: str
) -> BytesIO:
    """
    Create an order-vs-RPM PNG image for embedding in Excel.
    """
    figure, axis = plt.subplots(
        figsize=(14, 8)
    )

    for channel_name in CHANNEL_NAMES:
        if (
            channel_name
            in curve_dataframe.columns
        ):
            axis.plot(
                curve_dataframe["RPM"],
                curve_dataframe[channel_name],
                label=channel_name,
                linewidth=2,
            )

    if (
        "Target"
        in curve_dataframe.columns
    ):
        axis.plot(
            curve_dataframe["RPM"],
            curve_dataframe["Target"],
            label="Target Curve",
            linewidth=5,
        )

    axis.set_title(
        f"{order_label} vs RPM | "
        f"VIN: {vin_number} | "
        f"{analysis_type} | "
        f"{vehicle_configuration}",
        fontsize=16,
    )

    axis.set_xlabel(
        "RPM",
        fontsize=13,
    )

    axis.set_ylabel(
        "Order Amplitude [m/s²]",
        fontsize=13,
    )

    axis.grid(
        True,
        alpha=0.3,
    )

    axis.legend(
        loc="upper right",
        fontsize=12,
    )

    rpm_minimum = min(
        1000.0,
        float(
            curve_dataframe[
                "RPM"
            ].min()
        ),
    )

    rpm_maximum = max(
        4500.0,
        float(
            curve_dataframe[
                "RPM"
            ].max()
        ),
    )

    axis.set_xlim(
        rpm_minimum,
        rpm_maximum,
    )

    figure.tight_layout()

    image_buffer = BytesIO()

    figure.savefig(
        image_buffer,
        format="png",
        dpi=180,
        bbox_inches="tight",
    )

    plt.close(
        figure
    )

    image_buffer.seek(
        0
    )

    return image_buffer


def add_curve_plot_to_sheet(
    writer,
    sheet_name: str,
    curve_dataframe: pd.DataFrame,
    order_label: str,
    vin_number: str,
    analysis_type: str,
    vehicle_configuration: str
) -> None:
    """
    Add the calculated order curve PNG to an Excel worksheet.
    """
    worksheet = writer.book[
        sheet_name
    ]

    image_buffer = create_curve_plot_png(
        curve_dataframe=curve_dataframe,
        order_label=order_label,
        vin_number=vin_number,
        analysis_type=analysis_type,
        vehicle_configuration=vehicle_configuration,
    )

    excel_image = XLImage(
        image_buffer
    )

    excel_image.width = 900
    excel_image.height = 520

    worksheet.add_image(
        excel_image,
        "G2",
    )


# =============================================================================
# EXCEL REPORT GENERATION
# =============================================================================

def make_excel_report(
    vehicle_information: dict,
    results_by_order: Mapping[
        float,
        pd.DataFrame
    ],
    curves_by_order: Mapping[
        float,
        pd.DataFrame
    ],
    order_definitions: Mapping[
        float,
        dict
    ]
) -> BytesIO:
    """
    Create the complete Excel report.

    Sheets:
        Vehicle Info
        One comparison sheet per order
        One curve-data sheet per order
        Embedded PNG plot in every curve sheet
    """
    output = BytesIO()

    vin_number = vehicle_information[
        "VIN"
    ]

    analysis_type = vehicle_information[
        "Analysis Type"
    ]

    vehicle_configuration = vehicle_information[
        "Vehicle Configuration"
    ]

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        pd.DataFrame(
            [
                vehicle_information
            ]
        ).to_excel(
            writer,
            sheet_name="Vehicle Info",
            index=False,
        )

        for order_value, result_dataframe in (
            results_by_order.items()
        ):
            comparison_sheet_name = (
                f"{str(order_value).replace('.', '_')} "
                f"Comparison"
            )[:31]

            result_dataframe.to_excel(
                writer,
                sheet_name=comparison_sheet_name,
                index=False,
            )

            format_comparison_sheet(
                writer,
                comparison_sheet_name,
            )

        for order_value, curve_dataframe in (
            curves_by_order.items()
        ):
            curve_sheet_name = (
                f"{str(order_value).replace('.', '_')} "
                f"Curves"
            )[:31]

            curve_dataframe.to_excel(
                writer,
                sheet_name=curve_sheet_name,
                index=False,
            )

            format_curve_sheet(
                writer,
                curve_sheet_name,
            )

            add_curve_plot_to_sheet(
                writer=writer,
                sheet_name=curve_sheet_name,
                curve_dataframe=curve_dataframe,
                order_label=order_definitions[
                    order_value
                ]["label"],
                vin_number=vin_number,
                analysis_type=analysis_type,
                vehicle_configuration=vehicle_configuration,
            )

    output.seek(
        0
    )

    return output
    # =============================================================================
# USER INTERFACE
# =============================================================================

st.subheader(
    "Vehicle Information"
)

vehicle_column, analysis_column, option_column_1, option_column_2 = (
    st.columns(4)
)


# -----------------------------------------------------------------------------
# VIN
# -----------------------------------------------------------------------------

with vehicle_column:
    vin_number = st.text_input(
        "VIN Number",
        placeholder="Enter 17-character VIN",
        max_chars=17,
    ).upper().strip()


vin_valid = bool(
    re.fullmatch(
        r"[A-Z0-9]{17}",
        vin_number,
    )
)


# -----------------------------------------------------------------------------
# Analysis type
# -----------------------------------------------------------------------------

with analysis_column:
    analysis_type = st.selectbox(
        "Analysis Type",
        [
            ANALYSIS_AXLE,
            ANALYSIS_TRANSFER_CASE,
        ],
        disabled=not vin_valid,
    )


# -----------------------------------------------------------------------------
# Vehicle configuration
# -----------------------------------------------------------------------------

if analysis_type == ANALYSIS_AXLE:

    with option_column_1:
        fuel_type = st.selectbox(
            "Fuel Type",
            [
                "Select fuel type",
                "Diesel",
                "Gasoline",
            ],
            disabled=not vin_valid,
        )

    with option_column_2:
        axle_type = st.selectbox(
            "Axle Type",
            [
                "Select axle type",
                "Front Axle",
                "Rear Axle",
            ],
            disabled=not vin_valid,
        )

    vehicle_configuration = (
        f"{fuel_type} | {axle_type}"
    )

else:
    fuel_type = "N/A"

    axle_type = (
        "Transfer Case / 6th Gear"
    )

    vehicle_configuration = (
        "Transfer Case | 6th Gear"
    )

    with option_column_1:
        st.text_input(
            "Gear",
            value="6th Gear",
            disabled=True,
        )

    with option_column_2:
        st.text_input(
            "Component",
            value="Transfer Case",
            disabled=True,
        )


if vin_number and not vin_valid:
    st.error(
        "VIN must be exactly 17 characters "
        "and contain only letters and numbers."
    )


# =============================================================================
# MEASUREMENT FILE
# =============================================================================

st.markdown(
    """
    <div class="section-title">
        📂 Measurement Data
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Upload Measurement File",
    type=[
        "xlsx",
        "csv",
    ],
    disabled=not vin_valid,
    help=(
        "Supported formats: XLSX and CSV. "
        "Expected columns: Time, ChA, ChB, ChC, RPM."
    ),
)


# =============================================================================
# INPUT COMPLETENESS CHECK
# =============================================================================

if analysis_type == ANALYSIS_AXLE:

    can_continue = (
        vin_valid
        and fuel_type
        != "Select fuel type"
        and axle_type
        != "Select axle type"
        and uploaded_file
        is not None
    )

else:
    can_continue = (
        vin_valid
        and uploaded_file
        is not None
    )


if not can_continue:

    if not vin_valid:
        st.warning(
            "Please enter a valid 17-character VIN "
            "before continuing."
        )

    elif analysis_type == ANALYSIS_AXLE:
        st.warning(
            "Please select fuel type, select axle type "
            "and upload a measurement file."
        )

    else:
        st.warning(
            "Please upload a measurement file."
        )

    st.stop()


# =============================================================================
# ANALYSIS-SPECIFIC CONFIGURATION
# =============================================================================

if analysis_type == ANALYSIS_AXLE:

    order_definitions = (
        build_axle_order_definitions(
            fuel_type=fuel_type,
            axle_type=axle_type,
        )
    )

    fixed_samples_per_rev = 512

    fixed_revs_per_block = 8

    fixed_overlap = 0.75

    fixed_rpm_step = 10.0

    fixed_calibration_factor = 1.0

    minimum_max_order = 20

    default_max_order = 30

else:

    order_definitions = (
        TRANSFER_CASE_ORDERS
    )

    fixed_samples_per_rev = 512

    # 20 revolutions produce 0.05 order resolution.
    fixed_revs_per_block = 20

    fixed_overlap = 0.75

    fixed_rpm_step = 10.0

    fixed_calibration_factor = 1.0

    # Highest requested Transfer Case order is 170.10.
    minimum_max_order = 171

    default_max_order = 200


# =============================================================================
# READY STATUS
# =============================================================================

st.success(
    "Vehicle information and measurement file "
    "are ready for analysis."
)


information_columns = st.columns(
    4
)

information_columns[0].metric(
    "VIN",
    vin_number,
)

information_columns[1].metric(
    "Analysis",
    analysis_type,
)

information_columns[2].metric(
    "Fuel Type",
    fuel_type,
)

information_columns[3].metric(
    "Configuration",
    axle_type,
)


# =============================================================================
# ANALYSIS SETTINGS
# =============================================================================

st.markdown(
    """
    <div class="section-title">
        ⚙️ Analysis Settings
    </div>
    """,
    unsafe_allow_html=True,
)
with st.expander(
    "Advanced Settings",
    expanded=False,
):

    selected_channel = st.selectbox(
        "Order Map Channel",
        CHANNEL_NAMES,
    )

    max_order = st.slider(
        "Max Order",
        min_value=minimum_max_order,
        max_value=250,
        value=default_max_order,
        step=1,
    )

    order_width = st.number_input(
        "Order Width",
        min_value=0.05,
        max_value=2.0,
        value=0.15,
        step=0.05,
        format="%.2f",
    )


# =============================================================================
# SETTINGS SUMMARY
# =============================================================================

settings_columns = st.columns(
    5
)

settings_columns[0].metric(
    "Samples / Rev",
    fixed_samples_per_rev,
)

settings_columns[1].metric(
    "Revs / Block",
    fixed_revs_per_block,
)

settings_columns[2].metric(
    "Overlap",
    f"{fixed_overlap * 100:.0f}%",
)

settings_columns[3].metric(
    "RPM Step",
    f"{fixed_rpm_step:.0f}",
)

settings_columns[4].metric(
    "Max Order",
    max_order,
)

# =============================================================================
# MODULE STATUS DISPLAY
# =============================================================================

with st.expander(
    "System Status",
    expanded=False,
):
    st.success(
        "Axle Whine analysis engine loaded successfully."
    )

    st.success(
        "Transfer Case analysis engine loaded successfully."
    )

    st.write(
        "Transfer Case module validation:"
    )

    st.json(
        transfer_case_validation
    )

    st.write(
        {
            "App Status": "READY",
            "Supported Analysis Types": [
                ANALYSIS_AXLE,
                ANALYSIS_TRANSFER_CASE,
            ],
            "Supported File Types": [
                "XLSX",
                "CSV",
            ],
            "Supported Channels": CHANNEL_NAMES,
            "CSV Unit Conversion": "g to m/s²",
        }
    )
# =============================================================================
# RUN ANALYSIS
# =============================================================================
# =============================================================================
# SESSION STATE SAFETY
# =============================================================================

def clear_analysis_session_state() -> None:
    """
    Clear previously stored analysis results.

    This prevents results from a previous vehicle, file or analysis
    configuration from remaining visible after the inputs change.
    """
    state_keys_to_clear = [
        "analysis_completed",
        "analysis_type_result",
        "time_result",
        "rpm_result",
        "channels_result",
        "curves_by_order",
        "results_by_order",
        "raw_curves_by_order",
        "order_definitions_result",
        "overall_status",
        "vehicle_configuration_result",
        "selected_channel_result",
        "analysis_settings_result",
        "excel_report",
        "vehicle_information",
        "vin_result",
    ]

    for state_key in state_keys_to_clear:
        if state_key in st.session_state:
            del st.session_state[state_key]


def build_input_signature(
    vin_value: str,
    selected_analysis: str,
    selected_fuel: str,
    selected_axle: str,
    uploaded_measurement_file,
    selected_max_order: float,
    selected_order_width: float,
    selected_map_channel: str,
) -> tuple:
    """
    Build a signature representing the current analysis inputs.

    If this signature changes, stored results are cleared.
    """
    if uploaded_measurement_file is None:
        uploaded_file_name = None
        uploaded_file_size = None

    else:
        uploaded_file_name = str(
            uploaded_measurement_file.name
        )

        uploaded_file_size = int(
            uploaded_measurement_file.size
        )

    return (
        vin_value,
        selected_analysis,
        selected_fuel,
        selected_axle,
        uploaded_file_name,
        uploaded_file_size,
        float(selected_max_order),
        float(selected_order_width),
        selected_map_channel,
    )


current_input_signature = build_input_signature(
    vin_value=vin_number,
    selected_analysis=analysis_type,
    selected_fuel=fuel_type,
    selected_axle=axle_type,
    uploaded_measurement_file=uploaded_file,
    selected_max_order=max_order,
    selected_order_width=order_width,
    selected_map_channel=selected_channel,
)


previous_input_signature = st.session_state.get(
    "input_signature"
)


if (
    previous_input_signature is not None
    and previous_input_signature != current_input_signature
):
    clear_analysis_session_state()


st.session_state[
    "input_signature"
] = current_input_signature
if st.button(
    "Run Analysis",
    type="primary",
    use_container_width=True,
):

    try:
        # ---------------------------------------------------------------------
        # Load measurement data
        # ---------------------------------------------------------------------

        headers, data = load_measurement_file(
            uploaded_file=uploaded_file,
            analysis_type=analysis_type,
        )

        time = np.asarray(
            data[:, 0],
            dtype=float,
        )

        rpm = np.asarray(
            data[:, 4],
            dtype=float,
        )

        channels = {
            "ChA": np.asarray(
                data[:, 1],
                dtype=float,
            ),
            "ChB": np.asarray(
                data[:, 2],
                dtype=float,
            ),
            "ChC": np.asarray(
                data[:, 3],
                dtype=float,
            ),
        }

        # ---------------------------------------------------------------------
        # Execute selected analysis engine
        # ---------------------------------------------------------------------

        with st.spinner(
            "Analysis is running..."
        ):

            if (
                analysis_type
                == ANALYSIS_TRANSFER_CASE
            ):
                (
                    curves_by_order,
                    results_by_order,
                    raw_curves_by_order,
                ) = analyze_transfer_case_orders(
                    time=time,
                    rpm=rpm,
                    channels=channels,
                    order_definitions=order_definitions,
                    samples_per_rev=fixed_samples_per_rev,
                    revs_per_block=fixed_revs_per_block,
                    overlap=fixed_overlap,
                    max_order=max_order,
                    order_width=order_width,
                    rpm_step=fixed_rpm_step,
                    calibration_factor=(
                        fixed_calibration_factor
                    ),
                )

            else:
                (
                    curves_by_order,
                    results_by_order,
                    raw_curves_by_order,
                ) = analyze_axle_orders(
                    time=time,
                    rpm=rpm,
                    channels=channels,
                    order_definitions=order_definitions,
                    samples_per_rev=fixed_samples_per_rev,
                    revs_per_block=fixed_revs_per_block,
                    overlap=fixed_overlap,
                    max_order=max_order,
                    order_width=order_width,
                    rpm_step=fixed_rpm_step,
                    calibration_factor=(
                        fixed_calibration_factor
                    ),
                )

        # ---------------------------------------------------------------------
        # Validate returned result structures
        # ---------------------------------------------------------------------

        if not curves_by_order:
            raise ValueError(
                "Analysis returned no order curves."
            )

        if not results_by_order:
            raise ValueError(
                "Analysis returned no result tables."
            )

        if not raw_curves_by_order:
            raise ValueError(
                "Analysis returned no raw curve data."
            )

        # ---------------------------------------------------------------------
        # Overall status
        # ---------------------------------------------------------------------

        evaluated_statuses = []

        for result_dataframe in (
            results_by_order.values()
        ):
            if (
                "Status"
                not in result_dataframe.columns
            ):
                raise ValueError(
                    "A result table does not contain "
                    "the Status column."
                )

            evaluated_rows = result_dataframe[
                result_dataframe[
                    "Status"
                ] != "INFO"
            ]

            if len(
                evaluated_rows
            ) > 0:
                evaluated_statuses.extend(
                    evaluated_rows[
                        "Status"
                    ].tolist()
                )

        if len(
            evaluated_statuses
        ) == 0:
            overall_status = "INFO"

        elif any(
            status == "FAIL"
            for status in evaluated_statuses
        ):
            overall_status = "FAIL"

        else:
            overall_status = "PASS"

               # ---------------------------------------------------------------------
        # Overall assessment display
        # ---------------------------------------------------------------------

        st.markdown(
            """
            <div class="section-title">
                📊 Overall Assessment
            </div>
            """,
            unsafe_allow_html=True,
        )

        if overall_status == "PASS":
            st.success(
                "Overall Assessment: PASS"
            )

        elif overall_status == "FAIL":
            st.error(
                "Overall Assessment: FAIL"
            )

        else:
            st.info(
                "Overall Assessment: INFO"
            )
        # ---------------------------------------------------------------------
        # Vehicle and analysis metadata
        # ---------------------------------------------------------------------

        vehicle_information = {
            "VIN": vin_number,
            "Analysis Type": analysis_type,
            "Fuel Type": fuel_type,
            "Vehicle Configuration": (
                vehicle_configuration
            ),
            "Target Orders": ", ".join(
                str(
                    order_value
                )
                for order_value in (
                    order_definitions.keys()
                )
            ),
            "Order Width": order_width,
            "RPM Step": fixed_rpm_step,
            "Samples per Rev": (
                fixed_samples_per_rev
            ),
            "Revs per Block": (
                fixed_revs_per_block
            ),
            "Overlap": fixed_overlap,
            "Calibration Factor": (
                fixed_calibration_factor
            ),
            "Max Order": max_order,
            "Overall Assessment": (
                overall_status
            ),
        }

              # ---------------------------------------------------------------------
        # Excel report
        # ---------------------------------------------------------------------

        excel_report = make_excel_report(
            vehicle_information=(
                vehicle_information
                
            ),
            results_by_order=(
                results_by_order
            ),
            curves_by_order=(
                raw_curves_by_order
            ),
            order_definitions=(
                order_definitions
            ),
        )

        # ---------------------------------------------------------------------
        # Store results in Streamlit session state
        # ---------------------------------------------------------------------

        st.session_state[
            "excel_report"
        ] = excel_report

        st.session_state[
            "vehicle_information"
        ] = vehicle_information

        st.session_state[
            "vin_result"
        ] = vin_number

        st.session_state[
            "analysis_completed"
        ] = True

        st.session_state[
            "analysis_type_result"
        ] = analysis_type

        st.session_state[
            "time_result"
        ] = time

        st.session_state[
            "rpm_result"
        ] = rpm

        st.session_state[
            "channels_result"
        ] = channels

        st.session_state[
            "curves_by_order"
        ] = curves_by_order

        st.session_state[
            "results_by_order"
        ] = results_by_order

        st.session_state[
            "raw_curves_by_order"
        ] = raw_curves_by_order

        st.session_state[
            "order_definitions_result"
        ] = order_definitions

        st.session_state[
            "overall_status"
        ] = overall_status

        st.session_state[
            "vehicle_configuration_result"
        ] = vehicle_configuration

        st.session_state[
            "selected_channel_result"
        ] = selected_channel

        st.session_state[
            "analysis_settings_result"
        ] = {
            "samples_per_rev": (
                fixed_samples_per_rev
            ),
            "revs_per_block": (
                fixed_revs_per_block
            ),
            "overlap": fixed_overlap,
            "max_order": max_order,
            "calibration_factor": (
                fixed_calibration_factor
            ),
        }

    except Exception as error:
        st.session_state[
            "analysis_completed"
        ] = False

        st.error(
            "An error occurred while running the analysis."
        )

        st.exception(
            error
        )
        # =============================================================================
# PERSISTENT RESULT DISPLAY
# =============================================================================

if st.session_state.get(
    "analysis_completed",
    False,
):
    result_analysis_type = st.session_state[
        "analysis_type_result"
    ]

    result_time = st.session_state[
        "time_result"
    ]

    result_rpm = st.session_state[
        "rpm_result"
    ]

    result_channels = st.session_state[
        "channels_result"
    ]

    result_curves_by_order = st.session_state[
        "curves_by_order"
    ]

    result_tables_by_order = st.session_state[
        "results_by_order"
    ]

    result_raw_curves_by_order = st.session_state[
        "raw_curves_by_order"
    ]

    result_order_definitions = st.session_state[
        "order_definitions_result"
    ]

    result_overall_status = st.session_state[
        "overall_status"
    ]

    result_vehicle_configuration = st.session_state[
        "vehicle_configuration_result"
    ]

    result_selected_channel = st.session_state[
        "selected_channel_result"
    ]

    result_analysis_settings = st.session_state[
        "analysis_settings_result"
    ]

    result_vin = st.session_state.get(
        "vin_result",
        vin_number,
    )

    stored_excel_report = st.session_state.get(
        "excel_report"
    )


    # =========================================================================
    # OVERALL ASSESSMENT
    # =========================================================================

    st.divider()

    st.subheader(
        "Overall Assessment"
    )

    if result_overall_status == "PASS":
        st.success(
            "Overall Assessment: PASS"
        )

    elif result_overall_status == "FAIL":
        st.error(
            "Overall Assessment: FAIL"
        )

    else:
        st.info(
            "Overall Assessment: INFO"
        )


    # =========================================================================
    # EXCEL REPORT DOWNLOAD
    # =========================================================================

    if stored_excel_report is not None:
        st.download_button(
            label="Download Excel Report",
            data=stored_excel_report,
            file_name=(
                f"{result_vin}_"
                f"{result_analysis_type.replace(' ', '_')}"
                f"_report.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
            key="persistent_excel_download",
        )


    # =========================================================================
    # SMALL DISPLAY HELPERS
    # =========================================================================

    def get_channel_peak(
        result_dataframe: pd.DataFrame,
        channel_name: str
    ) -> float:
        """
        Return the peak amplitude for one channel.
        """
        matching_rows = result_dataframe.loc[
            result_dataframe[
                "Channel"
            ] == channel_name,
            "Peak Amplitude [m/s²]",
        ]

        if len(matching_rows) == 0:
            return float("nan")

        return float(
            matching_rows.iloc[0]
        )


    def determine_order_status(
        result_dataframe: pd.DataFrame
    ) -> str:
        """
        Determine order-level PASS, FAIL or INFO status.
        """
        if (
            "Status"
            not in result_dataframe.columns
        ):
            return "INFO"

        evaluated_rows = result_dataframe[
            result_dataframe[
                "Status"
            ] != "INFO"
        ]

        if len(evaluated_rows) == 0:
            return "INFO"

        if (
            evaluated_rows[
                "Status"
            ] == "PASS"
        ).all():
            return "PASS"

        return "FAIL"


    def display_order_result(
        order_value: float,
        definition: Mapping[str, object],
        result_dataframe: pd.DataFrame,
        channel_curves: Mapping[str, dict],
        show_separator: bool = True
    ) -> None:
        """
        Display KPI cards, comparison plot and result table for one order.
        """
        order_status = determine_order_status(
            result_dataframe
        )

        st.subheader(
            f"{definition['label']} Result Summary"
        )

        metric_columns = st.columns(
            4
        )

        for metric_column, channel_name in zip(
            metric_columns[:3],
            CHANNEL_NAMES,
        ):
            peak_value = get_channel_peak(
                result_dataframe,
                channel_name,
            )

            if np.isfinite(
                peak_value
            ):
                displayed_peak = (
                    f"{peak_value:.2f} m/s²"
                )
            else:
                displayed_peak = "N/A"

            metric_column.metric(
                f"Peak {channel_name}",
                displayed_peak,
            )

        metric_columns[3].metric(
            "Assessment",
            order_status,
        )

        comparison_figure = plot_order_comparison(
            order_label=definition[
                "label"
            ],
            channel_curves=channel_curves,
            target_rpm=definition.get(
                "target_rpm"
            ),
            target_amp=definition.get(
                "target_amp"
            ),
            vin_number=result_vin,
            analysis_type=(
                result_analysis_type
            ),
            vehicle_configuration=(
                result_vehicle_configuration
            ),
        )

        st.pyplot(
            comparison_figure,
            use_container_width=True,
        )

        plt.close(
            comparison_figure
        )

        if (
            definition.get(
                "target_rpm"
            )
            is not None
            and definition.get(
                "target_amp"
            )
            is not None
        ):
            table_title = (
                f"{definition['label']} "
                "Target Compliance"
            )

        else:
            table_title = (
                f"{definition['label']} "
                "Informational Results"
            )

        st.subheader(
            table_title
        )

        st.dataframe(
            result_dataframe,
            use_container_width=True,
            hide_index=True,
        )

        if order_status == "PASS":
            st.success(
                f"{definition['label']} Assessment: PASS"
            )

        elif order_status == "FAIL":
            st.error(
                f"{definition['label']} Assessment: FAIL"
            )

        else:
            st.info(
                f"{definition['label']} Assessment: INFO — "
                "No target has been defined for this harmonic."
            )

        if show_separator:
            st.markdown(
                "---"
            )


    # =========================================================================
    # TRANSFER CASE RESULT TABS
    # =========================================================================

    if (
        result_analysis_type
        == ANALYSIS_TRANSFER_CASE
    ):
        (
            gear_mesh_results_tab,
            order_map_tab,
            raw_results_tab,
        ) = st.tabs(
            [
                "Gear Mesh Order Results",
                "Order Map / Waterfall",
                "Raw Results",
            ]
        )

        with gear_mesh_results_tab:
            # Show base orders first, followed by second harmonics.
            transfer_case_display_order = [
                63.0,
                85.05,
                126.0,
                170.10,
            ]

            for display_index, order_value in enumerate(
                transfer_case_display_order
            ):
                if (
                    order_value
                    not in result_order_definitions
                ):
                    continue

                if (
                    order_value
                    not in result_tables_by_order
                ):
                    continue

                display_order_result(
                    order_value=order_value,
                    definition=(
                        result_order_definitions[
                            order_value
                        ]
                    ),
                    result_dataframe=(
                        result_tables_by_order[
                            order_value
                        ]
                    ),
                    channel_curves=(
                        result_curves_by_order[
                            order_value
                        ]
                    ),
                    show_separator=(
                        display_index
                        < len(
                            transfer_case_display_order
                        ) - 1
                    ),
                )


    # =========================================================================
    # AXLE WHINE RESULT TABS
    # =========================================================================

    else:
        (
            order_10_tab,
            order_20_tab,
            order_map_tab,
            raw_results_tab,
        ) = st.tabs(
            [
                "10th Order Target Comparison",
                "20th Order Target Comparison",
                "Order Map / Waterfall",
                "Raw Results",
            ]
        )

        axle_order_tabs = {
            10.0: order_10_tab,
            20.0: order_20_tab,
        }

        for order_value, result_tab in (
            axle_order_tabs.items()
        ):
            if (
                order_value
                not in result_order_definitions
            ):
                continue

            with result_tab:
                display_order_result(
                    order_value=order_value,
                    definition=(
                        result_order_definitions[
                            order_value
                        ]
                    ),
                    result_dataframe=(
                        result_tables_by_order[
                            order_value
                        ]
                    ),
                    channel_curves=(
                        result_curves_by_order[
                            order_value
                        ]
                    ),
                    show_separator=False,
                )


    # =========================================================================
    # ORDER MAP / WATERFALL TAB
    # =========================================================================

    with order_map_tab:
        st.subheader(
            f"Order Map / Waterfall - "
            f"{result_selected_channel}"
        )

        map_figure = create_order_map_figure(
            time=result_time,
            rpm=result_rpm,
            signal=result_channels[
                result_selected_channel
            ],
            selected_channel=(
                result_selected_channel
            ),
            analysis_type=(
                result_analysis_type
            ),
            vin_number=result_vin,
            samples_per_rev=(
                result_analysis_settings[
                    "samples_per_rev"
                ]
            ),
            revs_per_block=(
                result_analysis_settings[
                    "revs_per_block"
                ]
            ),
            overlap=(
                result_analysis_settings[
                    "overlap"
                ]
            ),
            max_order=(
                result_analysis_settings[
                    "max_order"
                ]
            ),
            calibration_factor=(
                result_analysis_settings[
                    "calibration_factor"
                ]
            ),
        )

        st.pyplot(
            map_figure,
            use_container_width=True,
        )

        plt.close(
            map_figure
        )

        if (
            result_analysis_type
            == ANALYSIS_TRANSFER_CASE
        ):
            st.caption(
                "Transfer Case order map calculated using "
                "20 revolutions per FFT block, providing "
                "0.05 order resolution."
            )


    # =========================================================================
    # RAW RESULTS TAB
    # =========================================================================

    with raw_results_tab:
        for order_value, curve_dataframe in (
            result_raw_curves_by_order.items()
        ):
            definition = (
                result_order_definitions[
                    order_value
                ]
            )

            st.subheader(
                f"{definition['label']} "
                "Raw Curve Data"
            )

            st.dataframe(
                curve_dataframe,
                use_container_width=True,
                hide_index=True,
            )

            st.markdown(
                "---"
            )
            # =============================================================================
