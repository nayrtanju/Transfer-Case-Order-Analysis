from transfer_case_analysis import (
    read_xlsx_numeric,
    angular_resample,
    order_map,
    extract_order_vs_rpm,
    analyze_transfer_case_orders,
    TRANSFER_CASE_ORDERS
)


st.set_page_config(
    page_title="NVH Analysis Suite",
    layout="wide"
)

st.title("NVH Analysis Suite")


try:
    from order_analysis import (
        read_xlsx_numeric,
        angular_resample,
        order_map,
        extract_order_vs_rpm
    )
except Exception:
    st.error("order_analysis.py yüklenirken hata oluştu")
    st.code(traceback.format_exc())
    st.stop()


MAX_FILE_SIZE_MB = 500
MAX_ROWS = 3000000
G_TO_MS2 = 9.80665


AXLE_TARGETS = {
    "Diesel": {
        "Front Axle": {
            "rpm": np.array([1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500]),
            "amp": np.array([2.5, 2.5, 2.5, 7.5, 7.5, 7.5, 7.5, 7.5])
        },
        "Rear Axle": {
            "rpm": np.array([1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500]),
            "amp": np.array([2.5, 2.5, 2.5, 7.5, 7.5, 7.5, 7.5, 7.5])
        }
    },
    "Gasoline": {
        "Front Axle": {
            "rpm": np.array([1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500]),
            "amp": np.array([2.5, 2.5, 2.5, 6.25, 10.0, 10.0, 10.0, 10.0])
        },
        "Rear Axle": {
            "rpm": np.array([1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500]),
            "amp": np.array([5.0, 5.0, 5.0, 10.0, 12.5, 12.5, 12.5, 12.5])
        }
    }
}


TRANSFER_CASE_TARGET_RPM = np.array(
    [1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500]
)

TRANSFER_CASE_ORDERS = {
    63.0: {
        "label": "63.00 Order - Gear Mesh",
        "target_rpm": TRANSFER_CASE_TARGET_RPM,
        "target_amp": np.array([5.0, 7.5, 10.0, 12.5, 15.0, 17.5, 20.0, 22.5]),
        "harmonic": "1st"
    },
    126.0: {
        "label": "126.00 Order - 2nd Harmonic",
        "target_rpm": None,
        "target_amp": None,
        "harmonic": "2nd"
    },
    85.05: {
        "label": "85.05 Order - Gear Mesh",
        "target_rpm": TRANSFER_CASE_TARGET_RPM,
        "target_amp": np.array([2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]),
        "harmonic": "1st"
    },
    170.10: {
        "label": "170.10 Order - 2nd Harmonic",
        "target_rpm": None,
        "target_amp": None,
        "harmonic": "2nd"
    }
}


def convert_csv_g_to_ms2_if_needed(headers, data):
    converted_channels = []

    for col_idx in [1, 2, 3]:
        header = str(headers[col_idx]).lower()

        if "(g)" in header or header.strip().endswith(" g"):
            data[:, col_idx] = data[:, col_idx] * G_TO_MS2
            converted_channels.append(headers[col_idx])

    return data, converted_channels


def load_measurement_file(uploaded_file):
    file_extension = uploaded_file.name.split(".")[-1].lower()

    if uploaded_file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
        st.error(f"File exceeds maximum allowed size: {MAX_FILE_SIZE_MB} MB.")
        st.stop()

    converted_channels = []

    if file_extension == "xlsx":
        temp_file = None

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.read())
                temp_file = tmp.name

            headers, data = read_xlsx_numeric(temp_file)

        finally:
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass

    elif file_extension == "csv":
        try:
            df = pd.read_csv(
                uploaded_file,
                sep=None,
                engine="python"
            )

            headers = list(df.columns)
            data = df.to_numpy(dtype=float)

            data, converted_channels = convert_csv_g_to_ms2_if_needed(
                headers,
                data
            )

        except Exception:
            st.error("CSV file could not be read. Please check delimiter and numeric data format.")
            st.code(traceback.format_exc())
            st.stop()

    else:
        st.error("Unsupported file format. Please upload .xlsx or .csv file.")
        st.stop()

    if data.ndim != 2 or data.shape[1] < 5:
        st.error("Measurement file must contain at least 5 columns: Time, ChA, ChB, ChC, RPM.")
        st.stop()

    if data.shape[0] > MAX_ROWS:
        st.error(f"Dataset exceeds maximum row limit: {MAX_ROWS} rows.")
        st.stop()

    if data.shape[0] < 10:
        st.error("Dataset is too short for order analysis.")
        st.stop()

    if not np.all(np.isfinite(data[:, :5])):
        st.error("Dataset contains NaN or non-numeric values in the first 5 columns.")
        st.stop()

    time = data[:, 0]
    rpm = data[:, 4]

    if np.any(np.diff(time) < 0):
        st.error("Time column contains decreasing values.")
        st.stop()

    if np.any(rpm <= 0):
        st.error("RPM column must contain only positive values.")
        st.stop()

    if converted_channels:
        st.info(
            "CSV detected: vibration channels converted from g to m/s²: "
            + ", ".join(map(str, converted_channels))
        )

    return headers, data


def format_comparison_sheet(writer, sheet_name):
    ws = writer.book[sheet_name]

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    info_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    status_col = None

    for cell in ws[1]:
        if cell.value == "Status":
            status_col = cell.column
            break

    if status_col is not None:
        for row in range(2, ws.max_row + 1):
            status_cell = ws.cell(row=row, column=status_col)

            if status_cell.value == "PASS":
                fill = green_fill
            elif status_cell.value == "FAIL":
                fill = red_fill
            elif status_cell.value == "INFO":
                fill = info_fill
            else:
                fill = None

            if fill is not None:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill

    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 22


def format_curve_sheet(writer, sheet_name):
    ws = writer.book[sheet_name]

    header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 16


def create_curve_plot_png(curve_df, order_label, vin_number, analysis_type, vehicle_config):
    fig, ax = plt.subplots(figsize=(14, 8))

    ax.plot(curve_df["RPM"], curve_df["ChA"], label="ChA", linewidth=2)
    ax.plot(curve_df["RPM"], curve_df["ChB"], label="ChB", linewidth=2)
    ax.plot(curve_df["RPM"], curve_df["ChC"], label="ChC", linewidth=2)

    if "Target" in curve_df.columns:
        ax.plot(
            curve_df["RPM"],
            curve_df["Target"],
            label="Target Curve",
            color="red",
            linewidth=5
        )

    ax.set_title(
        f"{order_label} vs RPM | VIN: {vin_number} | {analysis_type} | {vehicle_config}",
        fontsize=16
    )

    ax.set_xlabel("RPM", fontsize=13)
    ax.set_ylabel("Order Amplitude [m/s²]", fontsize=13)

    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper right", fontsize=12)

    rpm_min = min(1000, float(curve_df["RPM"].min()))
    rpm_max = max(4500, float(curve_df["RPM"].max()))
    ax.set_xlim(rpm_min, rpm_max)

    fig.tight_layout()

    img_buffer = BytesIO()
    fig.savefig(img_buffer, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    img_buffer.seek(0)
    return img_buffer


def add_png_plot_to_sheet(
    writer,
    sheet_name,
    curve_df,
    order_label,
    vin_number,
    analysis_type,
    vehicle_config
):
    ws = writer.book[sheet_name]

    img_buffer = create_curve_plot_png(
        curve_df=curve_df,
        order_label=order_label,
        vin_number=vin_number,
        analysis_type=analysis_type,
        vehicle_config=vehicle_config
    )

    img = XLImage(img_buffer)
    img.width = 900
    img.height = 520

    ws.add_image(img, "G2")


def make_excel_report(vehicle_info, results_by_order, curves_by_order, order_labels):
    output = BytesIO()

    vin_number = vehicle_info["VIN"]
    analysis_type = vehicle_info["Analysis Type"]
    vehicle_config = vehicle_info["Vehicle Config"]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([vehicle_info]).to_excel(
            writer,
            sheet_name="Vehicle Info",
            index=False
        )

        for order_value, result_df in results_by_order.items():
            sheet_name = f"{str(order_value).replace('.', '_')} Comparison"
            sheet_name = sheet_name[:31]

            result_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            format_comparison_sheet(writer, sheet_name)

        for order_value, curve_df in curves_by_order.items():
            sheet_name = f"{str(order_value).replace('.', '_')} Curves"
            sheet_name = sheet_name[:31]

            curve_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            format_curve_sheet(writer, sheet_name)

            add_png_plot_to_sheet(
                writer=writer,
                sheet_name=sheet_name,
                curve_df=curve_df,
                order_label=order_labels[order_value],
                vin_number=vin_number,
                analysis_type=analysis_type,
                vehicle_config=vehicle_config
            )

    output.seek(0)
    return output


def analyze_order(
    order_value,
    order_label,
    harmonic_label,
    time,
    rpm,
    channels,
    samples_per_rev,
    revs_per_block,
    overlap,
    max_order,
    order_width,
    rpm_step,
    cal_factor,
    target_rpm=None,
    target_amp=None
):
    channel_curves = {}
    peak_results = []

    has_target = target_rpm is not None and target_amp is not None

    for name, sig in channels.items():

        theta_u, x_u, rpm_u = angular_resample(
            time,
            rpm,
            sig,
            samples_per_rev=samples_per_rev
        )

        orders, rpms, spec = order_map(
            theta_u,
            x_u,
            rpm_u,
            samples_per_rev=samples_per_rev,
            revs_per_block=revs_per_block,
            overlap=overlap,
            max_order=max_order
        )

        rpm_sorted, amp_sorted = extract_order_vs_rpm(
            orders,
            rpms,
            spec,
            target_order=order_value,
            width=order_width,
            rpm_step=rpm_step,
            smooth=True
        )

        amp_sorted = amp_sorted * cal_factor

        channel_curves[name] = {
            "rpm": rpm_sorted,
            "amp": amp_sorted
        }

        peak_idx = np.argmax(amp_sorted)
        peak_rpm = float(rpm_sorted[peak_idx])
        peak_amp = float(amp_sorted[peak_idx])

        if has_target:
            target_curve = np.interp(
                rpm_sorted,
                target_rpm,
                target_amp
            )

            target_at_peak = float(
                np.interp(
                    peak_rpm,
                    target_rpm,
                    target_amp
                )
            )

            margin_curve = amp_sorted - target_curve
            exceedance = np.maximum(margin_curve, 0.0)

            exceedance_area = float(
                np.trapz(
                    exceedance,
                    rpm_sorted
                )
            )

            max_margin = float(np.max(margin_curve))
            max_margin_percent = (
                max_margin / target_at_peak * 100.0
                if target_at_peak > 0
                else np.nan
            )

            status = "PASS" if exceedance_area <= 1e-9 else "FAIL"

        else:
            target_at_peak = np.nan
            max_margin = np.nan
            max_margin_percent = np.nan
            exceedance_area = np.nan
            status = "INFO"

        peak_results.append({
            "Order": order_value,
            "Order Label": order_label,
            "Harmonic": harmonic_label,
            "Channel": name,
            "Peak RPM": peak_rpm,
            "Peak Amplitude [m/s²]": peak_amp,
            "Target at Peak RPM [m/s²]": target_at_peak,
            "Max Margin [m/s²]": max_margin,
            "Max Margin [%]": max_margin_percent,
            "Exceedance Area [m/s²·RPM]": exceedance_area,
            "Status": status
        })

    result_df = pd.DataFrame(peak_results)

    curve_df = pd.DataFrame()
    base_rpm = None

    for name, curve in channel_curves.items():
        if base_rpm is None:
            base_rpm = curve["rpm"]
            curve_df["RPM"] = base_rpm

        curve_df[name] = np.interp(
            base_rpm,
            curve["rpm"],
            curve["amp"]
        )

    if has_target:
        curve_df["Target"] = np.interp(
            curve_df["RPM"],
            target_rpm,
            target_amp
        )

    return channel_curves, result_df, curve_df


def plot_order_comparison(
    order_label,
    channel_curves,
    target_rpm,
    target_amp,
    vin_number,
    analysis_type,
    vehicle_config
):
    fig, ax = plt.subplots(figsize=(12, 7))

    for name, curve in channel_curves.items():
        ax.plot(
            curve["rpm"],
            curve["amp"],
            label=name
        )

    if target_rpm is not None and target_amp is not None:
        ax.plot(
            target_rpm,
            target_amp,
            color="red",
            linewidth=4,
            label="Target Curve"
        )

    ax.set_xlabel("RPM")
    ax.set_ylabel("Order Amplitude [m/s²]")
    ax.set_title(
        f"{order_label} vs RPM | VIN: {vin_number} | {analysis_type} | {vehicle_config}"
    )

    ax.grid(True, alpha=0.3)
    ax.legend()

    return fig


st.subheader("Vehicle Information")

col1, col2, col3, col4 = st.columns(4)

with col1:
    vin_number = st.text_input(
        "VIN Number",
        placeholder="Enter 17-character VIN",
        max_chars=17
    ).upper().strip()

vin_valid = bool(re.fullmatch(r"[A-Z0-9]{17}", vin_number))

with col2:
    analysis_type = st.selectbox(
        "Analysis Type",
        [
            "Axle Whine Order Analysis",
            "Transfer Case Gear Mesh Analysis"
        ],
        disabled=not vin_valid
    )

if analysis_type == "Axle Whine Order Analysis":
    with col3:
        fuel_type = st.selectbox(
            "Fuel Type",
            ["Select fuel type", "Diesel", "Gasoline"],
            disabled=not vin_valid
        )

    with col4:
        axle_type = st.selectbox(
            "Axle Type",
            ["Select axle type", "Front Axle", "Rear Axle"],
            disabled=not vin_valid
        )

else:
    fuel_type = "N/A"
    axle_type = "Transfer Case / 6th Gear"

    with col3:
        st.text_input(
            "Gear",
            value="6th Gear",
            disabled=True
        )

    with col4:
        st.text_input(
            "Component",
            value="Transfer Case",
            disabled=True
        )

if vin_number and not vin_valid:
    st.error("VIN must be exactly 17 characters and contain only letters and numbers.")


st.subheader("Measurement Data")

uploaded_file = st.file_uploader(
    "Upload Measurement File",
    type=["xlsx", "csv"],
    disabled=not vin_valid,
    help="Supported formats: .xlsx and .csv"
)


if analysis_type == "Axle Whine Order Analysis":
    can_continue = (
        vin_valid
        and fuel_type != "Select fuel type"
        and axle_type != "Select axle type"
        and uploaded_file is not None
    )
else:
    can_continue = (
        vin_valid
        and uploaded_file is not None
    )

if not can_continue:
    if not vin_valid:
        st.warning("Please enter a valid 17-character VIN before selecting analysis type and uploading data.")
    elif analysis_type == "Axle Whine Order Analysis":
        st.warning("Please select fuel type, select axle type, and upload measurement file.")
    else:
        st.warning("Please upload measurement file.")
    st.stop()


if analysis_type == "Axle Whine Order Analysis":
    base_target = AXLE_TARGETS[fuel_type][axle_type]

    ANALYSIS_ORDERS = {
        10.0: {
            "label": "10th Order",
            "target_rpm": base_target["rpm"],
            "target_amp": base_target["amp"],
            "harmonic": "Base"
        },
        20.0: {
            "label": "20th Order",
            "target_rpm": base_target["rpm"],
            "target_amp": base_target["amp"],
            "harmonic": "2nd"
        }
    }

    default_max_order = 30
    vehicle_config = f"{fuel_type} | {axle_type}"

else:
    ANALYSIS_ORDERS = TRANSFER_CASE_ORDERS
    default_max_order = 200
    vehicle_config = "Transfer Case | 6th Gear"


st.success("Vehicle information and measurement file are ready for analysis.")

info_cols = st.columns(4)
info_cols[0].metric("VIN", vin_number)
info_cols[1].metric("Analysis", analysis_type)
info_cols[2].metric("Fuel Type", fuel_type)
info_cols[3].metric("Configuration", axle_type)


st.subheader("Analysis Settings")

samples_per_rev = 512
revs_per_block = 8
overlap = 0.75
rpm_step = 10
cal_factor = 1.0

with st.expander("Advanced Settings", expanded=False):

    selected_channel = st.selectbox(
        "Order Map Channel",
        ["ChA", "ChB", "ChC"]
    )

    max_order = st.slider(
        "Max order",
        5,
        250,
        default_max_order
    )

    order_width = st.number_input(
        "Order width",
        min_value=0.05,
        max_value=2.0,
        value=0.15,
        step=0.05
    )


if st.button("Run Analysis", type="primary"):

    try:
        headers, data = load_measurement_file(uploaded_file)

        time = data[:, 0]
        rpm = data[:, 4]

        channels = {
            "ChA": data[:, 1],
            "ChB": data[:, 2],
            "ChC": data[:, 3],
        }

        with st.spinner("Analysis is running..."):

            curves_by_order = {}
            results_by_order = {}
            raw_curves_by_order = {}
            order_labels = {}

            for order_value, order_info in ANALYSIS_ORDERS.items():

                channel_curves, result_df, curve_df = analyze_order(
                    order_value=order_value,
                    order_label=order_info["label"],
                    harmonic_label=order_info["harmonic"],
                    time=time,
                    rpm=rpm,
                    channels=channels,
                    samples_per_rev=samples_per_rev,
                    revs_per_block=revs_per_block,
                    overlap=overlap,
                    max_order=max_order,
                    order_width=order_width,
                    rpm_step=rpm_step,
                    cal_factor=cal_factor,
                    target_rpm=order_info["target_rpm"],
                    target_amp=order_info["target_amp"]
                )

                curves_by_order[order_value] = channel_curves
                results_by_order[order_value] = result_df
                raw_curves_by_order[order_value] = curve_df
                order_labels[order_value] = order_info["label"]

            overall_status = "PASS"

            for result_df in results_by_order.values():
                target_rows = result_df[result_df["Status"] != "INFO"]

                if len(target_rows) > 0 and not (target_rows["Status"] == "PASS").all():
                    overall_status = "FAIL"

            st.subheader("Overall Assessment")

            if overall_status == "PASS":
                st.success("Overall Assessment: PASS")
            else:
                st.error("Overall Assessment: FAIL")

            vehicle_info = {
                "VIN": vin_number,
                "Analysis Type": analysis_type,
                "Fuel Type": fuel_type,
                "Vehicle Config": vehicle_config,
                "Target Orders": ", ".join([str(o) for o in ANALYSIS_ORDERS.keys()]),
                "Order Width": order_width,
                "RPM Step": rpm_step,
                "Samples per Rev": samples_per_rev,
                "Revs per Block": revs_per_block,
                "Overlap": overlap,
                "Calibration Factor": cal_factor,
                "Max Order": max_order,
                "Overall Assessment": overall_status
            }

            excel_report = make_excel_report(
                vehicle_info,
                results_by_order,
                raw_curves_by_order,
                order_labels
            )

            st.download_button(
                label="📊 Download Excel Report",
                data=excel_report,
                file_name=f"{vin_number}_{analysis_type.replace(' ', '_')}_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            if analysis_type == "Axle Whine Order Analysis":
                tab1, tab2, tab3, tab4 = st.tabs(
                    [
                        "10th Order Target Comparison",
                        "20th Order Target Comparison",
                        "Order Map / Waterfall",
                        "Raw Results"
                    ]
                )

                result_tabs = [tab1, tab2]

                for tab, order_value in zip(result_tabs, ANALYSIS_ORDERS.keys()):

                    with tab:
                        order_info = ANALYSIS_ORDERS[order_value]
                        result_df = results_by_order[order_value]
                        channel_curves = curves_by_order[order_value]

                        order_status = (
                            "PASS"
                            if (result_df["Status"] == "PASS").all()
                            else "FAIL"
                        )

                        st.subheader(f"{order_info['label']} Result Summary")

                        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

                        kpi1.metric(
                            "Peak ChA",
                            f"{result_df.loc[result_df['Channel'] == 'ChA', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi2.metric(
                            "Peak ChB",
                            f"{result_df.loc[result_df['Channel'] == 'ChB', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi3.metric(
                            "Peak ChC",
                            f"{result_df.loc[result_df['Channel'] == 'ChC', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi4.metric(
                            f"{order_info['label']} Assessment",
                            order_status
                        )

                        fig_cmp = plot_order_comparison(
                            order_label=order_info["label"],
                            channel_curves=channel_curves,
                            target_rpm=order_info["target_rpm"],
                            target_amp=order_info["target_amp"],
                            vin_number=vin_number,
                            analysis_type=analysis_type,
                            vehicle_config=vehicle_config
                        )

                        st.pyplot(fig_cmp)

                        st.subheader(f"{order_info['label']} Target Compliance")

                        st.dataframe(
                            result_df,
                            use_container_width=True
                        )

                order_map_tab = tab3
                raw_tab = tab4

            else:
                tab1, tab2, tab3 = st.tabs(
                    [
                        "Gear Mesh Order Results",
                        "Order Map / Waterfall",
                        "Raw Results"
                    ]
                )

                with tab1:
                    for order_value, order_info in ANALYSIS_ORDERS.items():
                        result_df = results_by_order[order_value]
                        channel_curves = curves_by_order[order_value]

                        if (result_df["Status"] == "INFO").all():
                            order_status = "INFO"
                        else:
                            order_status = (
                                "PASS"
                                if (result_df[result_df["Status"] != "INFO"]["Status"] == "PASS").all()
                                else "FAIL"
                            )

                        st.subheader(f"{order_info['label']} Result Summary")

                        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

                        kpi1.metric(
                            "Peak ChA",
                            f"{result_df.loc[result_df['Channel'] == 'ChA', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi2.metric(
                            "Peak ChB",
                            f"{result_df.loc[result_df['Channel'] == 'ChB', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi3.metric(
                            "Peak ChC",
                            f"{result_df.loc[result_df['Channel'] == 'ChC', 'Peak Amplitude [m/s²]'].iloc[0]:.2f} m/s²"
                        )

                        kpi4.metric(
                            "Assessment",
                            order_status
                        )

                        fig_cmp = plot_order_comparison(
                            order_label=order_info["label"],
                            channel_curves=channel_curves,
                            target_rpm=order_info["target_rpm"],
                            target_amp=order_info["target_amp"],
                            vin_number=vin_number,
                            analysis_type=analysis_type,
                            vehicle_config=vehicle_config
                        )

                        st.pyplot(fig_cmp)

                        st.subheader(f"{order_info['label']} Compliance / Severity Table")

                        st.dataframe(
                            result_df,
                            use_container_width=True
                        )

                        st.markdown("---")

                order_map_tab = tab2
                raw_tab = tab3

            with order_map_tab:

                st.subheader(f"Order Map / Waterfall - {selected_channel}")

                sig = channels[selected_channel]

                theta_u, x_u, rpm_u = angular_resample(
                    time,
                    rpm,
                    sig,
                    samples_per_rev=samples_per_rev
                )

                orders, rpms, spec = order_map(
                    theta_u,
                    x_u,
                    rpm_u,
                    samples_per_rev=samples_per_rev,
                    revs_per_block=revs_per_block,
                    overlap=overlap,
                    max_order=max_order
                )

                idx = np.argsort(rpms)
                r = rpms[idx]
                s = spec[idx]

                db = 20 * np.log10(np.maximum(s * cal_factor, 1e-12))

                fig, ax = plt.subplots(figsize=(12, 7))

                im = ax.imshow(
                    db,
                    aspect="auto",
                    origin="lower",
                    extent=[orders[0], orders[-1], r[0], r[-1]],
                    interpolation="nearest",
                    cmap="jet"
                )

                fig.colorbar(
                    im,
                    ax=ax,
                    label="Amplitude [dB re 1 m/s²]"
                )

                ax.set_xlabel("Order")
                ax.set_ylabel("RPM")
                ax.set_title(
                    f"Order Map / Waterfall - {selected_channel} | VIN: {vin_number} | {analysis_type}"
                )

                st.pyplot(fig)

            with raw_tab:

                for order_value, curve_df in raw_curves_by_order.items():
                    st.subheader(f"{order_labels[order_value]} Raw Curve Data")

                    st.dataframe(
                        curve_df,
                        use_container_width=True
                    )

    except Exception:
        st.error("Uygulama çalışırken hata oluştu")
        st.code(traceback.format_exc())
