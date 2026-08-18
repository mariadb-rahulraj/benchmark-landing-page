import io
import json
import os
import requests
import streamlit as st
import pandas as pd
import plotly.express as px
import base64

# ==============================================================================
# 1. GLOBAL CONFIGURATION & LINKS
# ==============================================================================
st.set_page_config(page_title="Benchmark Analytics Portal", layout="wide")

CONFIG_FILE_PATH = "sheets_config.json"

def load_links_config(json_path="links.json"):
    """Loads external link configurations from a JSON file with fallback defaults."""
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return (
                    data.get("SYSBENCH_LINKS", {}), 
                    data.get("HAMMERDB_LINKS", {}), 
                    data.get("MACHINE_DETAILS", {})
                )
        except Exception as e:
            print(f"Error loading {json_path}: {e}")
    
    # Fallback default links if file is missing or unreadable
    return {}, {}, {}

# Load links dynamically from JSON
SYSBENCH_LINKS, HAMMERDB_LINKS, MACHINE_DETAILS = load_links_config()

def load_quarterly_sheet_urls(selected_edition: str, selected_tool: str, config_path: str = CONFIG_FILE_PATH) -> dict:
    """Loads Google Sheet URLs matching the active Edition AND Benchmark Tool."""
    if not os.path.exists(config_path):
        st.error(f"Configuration file `{config_path}` not found! Please create it in the app directory.")
        return {}
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            all_configs = json.load(f)
            
        # Drill down into selected edition -> selected benchmark tool
        return all_configs.get(selected_edition, {}).get(selected_tool, {})
    except Exception as e:
        st.error(f"Error reading `{config_path}`: {e}")
        return {}


# ==============================================================================
# 2. DATA LOADING & PARSING FUNCTIONS
# ==============================================================================
@st.cache_data(ttl=300)
def get_google_sheet_bytes(url: str) -> io.BytesIO:
    export_url = url.split("/edit")[0] + "/export?format=xlsx"
    response = requests.get(export_url, timeout=15)
    response.raise_for_status()
    return io.BytesIO(response.content)


@st.cache_data(ttl=300, show_spinner=False)
def parse_stacked_matrix_sheet(sheet_urls_dict: dict) -> pd.DataFrame:
    """Scans all registered quarterly sheets, preserves tab order and names, and flattens data."""
    all_flattened_data = []

    for quarter_name, url in sheet_urls_dict.items():
        if not url or "YOUR_" in url:
            continue

        try:
            excel_bytes = get_google_sheet_bytes(url)
            excel_file = pd.ExcelFile(excel_bytes, engine='openpyxl')

            for sheet_name in excel_file.sheet_names:
                raw_sheet_name = str(sheet_name).strip()

                if "_" in raw_sheet_name:
                    parts = raw_sheet_name.split("_", 1)
                    tab_config = parts[1].strip()
                else:
                    tab_config = raw_sheet_name

                # Rename configurations to meaningful names
                CONFIG_RENAME_MAP = {
                    "Default": "Default",
                    "AHI": "Adaptive Hash Index",
                    "TRX_1": "innodb_flush_log_at_trx_commit=1",
                    "TRX_1_DW_1": "innodb_flush_log_at_trx_commit=1, double_write=1",
                    "Binlog": "Binlog Enabled",
                    "Serial": "transaction_isolation=SERIALIZABLE",
                    "UAW_TRX_1_DW_1": "UAW innodb_flush_log_at_trx_commit=1, double_write=1",
                    "UAW_Binlog": "UAW_Binlog Enabled"
                }

                tab_config = CONFIG_RENAME_MAP.get(tab_config, tab_config)

                df_tab_raw = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None, engine='openpyxl')
                if df_tab_raw.empty:
                    continue

                block_indices = []
                for r_idx, val in enumerate(df_tab_raw.iloc[:, 0]):
                    val_str = str(val).strip().upper()
                    if val_str in ["AMD", "INTEL"]:
                        block_indices.append((r_idx, val_str))

                if not block_indices:
                    block_indices = [(0, "AMD")]

                for idx_pos, (start_row, processor) in enumerate(block_indices):
                    end_row = block_indices[idx_pos + 1][0] if idx_pos + 1 < len(block_indices) else len(df_tab_raw)
                    df_block = df_tab_raw.iloc[start_row:end_row].reset_index(drop=True)

                    if len(df_block) < 4:
                        continue

                    builds_row = df_block.iloc[1]
                    data_rows = df_block.iloc[3:]

                    for col_i in range(1, len(df_block.columns), 2):
                        build_name = str(builds_row[col_i]).strip()
                        if build_name.lower() in ["nan", "none", ""]:
                            continue

                        vu_col = pd.to_numeric(data_rows.iloc[:, 0], errors='coerce')
                        nopm_col = pd.to_numeric(data_rows.iloc[:, col_i], errors='coerce')
                        tpm_col = pd.to_numeric(data_rows.iloc[:, col_i + 1], errors='coerce')

                        sub_df = pd.DataFrame({
                            'Quarter': quarter_name,
                            'Configuration': tab_config,
                            'Processor': processor,
                            'Build': build_name,
                            'VU': vu_col,
                            'NOPM': nopm_col,
                            'TPM': tpm_col
                        })

                        sub_df = sub_df.dropna(subset=['VU', 'NOPM'])
                        all_flattened_data.append(sub_df)

        except Exception as e:
            st.error(f"Error parsing Google Sheet for {quarter_name}: {e}")

    if not all_flattened_data:
        return pd.DataFrame()

    combined = pd.concat(all_flattened_data, ignore_index=True)
    combined['VU'] = pd.to_numeric(combined['VU'], errors='coerce')
    combined['NOPM'] = pd.to_numeric(combined['NOPM'], errors='coerce')
    combined['TPM'] = pd.to_numeric(combined['TPM'], errors='coerce')
    return combined.dropna(subset=['VU', 'NOPM', 'TPM'])

def get_base64_logo(image_path: str) -> str:
    """Reads a local image file and converts it into a Base64 Data URI string."""
    if os.path.exists(image_path):
        with open(image_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            return f"data:image/png;base64,{encoded}"
    return ""

# ==============================================================================
# 3. UI RENDERING MODULES
# ==============================================================================
def render_landing_page():
    """Renders the primary landing page with Edition and Benchmark Tool selectors."""
    
    # CSS injection to make Sysbench & HammerDB containers equal in height
    st.markdown(
        """
        <style>
            div[data-testid="column"] > div[data-testid="stVerticalBlock"] {
               height: 100%;
            }
            div[data-testid="column"] div[data-testid="stVerticalBlockBorderWrapper"] {
                height: 100% !important;
                min-height: 280px !important;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<h1 style='text-align: center;'>🚀 Performance Analytics Portal</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Select a benchmark module or jump directly to the live dashboard.</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Load local HammerDB icon as base64 (Change 'hammerdb.png' to your actual file name if different)
    hammerdb_img_src = get_base64_logo("hammerdb_logo.png")

    col1, col2, col3 = st.columns(3)

    with col1:
        with st.container(border=True):
            st.markdown(
                """
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 5px;">
                    <span style="font-size: 28px;">📊</span>
                    <h3 style="margin: 0; padding: 0;">Quarter Performance</h3>
                </div>
                """, 
                unsafe_allow_html=True
            )
            st.caption("Live interactive analytics dashboard.")
            
            # Side-by-side radio options for Edition and Benchmark Tool
            sel_col1, sel_col2 = st.columns(2)
            
            with sel_col1:
                server_type = st.radio(
                    "Select Edition:",
                    options=["Enterprise Server", "Community Server"],
                    key="selected_server_edition"
                )

            with sel_col2:
                benchmark_tool = st.radio(
                    "Benchmark Tool:",
                    options=["HammerDB", "Sysbench"],
                    key="selected_benchmark_tool"
                )
            
            st.write("")
            
            if st.button("📈 Launch Dashboard", use_container_width=True, type="primary"):
                st.session_state["server_edition"] = server_type
                st.session_state["benchmark_tool"] = benchmark_tool
                st.session_state["view_dashboard"] = True
                st.rerun()

    with col2:
        with st.container(border=True):
            if hammerdb_img_src:
                hammer_icon_html = f'<img src="{hammerdb_img_src}" width="30" height="30" style="object-fit: contain;">'
            else:
                hammer_icon_html = '<span style="font-size: 28px;">🔨</span>'

            st.markdown(
                f"""
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 15px;">
                    {hammer_icon_html}
                    <h3 style="margin: 0; padding: 0;">HammerDB</h3>
                </div>
                """, 
                unsafe_allow_html=True
            )
            st.caption("Quarterly HammerDB raw sheets and Daily Run")
            with st.expander("🔗 View HammerDB Sheets and Runs"):
                for label, url in HAMMERDB_LINKS.items():
                    st.link_button(f"📄 {label}", url, use_container_width=True)

    with col3:
        with st.container(border=True):
            st.markdown(
                """
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 5px;">
                    <span style="font-size: 28px;">⚡</span>
                    <h3 style="margin: 0; padding: 0;">Sysbench</h3>
                </div>
                """, 
                unsafe_allow_html=True
            )
            st.caption("Quarterly Sysbench raw sheets")
            with st.expander("🔗 View Sysbench Sheets"):
                for label, url in SYSBENCH_LINKS.items():
                    st.link_button(f"📄 {label}", url, use_container_width=True)

def render_peak_trends_tab(df_proc: pd.DataFrame, processor_name: str, selected_configs: list, selected_builds: list = None):
    """
    Renders Tab 1: Peak Trends in a single self-contained function.
    Includes:
    - Auto-synced Quarter filter when sidebar builds/configs change.
    - Non-overlapping line chart labels with directional padding and version prefixes.
    - Y-axis headroom and inward edge label alignment to prevent right-edge clipping.
    - Center-aligned table headers/values with major version demarcation lines (e.g. 10.6, 11.4, 11.8).
    """
    st.subheader(f"Quarter-over-Quarter Peak Performance ({processor_name})")

    # --------------------------------------------------------------------------
    # 0. INJECT CSS TO OVERRIDE STREAMLIT TABLE HEADER & CELL ALIGNMENT + PRINT PAGE BREAKS
    # --------------------------------------------------------------------------
    st.markdown(
        """
        <style>
            div[data-testid="stMarkdownContainer"] table th:nth-child(3),
            div[data-testid="stMarkdownContainer"] table th:nth-child(4) {
                text-align: center !important;
            }
            div[data-testid="stMarkdownContainer"] table td:nth-child(3),
            div[data-testid="stMarkdownContainer"] table td:nth-child(4) {
                text-align: center !important;
            }
            .version-divider {
                border-top: 2px solid #3b82f6 !important;
            }

            /* FORCE PAGE BREAK FOR EVERY CONFIGURATION WHEN PRINTING / SAVING TO PDF */
            @media print {
                hr {
                    page-break-after: always !important;
                    break-after: page !important;
                    visibility: hidden !important;
                }
                div[data-testid="stPlotlyChart"],
                div[data-testid="stMarkdownContainer"] {
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                }
            }
        </style>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------------------------
    # 1. FILTER DATA BY SIDEBAR BUILDS & CONFIGS
    # --------------------------------------------------------------------------
    df_active = df_proc.copy()
    if selected_builds:
        df_active = df_active[df_active['Build'].isin(selected_builds)]

    df_filtered = df_active[df_active['Configuration'].isin(selected_configs)] if selected_configs else df_active
    relevant_quarters = list(dict.fromkeys(df_filtered['Quarter']))

    if not relevant_quarters:
        st.warning("No data available for the selected build and configuration filters.")
        return

    # --------------------------------------------------------------------------
    # 2. FORCE SESSION STATE SYNC FOR MULTISELECT
    # --------------------------------------------------------------------------
    multiselect_key = f"peak_q_{processor_name}"
    tracker_key = f"filter_tracker_{processor_name}"
    
    current_filters = (tuple(selected_configs or []), tuple(selected_builds or []))

    # Initialize key if missing or filter builds changed
    if tracker_key not in st.session_state or st.session_state[tracker_key] != current_filters:
        st.session_state[tracker_key] = current_filters
        st.session_state[multiselect_key] = relevant_quarters
    elif multiselect_key not in st.session_state or not st.session_state[multiselect_key]:
        # Keeps quarters selected when toggling back from Landing Page
        st.session_state[multiselect_key] = [q for q in relevant_quarters if q in relevant_quarters]

    # --------------------------------------------------------------------------
    # 3. CONTROLS LAYOUT
    # --------------------------------------------------------------------------
    ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1, 1, 2])

    with ctrl_col1:
        metric_choice = st.radio(
            f"Select Metric ({processor_name}):", ["NOPM", "TPM"], horizontal=True, key=f"metric_{processor_name}"
        )

    with ctrl_col2:
        chart_type = st.radio(
            "Graph Type:", ["Bar Chart", "Line Chart"], index=1, horizontal=True, key=f"chart_type_{processor_name}"
        )

    with ctrl_col3:
        selected_quarters = st.multiselect(
            "Filter Quarters:", 
            options=relevant_quarters, 
            key=multiselect_key
        )

    if not selected_quarters:
        st.warning("Please select at least one quarter to display peak trends.")
        return

    # --------------------------------------------------------------------------
    # 4. CONFIGURATION LOOP & RENDERING
    # --------------------------------------------------------------------------
    for cfg in selected_configs:
        df_cfg = df_active[
            (df_active['Configuration'] == cfg) & 
            (df_active['Quarter'].isin(selected_quarters))
        ]
        if df_cfg.empty:
            continue

        st.markdown(f"### ⚙️ Configuration: `{cfg}`")

        peak_df = df_cfg.groupby(['Quarter', 'Build'], as_index=False).agg({
            'NOPM': 'max', 'TPM': 'max'
        })

        # Extract major version family
        peak_df['Build_Family'] = peak_df['Build'].astype(str).str.extract(r'(\d+\.\d+)')[0].fillna(peak_df['Build'])

        # Order Quarters chronologically
        quarter_order = [q for q in relevant_quarters if q in selected_quarters]
        peak_df['Quarter_Cat'] = pd.Categorical(peak_df['Quarter'], categories=quarter_order, ordered=True)
        
        peak_df = peak_df.sort_values(by=['Build_Family', 'Quarter_Cat']).reset_index(drop=True)

        # Global Baseline calculation
        global_base = peak_df[metric_choice].iloc[0] if not peak_df.empty else 1
        peak_df[f'{metric_choice} Ratio'] = peak_df[metric_choice] / global_base

        # Calculate Y-axis headroom
        max_val = peak_df[metric_choice].max() if not peak_df.empty else 100
        y_max = max_val * 1.25  # 25% upper headroom for top labels

        # Graph Rendering
        if chart_type == "Bar Chart":
            fig_peak = px.bar(
                peak_df, x="Quarter", y=metric_choice, color="Build_Family", hover_data=["Build"], barmode="group",
                text_auto='.2s', height=450, title=f"Peak {metric_choice} Progression — {cfg}",
                labels={"Quarter": "Quarter", metric_choice: f"Highest {metric_choice} Value", "Build_Family": "Build Version"}
            )
        else:
            fig_peak = px.line(
                peak_df, x="Quarter", y=metric_choice, color="Build_Family", hover_data=["Build"],
                markers=True, height=450, title=f"Peak {metric_choice} Progression (Line View) — {cfg}",
                labels={"Quarter": "Quarter", metric_choice: f"Highest {metric_choice} Value", "Build_Family": "Build Version"}
            )

            unique_fams = list(dict.fromkeys(peak_df['Build_Family']))

            for i, fam in enumerate(unique_fams):
                fam_df = peak_df[peak_df['Build_Family'] == fam].reset_index(drop=True)
                n_points = len(fam_df)

                # Force 10.6 below the line, alternate others
                if str(fam).strip() == "10.6":
                    is_top = False
                else:
                    is_top = (i % 2 == 0)
                
                text_labels = []
                text_positions = []

                for p_idx, row in fam_df.iterrows():
                    ratio = row[f'{metric_choice} Ratio']
                    global_row_idx = peak_df[peak_df['Build'] == row['Build']].index[0]

                    if global_row_idx == 0:
                        val_str = "Baseline"
                    elif ratio > 1.0:
                        val_str = f"{ratio:.2f}x Gain"
                    elif ratio < 1.0:
                        val_str = f"{ratio:.2f}x Drop"
                    else:
                        val_str = "1.00x"

                    label_text = f"<b>{fam}:</b> {val_str}"

                    # Add linebreaks (<br>) to offset text above or below line strokes
                    if is_top:
                        text_labels.append(f"{label_text}<br>&nbsp;")
                    else:
                        text_labels.append(f"&nbsp;<br>{label_text}")

                    # Force inward alignment for edge points
                    if p_idx == n_points - 1 and n_points > 1:
                        text_positions.append("top left" if is_top else "bottom left")
                    elif p_idx == 0 and n_points > 1:
                        text_positions.append("top right" if is_top else "bottom right")
                    else:
                        text_positions.append("top center" if is_top else "bottom center")

                fig_peak.for_each_trace(
                    lambda trace, fam_name=fam, lbls=text_labels, pos=text_positions: trace.update(
                        mode="lines+markers+text",
                        text=lbls,
                        textposition=pos,
                        cliponaxis=False,
                        marker=dict(size=8),
                        line=dict(width=3)
                    ) if trace.name == fam_name else None
                )

        fig_peak.update_layout(
            hovermode="x unified",
            margin=dict(r=100, t=60, l=10, b=10),
            xaxis=dict(type='category', categoryorder='array', categoryarray=quarter_order),
            yaxis=dict(rangemode="tozero", range=[0, y_max])
        )
        st.plotly_chart(fig_peak, use_container_width=True)

        # ----------------------------------------------------------------------
        # 5. SUMMARY TABLE RENDERING WITH VERSION DEMARCATIONS
        # ----------------------------------------------------------------------
        gain_summary = peak_df[['Quarter', 'Build', 'Build_Family', metric_choice, f'{metric_choice} Ratio']].copy()

        table_rows = []
        for idx, row in gain_summary.iterrows():
            ratio = row[f'{metric_choice} Ratio']
            
            if idx == 0:
                gain_str = "Baseline"
            elif pd.isnull(ratio):
                gain_str = "N/A"
            elif ratio > 1.0:
                gain_str = f"{ratio:.2f}x Gain"
            elif ratio < 1.0:
                gain_str = f"{ratio:.2f}x Drop"
            else:
                gain_str = "1.00x Same"

            val_formatted = f"{int(row[metric_choice]):,}" if pd.notnull(row[metric_choice]) else "N/A"

            is_new_family = (idx > 0) and (row['Build_Family'] != gain_summary.at[idx - 1, 'Build_Family'])
            divider_class = "class='version-divider'" if is_new_family else ""

            table_rows.append(
                f"<tr {divider_class} style='border-bottom: 1px solid rgba(128, 128, 128, 0.2);'>"
                f"<td style='text-align: left; padding: 10px 12px;'>{row['Quarter']}</td>"
                f"<td style='text-align: left; padding: 10px 12px;'>{row['Build']}</td>"
                f"<td style='padding: 10px 12px;'>{val_formatted}</td>"
                f"<td style='padding: 10px 12px;'>{gain_str}</td>"
                f"</tr>"
            )

        html_table = f"""
        <div style="border: 1px solid rgba(128, 128, 128, 0.35); border-radius: 6px; overflow: hidden; margin: 12px 0 20px 0;">
            <table style="width: 100%; border-collapse: collapse; font-family: inherit; margin: 0;">
                <thead>
                    <tr style="background-color: rgba(128, 128, 128, 0.12); border-bottom: 2px solid rgba(128, 128, 128, 0.35);">
                        <th style="text-align: left; padding: 10px 12px; font-weight: 600;">Quarter</th>
                        <th style="text-align: left; padding: 10px 12px; font-weight: 600;">Build</th>
                        <th style="padding: 10px 12px; font-weight: 600;">{metric_choice}</th>
                        <th style="padding: 10px 12px; font-weight: 600;">Performance Gain / Drop</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(table_rows)}
                </tbody>
            </table>
        </div>
        """

        st.markdown(html_table, unsafe_allow_html=True)
        st.markdown("---")

def render_vu_scaling_tab(df_proc: pd.DataFrame, processor_name: str, selected_configs: list, selected_builds: list = None):
    """Renders Tab 2: Dedicated VU Scaling Graphs with metric toggle and ordered legend traces."""
    st.subheader(f"VU Load Scaling Curves ({processor_name})")

    # --------------------------------------------------------------------------
    # 1. FILTER DATA BY SELECTED BUILDS & EXTRACT RELEVANT QUARTERS
    # --------------------------------------------------------------------------
    df_active = df_proc.copy()
    if selected_builds:
        df_active = df_active[df_active['Build'].isin(selected_builds)]

    df_filtered = df_active[df_active['Configuration'].isin(selected_configs)] if selected_configs else df_active
    quarters = list(dict.fromkeys(df_filtered['Quarter'])) if not df_filtered.empty else list(dict.fromkeys(df_active['Quarter']))

    # --------------------------------------------------------------------------
    # 2. FORCE SESSION STATE SYNC FOR AUTOMATIC QUARTER SELECTION
    # --------------------------------------------------------------------------
    multiselect_key = f"q_{processor_name}"
    tracker_key = f"vu_tracker_{processor_name}"
    current_filters = (tuple(selected_configs or []), tuple(selected_builds or []))

    if tracker_key not in st.session_state or st.session_state[tracker_key] != current_filters:
        st.session_state[tracker_key] = current_filters
        st.session_state[multiselect_key] = quarters
    elif multiselect_key not in st.session_state or not st.session_state[multiselect_key]:
        st.session_state[multiselect_key] = quarters

    # --------------------------------------------------------------------------
    # 3. CONTROLS
    # --------------------------------------------------------------------------
    metric_choice = st.radio(
        f"Select Scaling Metric ({processor_name}):", ["NOPM", "TPM"], horizontal=True, key=f"vu_metric_{processor_name}"
    )

    selected_quarters = st.multiselect(
        f"Select Quarter(s) to Compare ({processor_name}):", options=quarters, key=multiselect_key
    )

    if not selected_quarters:
        st.warning("Please select at least one quarter to display the scaling curves.")
        return

    # --------------------------------------------------------------------------
    # 4. CONFIGURATION LOOP & GRAPH RENDERING
    # --------------------------------------------------------------------------
    for cfg in selected_configs:
        df_cfg = df_active[
            (df_active['Configuration'] == cfg) & 
            (df_active['Quarter'].isin(selected_quarters))
        ]

        if df_cfg.empty:
            continue

        st.markdown(f"### ⚙️ Configuration: `{cfg}`")

        vu_df = df_cfg.copy()
        vu_df['Trace_Label'] = vu_df['Quarter'].astype(str) + " - " + vu_df['Build'].astype(str)

        # FIX: Preserve chronological quarter order (earliest first -> latest last)
        quarter_order = [q for q in quarters if q in vu_df['Quarter'].unique()]
        vu_df['Quarter'] = pd.Categorical(vu_df['Quarter'], categories=quarter_order, ordered=True)
        vu_df = vu_df.sort_values(by=['Quarter', 'Build', 'VU'])

        # Build legend trace order chronologically
        trace_order = list(dict.fromkeys(vu_df['Trace_Label']))

        unique_vus = sorted(vu_df['VU'].unique())
        vu_df['VU_Label'] = vu_df['VU'].astype(int).astype(str)
        vu_order_labels = [str(int(v)) for v in unique_vus]

        fig_vu = px.line(
            vu_df, x="VU_Label", y=metric_choice, color="Trace_Label", markers=True, height=420,
            title=f"{metric_choice} Scaling across VUs — {cfg}",
            labels={"VU_Label": "Active Virtual Users (VU)", metric_choice: f"{metric_choice} Value", "Trace_Label": "Quarter & Build Version"},
            category_orders={"Trace_Label": trace_order}
        )

        fig_vu.update_layout(
            hovermode="x unified",
            xaxis=dict(type='category', categoryorder='array', categoryarray=vu_order_labels),
            yaxis=dict(rangemode="tozero")
        )
        st.plotly_chart(fig_vu, use_container_width=True)
        st.markdown("---")


def render_raw_data_tab(df_proc: pd.DataFrame, processor_name: str):
    """Renders Tab 3: Raw Data Explorer Table."""
    st.subheader(f"📋 Raw Data Explorer ({processor_name})")
    st.dataframe(
        df_proc[['Quarter', 'Build', 'Configuration', 'VU', 'NOPM', 'TPM']],
        use_container_width=True, height=550, hide_index=True
    )


def render_processor_dashboard(raw_df: pd.DataFrame, selected_configs: list, selected_builds: list, processor_name: str):
    """Controls sub-tabs for AMD or Intel architectures."""
    df_proc = raw_df[
        (raw_df['Processor'] == processor_name) &
        (raw_df['Configuration'].isin(selected_configs)) &
        (raw_df['Build'].isin(selected_builds))
    ]

    if df_proc.empty:
        st.warning(f"No benchmark data available for {processor_name} under selected filters.")
        return

    sub_tab1, sub_tab2, sub_tab3 = st.tabs([
        "📊 Peak Trends & Gain %", "📈 VU Scaling Curves", "📋 Raw Data Explorer"
    ])

    with sub_tab1:
        render_peak_trends_tab(df_proc, processor_name, selected_configs, selected_builds)

    with sub_tab2:
        render_vu_scaling_tab(df_proc, processor_name, selected_configs, selected_builds)

    with sub_tab3:
        render_raw_data_tab(df_proc, processor_name)


# ==============================================================================
# 4. MAIN APPLICATION ROUTER
# ==============================================================================
def main():
    if "view_dashboard" not in st.session_state:
        st.session_state["view_dashboard"] = False

    if not st.session_state["view_dashboard"]:
        render_landing_page()
        return

    # --------------------------------------------------------------------------
    # CONSOLIDATED STYLING & TAB LOGO RESIZING
    # --------------------------------------------------------------------------
    st.markdown(
        """
        <style>
            /* 1. Tab Button Sizing & Vertical Alignment */
            [data-testid="stTab"],
            [data-baseweb="tab"],
            button[role="tab"] {
                height: auto !important;
                min-height: 52px !important;
                align-items: center !important;
            }

            /* 2. Independent Logo Sizing via Alt Attributes */
            [data-baseweb="tab"] img[alt="AMD"],
            [data-testid="stTab"] img[alt="AMD"],
            button[role="tab"] img[alt="AMD"] {
                height: 38px !important;
                width: auto !important;
                vertical-align: middle !important;
                margin-right: 8px !important;
            }

            [data-baseweb="tab"] img[alt="Intel"],
            [data-testid="stTab"] img[alt="Intel"],
            button[role="tab"] img[alt="Intel"] {
                height: 30px !important;
                width: auto !important;
                vertical-align: middle !important;
                margin-right: 8px !important;
            }

            /* 3. Dropdown Menu Items Wrapping */
            section[data-testid="stSidebar"] [data-baseweb="select"] [role="option"],
            div[data-baseweb="select"] > div {
                white-space: normal !important;
                word-break: break-word !important;
                overflow-wrap: anywhere !important;
                max-height: none !important;
                overflow-y: visible !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.button(
        "🏠 Back to Home Page",
        on_click=lambda: st.session_state.update({"view_dashboard": False}),
    )

    # 1. Retrieve Active Edition & Benchmark Tool selections
    edition = st.session_state.get("server_edition", "Enterprise Server")
    tool = st.session_state.get("benchmark_tool", "HammerDB")

    st.sidebar.info(
        f"Active Edition: **{edition}**\n\nBenchmark Tool: **{tool}**"
    )
    st.sidebar.markdown("---")

    # 2. Load URLs filtered by BOTH Edition and Benchmark Tool
    quarterly_urls = load_quarterly_sheet_urls(
        selected_edition=edition,
        selected_tool=tool,
        config_path=CONFIG_FILE_PATH,
    )

    if not quarterly_urls:
        st.warning(
            f"No Google Sheet URLs configured for **{edition} - {tool}** in"
            " `sheets_config.json`."
        )
        st.stop()

    # 3. Parse only the sheets belonging to the exact combination selected
    raw_df = parse_stacked_matrix_sheet(quarterly_urls)
    if raw_df.empty:
        st.warning(
            f"Awaiting valid data for **{edition} ({tool})**. Please check your"
            " Google Sheet links."
        )
        st.stop()

    st.sidebar.header("🛠️ Global Filters")

    # --------------------------------------------------------------------------
    # 4. CONFIGURATION FILTER (EXPANDER + CHECKBOXES WITH CALLBACKS)
    # --------------------------------------------------------------------------
    all_configs = list(dict.fromkeys(raw_df["Configuration"]))

    with st.sidebar.expander("⚙️ Filter Configurations", expanded=False):
        # Callback function for Select All / Clear All
        def set_all_configs(status: bool):
            for cfg in all_configs:
                st.session_state[f"cfg_{cfg}"] = status

        col_all, col_none = st.columns(2)
        col_all.button("Select All", key="btn_cfg_all", on_click=set_all_configs, args=(True,))
        col_none.button("Clear All", key="btn_cfg_none", on_click=set_all_configs, args=(False,))

        st.markdown("---")

        selected_configs = []
        for cfg in all_configs:
            key_name = f"cfg_{cfg}"
            # Initialize default value to True on first load
            if key_name not in st.session_state:
                st.session_state[key_name] = True

            # Checkbox reads/writes directly from session_state key
            if st.checkbox(cfg, key=key_name):
                selected_configs.append(cfg)

    # --------------------------------------------------------------------------
    # 5. BUILD FILTERS (STACKED VERTICALLY BY MAJOR VERSION)
    # --------------------------------------------------------------------------
    all_builds = list(dict.fromkeys(raw_df["Build"]))
    builds_10_6 = [b for b in all_builds if "10.6" in str(b)]
    builds_11_4 = [b for b in all_builds if "11.4" in str(b)]
    builds_11_8 = [b for b in all_builds if "11.8" in str(b)]
    other_builds = [b for b in all_builds if b not in builds_10_6 + builds_11_4 + builds_11_8]

    st.sidebar.markdown("### Filter Builds")
    sel_10_6 = st.sidebar.multiselect("10.6 Builds:", options=builds_10_6, default=builds_10_6)
    sel_11_4 = st.sidebar.multiselect("11.4 Builds:", options=builds_11_4, default=builds_11_4)
    sel_11_8 = st.sidebar.multiselect("11.8 Builds:", options=builds_11_8, default=builds_11_8)

    selected_builds = sel_10_6 + sel_11_4 + sel_11_8 + other_builds

    # --------------------------------------------------------------------------
    # 6. ARCHITECTURE TABS & MACHINE DETAILS DISPLAY
    # --------------------------------------------------------------------------
    # Load local logos as Base64 Data URIs
    amd_logo = get_base64_logo("amd_logo.png")
    intel_logo = get_base64_logo("intel_logo.png")

    amd_tab_label = f"![AMD]({amd_logo}) AMD Architecture" if amd_logo else "🔘 AMD Architecture"
    intel_tab_label = f"![Intel]({intel_logo}) Intel Architecture" if intel_logo else "🔘 Intel Architecture"

    # Fetch machine details from links.json
    _, _, machine_details = load_links_config("links.json")

    proc_tab_amd, proc_tab_intel = st.tabs([amd_tab_label, intel_tab_label])

    with proc_tab_amd:
        amd_specs = machine_details.get("AMD", {})
        if amd_specs:
            with st.expander("💻 System & Hardware Specs — AMD Architecture", expanded=False):
                for key, val in amd_specs.items():
                    st.markdown(f"• **{key}:** {val}")

        render_processor_dashboard(raw_df, selected_configs, selected_builds, "AMD")

    with proc_tab_intel:
        intel_specs = machine_details.get("INTEL", machine_details.get("Intel", {}))
        if intel_specs:
            with st.expander("💻 System & Hardware Specs — Intel Architecture", expanded=False):
                for key, val in intel_specs.items():
                    st.markdown(f"• **{key}:** {val}")

        render_processor_dashboard(raw_df, selected_configs, selected_builds, "INTEL")

if __name__ == "__main__":
    main()