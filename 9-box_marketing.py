import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

st.set_page_config(page_title="Marketing 9-Box Chart", layout="wide")

# ==========================================
# HEADER
# ==========================================
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); 
                padding: 1.2rem; border-radius: 10px; margin-bottom: 1rem; color: white;
                text-align: center;">
        <h2 style="margin:0; color: #38bdf8; font-weight: 800;">
            📊 Marketing Performance 9-Box Chart
        </h2>
        <p style="margin: 0.2rem 0 0 0; opacity: 0.9; font-size: 1rem;">
            Strategic Financial Output vs People Score (Automated Matrix Mapping)
        </p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# FILE UPLOADER & DATA LOADING
# ==========================================
st.markdown("### 📂 Upload Source Data")
uploaded_file = st.file_uploader("Upload Excel file (e.g., Financial_Metric_Marketing_Dummy.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("☝️ Please upload your Excel data file to begin.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)

    req_sheets = ['PL_Local', 'Matrix']
    missing_sheets = [s for s in req_sheets if s not in xls.sheet_names]
    if missing_sheets:
        st.error(f"⚠️ The following sheets were not found in the Excel file: {', '.join(missing_sheets)}")
        st.stop()

    # 1. READ PL_LOCAL (Raw Data)
    df_pl = pd.read_excel(xls, sheet_name='PL_Local')


    def clean_curr(x):
        if pd.isna(x): return 0.0
        if isinstance(x, (int, float)): return float(x)
        s = str(x).strip().replace('.', '').replace(',', '.')
        try:
            return float(s)
        except:
            return 0.0


    col_gs = 'Gross Sales' if 'Gross Sales' in df_pl.columns else None
    col_ret = 'Sales Return' if 'Sales Return' in df_pl.columns else ('Return' if 'Return' in df_pl.columns else None)
    col_cm = 'Contribution Margin'
    col_op = 'Operating Income' if 'Operating Income' in df_pl.columns else 'Operating Profit'

    if col_gs and col_ret:
        df_pl['Calc_Sales'] = df_pl[col_gs].apply(clean_curr) + df_pl[col_ret].apply(clean_curr)
    elif col_gs:
        df_pl['Calc_Sales'] = df_pl[col_gs].apply(clean_curr)
    else:
        df_pl['Calc_Sales'] = 0.0

    df_pl[col_cm] = df_pl[col_cm].apply(clean_curr) if col_cm in df_pl.columns else 0.0
    df_pl[col_op] = df_pl[col_op].apply(clean_curr) if col_op in df_pl.columns else 0.0

    # 2. FIND PEOPLE SCORE
    score_map = {}
    if 'Head-Count' in xls.sheet_names:
        df_hc = pd.read_excel(xls, sheet_name='Head-Count')
        ps_cands = [c for c in df_hc.columns if 'score' in c.lower() or 'people' in c.lower()]
        if ps_cands:
            score_map = dict(
                zip(df_hc['Name'].astype(str).str.strip(), pd.to_numeric(df_hc[ps_cands[0]], errors='coerce')))

    # 3. READ MATRIX SHEET
    df_mat_raw = pd.read_excel(xls, sheet_name='Matrix', header=None)

    results = []
    for idx, row in df_mat_raw.iterrows():
        try:
            no = int(row[0])
        except:
            continue

        name = str(row[1]).strip()
        if pd.isna(name) or name.lower() in ['nan', '']: continue

        pillar = str(row[2]).strip() if pd.notna(row[2]) else ""
        brands = [str(x).strip() for x in row[3:9] if pd.notna(x) and str(x).strip() != '']
        categories = [str(x).strip() for x in row[9:15] if pd.notna(x) and str(x).strip() != '']

        mask = pd.Series(False, index=df_pl.index)

        if pillar and pillar.lower() != 'nan':
            mask = mask | (df_pl['Pillar'].astype(str).str.strip() == pillar)
        if brands:
            mask = mask | (df_pl['Brand'].astype(str).str.strip().isin(brands))
        if categories:
            mask = mask | (df_pl['Category'].astype(str).str.strip().isin(categories))

        df_filtered = df_pl[mask]
        sales_val = df_filtered['Calc_Sales'].sum()
        cm_val = df_filtered[col_cm].sum()
        op_val = df_filtered[col_op].sum()

        p_score = score_map.get(name, 0.0)
        if pd.isna(p_score) or p_score == 0.0:
            if len(row) > 15 and pd.notna(row[15]):
                try:
                    p_score = float(row[15])
                except:
                    p_score = 0.0

        results.append({
            'Name': name,
            'People Score': p_score,
            'Sales': sales_val,
            'Contribution Margin': cm_val,
            'Operating Profit': op_val
        })

    df_res = pd.DataFrame(results)

    if df_res.empty:
        st.warning("⚠️ No names or marketing data were successfully mapped from the 'Matrix' sheet.")
        st.stop()

    total_persons = len(df_res)

    # 4. PERCENTAGE CALCULATION
    sales_safe = df_res['Sales'].replace(0, np.nan)
    df_res['% CM'] = (df_res['Contribution Margin'] / sales_safe) * 100
    df_res['% Op. Profit'] = (df_res['Operating Profit'] / sales_safe) * 100

    df_res['% CM'] = df_res['% CM'].fillna(0)
    df_res['% Op. Profit'] = df_res['% Op. Profit'].fillna(0)

    # ==========================================
    # SETTINGS: Y-AXIS & THRESHOLDS
    # ==========================================
    st.markdown("---")
    st.markdown("### ⚙️ 9-Box Configuration")

    col_set1, col_set2 = st.columns(2)
    with col_set1:
        y_metric = st.selectbox(
            "Select Y-Axis (Financial Output Metric):",
            ["Contribution Margin", "Operating Profit", "% CM", "% Op. Profit"]
        )
    with col_set2:
        st.info("💡 **X-Axis** is automatically mapped to **People Score**. **Bubble Size** represents **Sales**.")

    st.markdown("#### 🎚️ Midpoint Thresholds")
    mid_x = df_res['People Score'].mean() if df_res['People Score'].sum() > 0 else 5.0
    mid_y = df_res[y_metric].mean() if df_res[y_metric].sum() > 0 else (50.0 if '%' in y_metric else 1000000.0)

    step_y = 5.0 if '%' in y_metric else 1000000.0

    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
    with col_t1:
        x_low = st.number_input("X-Axis Low-Med (Score)", value=float(mid_x * 0.8), step=0.5)
    with col_t2:
        x_high = st.number_input("X-Axis Med-High (Score)", value=float(mid_x * 1.2), step=0.5)
    with col_t3:
        y_low = st.number_input(f"Y-Axis Low-Med ({y_metric})", value=float(mid_y * 0.8), step=step_y)
    with col_t4:
        y_high = st.number_input(f"Y-Axis Med-High ({y_metric})", value=float(mid_y * 1.2), step=step_y)

    # ==========================================
    # ASSIGN 9-BOX CATEGORY
    # ==========================================
    cat_protect = "Protect & scale"
    cat_keep = "Keep & strengthen"
    cat_move = "Move to higher output"
    cat_share = "Share across brands"
    cat_simplify = "Simplify the work"
    cat_reduce_some = "Reduce some"
    cat_upskill = "Upskill or replace"
    cat_replace = "Replace with tools"
    cat_reduce_most = "Reduce most (over half)"


    def assign_9box(df_target):
        conds = [
            (df_target['People Score'] > x_high) & (df_target[y_metric] > y_high),
            (df_target['People Score'] > x_high) & (df_target[y_metric] >= y_low) & (df_target[y_metric] <= y_high),
            (df_target['People Score'] > x_high) & (df_target[y_metric] < y_low),

            (df_target['People Score'] >= x_low) & (df_target['People Score'] <= x_high) & (
                        df_target[y_metric] > y_high),
            (df_target['People Score'] >= x_low) & (df_target['People Score'] <= x_high) & (
                        df_target[y_metric] >= y_low) & (df_target[y_metric] <= y_high),
            (df_target['People Score'] >= x_low) & (df_target['People Score'] <= x_high) & (
                        df_target[y_metric] < y_low),

            (df_target['People Score'] < x_low) & (df_target[y_metric] > y_high),
            (df_target['People Score'] < x_low) & (df_target[y_metric] >= y_low) & (df_target[y_metric] <= y_high),
            (df_target['People Score'] < x_low) & (df_target[y_metric] < y_low)
        ]
        choices = [
            cat_protect, cat_keep, cat_move,
            cat_share, cat_simplify, cat_reduce_some,
            cat_upskill, cat_replace, cat_reduce_most
        ]
        return np.select(conds, choices, default=cat_simplify)


    df_res['9-Box Category'] = assign_9box(df_res)
    df_res['Bubble_Size'] = df_res['Sales'].abs().replace(0, 1)

    # ==========================================
    # SCATTER PLOT VISUALIZATION
    # ==========================================
    color_map = {
        cat_protect: '#a7c957',  # Green
        cat_keep: '#c5e1a5',  # Light Green
        cat_move: '#a8dadc',  # Light Blue
        cat_share: '#ffd166',  # Yellow/Orange
        cat_simplify: '#ffe66d',  # Yellow
        cat_reduce_some: '#f4a261',  # Orange
        cat_upskill: '#f4a261',  # Orange
        cat_replace: '#e76f51',  # Dark Orange
        cat_reduce_most: '#d62828'  # Dark Red
    }

    fig = px.scatter(
        df_res,
        x='People Score',
        y=y_metric,
        color='9-Box Category',
        color_discrete_map=color_map,
        size='Bubble_Size',
        hover_name='Name',
        hover_data={'Bubble_Size': False, 'Sales': ':,.0f', 'Contribution Margin': ':,.0f',
                    'Operating Profit': ':,.0f'},
        title=f"9-Box Matrix: {y_metric} vs People Score (Bubble Size: Sales)",
        size_max=40,
        render_mode='svg'
    )

    fig.add_vline(x=x_low, line_dash="dash", line_color="black", opacity=0.7)
    fig.add_vline(x=x_high, line_dash="dash", line_color="black", opacity=0.7)
    fig.add_hline(y=y_low, line_dash="dash", line_color="black", opacity=0.7)
    fig.add_hline(y=y_high, line_dash="dash", line_color="black", opacity=0.7)

    # ---------------------------------------------------------
    # ANNOTATION (LABEL & COUNT) IN THE CENTER OF EACH QUADRANT
    # ---------------------------------------------------------
    x_min, x_max = df_res['People Score'].min(), df_res['People Score'].max()
    y_min, y_max = df_res[y_metric].min(), df_res[y_metric].max()

    x_min_plot = min(x_min, x_low) - (abs(x_max - x_min) * 0.1 if x_max != x_min else 1)
    x_max_plot = max(x_max, x_high) + (abs(x_max - x_min) * 0.1 if x_max != x_min else 1)
    y_min_plot = min(y_min, y_low) - (abs(y_max - y_min) * 0.1 if y_max != y_min else 1000)
    y_max_plot = max(y_max, y_high) + (abs(y_max - y_min) * 0.1 if y_max != y_min else 1000)

    mid_x_low = (x_min_plot + x_low) / 2
    mid_x_mid = (x_low + x_high) / 2
    mid_x_high = (x_high + x_max_plot) / 2

    mid_y_low = (y_min_plot + y_low) / 2
    mid_y_mid = (y_low + y_high) / 2
    mid_y_high = (y_high + y_max_plot) / 2

    box_counts = df_res['9-Box Category'].value_counts().to_dict()

    annotations_data = [
        {'cat': cat_move, 'x': mid_x_low, 'y': mid_y_high},
        {'cat': cat_keep, 'x': mid_x_mid, 'y': mid_y_high},
        {'cat': cat_protect, 'x': mid_x_high, 'y': mid_y_high},
        {'cat': cat_reduce_some, 'x': mid_x_low, 'y': mid_y_mid},
        {'cat': cat_simplify, 'x': mid_x_mid, 'y': mid_y_mid},
        {'cat': cat_share, 'x': mid_x_high, 'y': mid_y_mid},
        {'cat': cat_reduce_most, 'x': mid_x_low, 'y': mid_y_low},
        {'cat': cat_replace, 'x': mid_x_mid, 'y': mid_y_low},
        {'cat': cat_upskill, 'x': mid_x_high, 'y': mid_y_low}
    ]

    for item in annotations_data:
        count = box_counts.get(item['cat'], 0)
        if count > 0:
            person_text = "Person" if count == 1 else "Persons"
            fig.add_annotation(
                x=item['x'], y=item['y'],
                xref="x", yref="y",
                text=f"<span style='color:#1f2937;'><b>{item['cat']}</b></span><br><b>{count} {person_text}</b>",
                showarrow=False,
                font=dict(size=12, color="#374151"),
                bgcolor="rgba(255, 255, 255, 0.85)",
                bordercolor="rgba(15, 23, 42, 0.3)",
                borderwidth=1, borderpad=4
            )

    # ADD TOTAL OVERALL ANNOTATION
    total_person_text = "Person" if total_persons == 1 else "Persons"
    fig.add_annotation(
        x=0.98, y=0.98, xref="paper", yref="paper",
        text=f"<b>TOTAL:<br>{total_persons} {total_person_text}</b>",
        showarrow=False,
        font=dict(size=13, color="white"),
        bgcolor="#374151",
        bordercolor="black",
        borderwidth=1, borderpad=5
    )

    fig.update_xaxes(range=[x_min_plot, x_max_plot])
    fig.update_yaxes(range=[y_min_plot, y_max_plot])

    fig.update_layout(
        height=650,
        hovermode="closest",
        margin=dict(t=50, b=50, l=50, r=50),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=None)
    )

    st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False,
                                                           'toImageButtonOptions': {'format': 'png',
                                                                                    'filename': '9box_chart',
                                                                                    'scale': 2}})
    st.caption("📷 *Tip: Click the camera icon in the top right corner of the chart to download the image (PNG).*")

    # ==========================================
    # HTML 9-BOX VISUAL TABLE
    # ==========================================
    st.markdown("---")
    st.markdown("### 🗂️ Strategic 9-Box Matrix Summary")

    html_matrix = f"""
    <div style="font-family: Arial, sans-serif; border: 1px solid #ddd; border-collapse: collapse; width: 100%; text-align: center;">
        <div style="display: flex; background-color: #1e3a5f; color: white; padding: 10px; font-weight: bold;">
            <div style="flex: 1; display: flex; align-items: center; justify-content: center; background-color: #0f172a; border-radius: 5px;">TOTAL: {total_persons} {total_person_text.upper()}</div>
            <div style="flex: 3; display: flex; align-items: center; justify-content: center;">FINANCIAL OUTPUT ({y_metric}) → higher is better</div>
        </div>
        <div style="display: flex; background-color: #e6e6e6; font-weight: bold; border-bottom: 2px solid white;">
            <div style="flex: 1; background-color: #1e3a5f; color: white; display: flex; align-items: center; justify-content: center; writing-mode: vertical-rl; transform: rotate(180deg);">
                PEOPLE SCORE → higher is better
            </div>
            <div style="flex: 3; display: flex; flex-direction: column;">
                <div style="display: flex; background-color: #e6e6e6; padding: 5px 0;">
                    <div style="flex: 1;">LOW</div>
                    <div style="flex: 1;">MID</div>
                    <div style="flex: 1;">HIGH</div>
                </div>

                <div style="display: flex; height: 120px;">
                    <div style="flex: 0.3; background-color: #dbe4f0; display: flex; align-items: center; justify-content: center; font-weight: bold; border-right: 2px solid white; border-bottom: 2px solid white;">HIGH</div>
                    <div style="flex: 1; background-color: {color_map[cat_move]}; padding: 10px; border-right: 2px solid white; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_move}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_move, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_keep]}; padding: 10px; border-right: 2px solid white; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_keep}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_keep, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_protect]}; padding: 10px; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_protect}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_protect, 0)}</span>
                    </div>
                </div>

                <div style="display: flex; height: 120px;">
                    <div style="flex: 0.3; background-color: #dbe4f0; display: flex; align-items: center; justify-content: center; font-weight: bold; border-right: 2px solid white; border-bottom: 2px solid white;">MID</div>
                    <div style="flex: 1; background-color: {color_map[cat_reduce_some]}; padding: 10px; border-right: 2px solid white; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_reduce_some}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_reduce_some, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_simplify]}; padding: 10px; border-right: 2px solid white; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_simplify}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_simplify, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_share]}; padding: 10px; border-bottom: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_share}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_share, 0)}</span>
                    </div>
                </div>

                <div style="display: flex; height: 120px;">
                    <div style="flex: 0.3; background-color: #dbe4f0; display: flex; align-items: center; justify-content: center; font-weight: bold; border-right: 2px solid white;">LOW</div>
                    <div style="flex: 1; background-color: {color_map[cat_reduce_most]}; color: white; padding: 10px; border-right: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_reduce_most}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_reduce_most, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_replace]}; padding: 10px; border-right: 2px solid white; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_replace}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_replace, 0)}</span>
                    </div>
                    <div style="flex: 1; background-color: {color_map[cat_upskill]}; padding: 10px; display: flex; flex-direction: column; justify-content: center;">
                        <span style="font-size: 18px; font-weight: bold;">{cat_upskill}</span>
                        <span style="font-size: 24px; font-weight: 900; margin-top: 5px;">{box_counts.get(cat_upskill, 0)}</span>
                    </div>
                </div>
            </div>
        </div>
    </div>
    """.replace('\n', '')

    st.markdown(html_matrix, unsafe_allow_html=True)

    # ==========================================
    # DATA TABLE RESULTS & EXCEL EXPORT
    # ==========================================
    st.markdown("---")
    st.markdown("### 📋 Data Results per Person/Metric")

    df_display = df_res[['Name', 'People Score', y_metric, 'Sales', '9-Box Category']].copy()
    df_display.rename(columns={'Sales': 'Sales (Bubble Size)'}, inplace=True)
    df_display = df_display.sort_values('9-Box Category')

    format_dict = {'People Score': '{:.2f}'}
    if '%' in y_metric:
        format_dict[y_metric] = '{:.2f}%'
    else:
        format_dict[y_metric] = 'Rp {:,.0f}'
    format_dict['Sales (Bubble Size)'] = 'Rp {:,.0f}'

    # Setup columns for download buttons
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        # Create Excel File in Memory for the Box Summary Table
        buffer_summary = io.BytesIO()
        with pd.ExcelWriter(buffer_summary, engine='openpyxl') as writer:
            summary_data = [
                ['', 'LOW FIN', 'MID FIN', 'HIGH FIN', 'TOTAL'],
                ['HIGH PEOPLE', f"{cat_move}: {box_counts.get(cat_move, 0)}",
                 f"{cat_keep}: {box_counts.get(cat_keep, 0)}", f"{cat_protect}: {box_counts.get(cat_protect, 0)}", ""],
                ['MID PEOPLE', f"{cat_reduce_some}: {box_counts.get(cat_reduce_some, 0)}",
                 f"{cat_simplify}: {box_counts.get(cat_simplify, 0)}", f"{cat_share}: {box_counts.get(cat_share, 0)}",
                 ""],
                ['LOW PEOPLE', f"{cat_reduce_most}: {box_counts.get(cat_reduce_most, 0)}",
                 f"{cat_replace}: {box_counts.get(cat_replace, 0)}", f"{cat_upskill}: {box_counts.get(cat_upskill, 0)}",
                 f"{total_persons} {total_person_text}"]
            ]
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, index=False, header=False, sheet_name='Matrix_Summary')

        st.download_button(
            label="📥 Download Box Table (Excel)",
            data=buffer_summary.getvalue(),
            file_name="Marketing_9Box_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="btn_dl_summary"
        )

    with col_dl2:
        # Create Excel File in Memory for the Detailed Data Table
        buffer_data = io.BytesIO()
        with pd.ExcelWriter(buffer_data, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name='Data_Result')

        st.download_button(
            label="📥 Download Data Results (Excel)",
            data=buffer_data.getvalue(),
            file_name="Marketing_Data_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
            key="btn_dl_data"
        )

    st.dataframe(df_display.style.format(format_dict), use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"⚠️ An error occurred while processing the Excel file. Error Log: {e}")