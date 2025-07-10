import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(layout="wide", page_title="📊 Modern Fill Rate Dashboard")
st.markdown("<h1 style='text-align:center; color:#0e4d92;'>📦 Fill Rate Analytics Dashboard</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel file with 'Base Data' sheet", type=["xlsx", "xls"])

required_columns = ['Row Labels', 'Sum of sku_po_qty', 'Sum of sku_grn_qty', 'Sum of QFR', 'Sum of LFR', 'Manufacturer Name', 'WH Name']

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Base Data")

        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"❌ Required columns not found: {', '.join(missing)}")
            st.stop()

        df = df.rename(columns={
            'Row Labels': 'Category',
            'Sum of sku_po_qty': 'PO Qty',
            'Sum of sku_grn_qty': 'GRN Qty',
            'Sum of QFR': 'QFR',
            'Sum of LFR': 'LFR',
            'PO Date': 'PO Date',
            'Region': 'Region',
            'Manufacturer Name': 'Manufacturer',
            'WH Name': 'Warehouse',
        })

        for col in ['PO Qty', 'GRN Qty', 'QFR', 'LFR']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['Category'])

        if 'PO Date' in df.columns:
            df['PO Date'] = pd.to_datetime(df['PO Date'], errors='coerce')

        st.sidebar.header("Filters")
        selected_categories = st.sidebar.multiselect("Category", df['Category'].unique(), default=list(df['Category'].unique()))
        selected_manu = st.sidebar.multiselect("Manufacturer", df['Manufacturer'].dropna().unique())
        selected_wh = st.sidebar.multiselect("Warehouse", df['Warehouse'].dropna().unique())

        if 'Region' in df.columns:
            selected_region = st.sidebar.multiselect("Region", df['Region'].dropna().unique())
        else:
            selected_region = None

        df_filtered = df[df['Category'].isin(selected_categories)]
        if selected_manu:
            df_filtered = df_filtered[df_filtered['Manufacturer'].isin(selected_manu)]
        if selected_wh:
            df_filtered = df_filtered[df_filtered['Warehouse'].isin(selected_wh)]
        if selected_region:
            df_filtered = df_filtered[df_filtered['Region'].isin(selected_region)]

        total_po = df_filtered['PO Qty'].sum()
        total_grn = df_filtered['GRN Qty'].sum()
        fill_rate = (total_grn / total_po) * 100 if total_po > 0 else 0
        avg_qfr = df_filtered['QFR'].mean() * 100
        avg_lfr = df_filtered['LFR'].mean() * 100

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("PO Qty", f"{int(total_po):,}")
        col2.metric("GRN Qty", f"{int(total_grn):,}")
        col3.metric("Fill Rate", f"{fill_rate:.2f}%")
        col4.metric("QFR", f"{avg_qfr:.2f}%")
        col5.metric("LFR", f"{avg_lfr:.2f}%")

        st.markdown("### 📊 Fill Rate by Category")
        chart_data = df_filtered.groupby("Category")[['PO Qty', 'GRN Qty']].sum().reset_index()
        chart_data["Fill Rate"] = (chart_data["GRN Qty"] / chart_data["PO Qty"]) * 100
        fig_cat = px.bar(chart_data, x="Category", y="Fill Rate", color="Fill Rate", color_continuous_scale="Blues")
        st.plotly_chart(fig_cat, use_container_width=True)

        st.markdown("### 🧾 Manufacturer & Warehouse Breakdown")
        group_df = df_filtered.groupby(['Category', 'Manufacturer', 'Warehouse']).agg({
            'PO Qty': 'sum',
            'GRN Qty': 'sum',
            'QFR': 'mean',
            'LFR': 'mean'
        }).reset_index()
        group_df['QFR'] = group_df['QFR'] * 100
        group_df['LFR'] = group_df['LFR'] * 100
        st.dataframe(group_df)

        def to_excel(data):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                data.to_excel(writer, index=False, sheet_name="Filtered")
            output.seek(0)
            return output

        st.download_button("📥 Download Filtered Data", to_excel(df_filtered), "filtered_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
else:
    st.info("⬆️ Upload an Excel file with a 'Base Data' sheet to begin.")
