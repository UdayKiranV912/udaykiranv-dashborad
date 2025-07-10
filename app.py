import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import base64

st.set_page_config(layout="wide", page_title="📊 Modern Fill Rate Dashboard")
st.markdown("<h1 style='text-align:center; color:#0e4d92;'>📦 Fill Rate Analytics Dashboard</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel file with 'Base Data' sheet", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Base Data")

        df = df.rename(columns={
            'Row Labels': 'Category',
            'Sum of sku_po_qty': 'PO Qty',
            'Sum of sku_grn_qty': 'GRN Qty',
            'Sum of QFR': 'QFR',
            'Sum of sku_po_line': 'PO Lines',
            'Sum of sku_grn_line': 'GRN Lines',
            'Sum of LFR': 'LFR',
            'Sum of po_amount': 'PO Amount',
            'Sum of grn_amount': 'GRN Amount',
            'Sum of Vendor loss A/c': 'Vendor Loss',
            'PO Date': 'PO Date',
            'Region': 'Region',
            'Manufacturer Name': 'Manufacturer',
            'WH Name': 'Warehouse',
            'Vendor Name': 'Vendor'
        })

        numeric_cols = ['PO Qty', 'GRN Qty', 'QFR', 'LFR', 'PO Lines', 'GRN Lines', 'PO Amount', 'GRN Amount', 'Vendor Loss']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['Category'])

        if 'PO Date' in df.columns:
            df['PO Date'] = pd.to_datetime(df['PO Date'], errors='coerce')

        st.sidebar.header("Filters")
        categories = df['Category'].dropna().unique()
        selected_categories = st.sidebar.multiselect("Category", categories, default=categories)

        manufacturers = df['Manufacturer'].dropna().unique()
        selected_manu = st.sidebar.multiselect("Manufacturer", manufacturers, default=manufacturers)

        warehouses = df['Warehouse'].dropna().unique()
        selected_wh = st.sidebar.multiselect("Warehouse", warehouses, default=warehouses)

        if 'Region' in df.columns:
            regions = df['Region'].dropna().unique()
            selected_regions = st.sidebar.multiselect("Region", regions, default=regions)
        else:
            selected_regions = []

        df_filtered = df[
            (df['Category'].isin(selected_categories)) &
            (df['Manufacturer'].isin(selected_manu)) &
            (df['Warehouse'].isin(selected_wh))
        ]
        if 'Region' in df.columns:
            df_filtered = df_filtered[df_filtered['Region'].isin(selected_regions)]

        total_po = df_filtered['PO Qty'].sum()
        total_grn = df_filtered['GRN Qty'].sum()
        fill_rate = (total_grn / total_po) * 100 if total_po > 0 else 0
        avg_qfr = df_filtered['QFR'].mean() if 'QFR' in df_filtered.columns else 0
        avg_lfr = df_filtered['LFR'].mean() if 'LFR' in df_filtered.columns else 0

        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("PO Qty", f"{int(total_po):,}")
        kpi2.metric("GRN Qty", f"{int(total_grn):,}")
        kpi3.metric("Fill Rate", f"{fill_rate:.2f}%")
        kpi4.metric("QFR", f"{avg_qfr * 100:.2f}%")
        kpi5.metric("LFR", f"{avg_lfr * 100:.2f}%")

        st.markdown("### 📊 Category-Wise Fill Rate")
        chart_data = df_filtered.groupby("Category")[['PO Qty', 'GRN Qty']].sum().reset_index()
        chart_data["Fill Rate"] = (chart_data["GRN Qty"] / chart_data["PO Qty"]) * 100
        fig_cat = px.bar(chart_data, x="Category", y="Fill Rate", color="Fill Rate", color_continuous_scale="Blues")
        st.plotly_chart(fig_cat, use_container_width=True)

        st.markdown("### 🧾 Manufacturer & Warehouse Breakdown")
        sub_df = df_filtered.groupby(['Category', 'Manufacturer', 'Warehouse']).agg({
            'PO Qty': 'sum',
            'GRN Qty': 'sum',
            'QFR': 'mean',
            'LFR': 'mean'
        }).reset_index()
        sub_df['QFR'] = sub_df['QFR'] * 100
        sub_df['LFR'] = sub_df['LFR'] * 100
        st.dataframe(sub_df)

        def to_excel(data):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                data.to_excel(writer, index=False, sheet_name="Filtered")
            output.seek(0)
            return output

        st.download_button("📥 Download Filtered Data", to_excel(df_filtered), "filtered_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("⬆️ Upload an Excel file to begin.")
