import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import base64
from datetime import datetime

st.set_page_config(layout="wide", page_title="PowerBI-style Dashboard")
st.title("📊 Fill Rate Dashboard")

uploaded_file = st.file_uploader("📤 Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file, skiprows=3)

        df = df_raw.rename(columns={
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

        st.sidebar.header("🔎 Filters")
        categories = df['Category'].dropna().unique()
        selected_categories = st.sidebar.multiselect("Select Categories", categories, default=categories)

        manufacturers = df['Manufacturer'].dropna().unique() if 'Manufacturer' in df.columns else []
        selected_manu = st.sidebar.multiselect("Select Manufacturer", manufacturers, default=manufacturers)

        warehouses = df['Warehouse'].dropna().unique() if 'Warehouse' in df.columns else []
        selected_wh = st.sidebar.multiselect("Select WH Name", warehouses, default=warehouses)

        if 'Region' in df.columns:
            regions = df['Region'].dropna().unique()
            selected_regions = st.sidebar.multiselect("Select Regions", regions, default=regions)
        else:
            selected_regions = df['Region'].unique() if 'Region' in df.columns else []

        if 'PO Date' in df.columns:
            min_date = df['PO Date'].min()
            max_date = df['PO Date'].max()
            selected_date = st.sidebar.date_input("Select PO Date Range", [min_date, max_date])
        else:
            selected_date = None

        filtered_df = df[df['Category'].isin(selected_categories)]

        if selected_manu:
            filtered_df = filtered_df[filtered_df['Manufacturer'].isin(selected_manu)]

        if selected_wh:
            filtered_df = filtered_df[filtered_df['Warehouse'].isin(selected_wh)]

        if 'Region' in df.columns:
            filtered_df = filtered_df[filtered_df['Region'].isin(selected_regions)]

        if selected_date and isinstance(selected_date, list) and len(selected_date) == 2:
            start, end = selected_date
            filtered_df = filtered_df[(filtered_df['PO Date'] >= pd.to_datetime(start)) & (filtered_df['PO Date'] <= pd.to_datetime(end))]

        total_po = filtered_df['PO Qty'].sum()
        total_grn = filtered_df['GRN Qty'].sum()
        fill_rate = (total_grn / total_po) * 100 if total_po > 0 else 0
        avg_qfr = filtered_df['QFR'].mean() if 'QFR' in filtered_df.columns else 0
        avg_lfr = filtered_df['LFR'].mean() if 'LFR' in filtered_df.columns else 0

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📦 PO Qty", f"{int(total_po):,}")
        col2.metric("✅ GRN Qty", f"{int(total_grn):,}")
        col3.metric("📈 Fill Rate", f"{fill_rate:.2f}%")
        col4.metric("📊 Total QFR", f"{avg_qfr * 100:.2f}%")
        col5.metric("📉 Total LFR", f"{avg_lfr * 100:.2f}%")

        st.subheader("📌 Fill Rate by Category")
        df_bar = filtered_df[['Category', 'PO Qty', 'GRN Qty']].copy()
        df_bar["Fill Rate %"] = (df_bar["GRN Qty"] / df_bar["PO Qty"]) * 100
        fig_bar = px.bar(df_bar, x="Category", y="Fill Rate %", color="Fill Rate %", color_continuous_scale="Blues")
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("📆 PO vs GRN by Date")
        if 'PO Date' in filtered_df.columns:
            df_date = filtered_df.groupby('PO Date')[['PO Qty', 'GRN Qty']].sum().reset_index()
            fig_date = px.bar(df_date, x='PO Date', y=['PO Qty', 'GRN Qty'], barmode='group')
            st.plotly_chart(fig_date, use_container_width=True)

        st.subheader("🥧 GRN Distribution by Category")
        fig_pie = px.pie(filtered_df, names="Category", values="GRN Qty", title="GRN Share by Category")
        st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("📈 Cumulative PO vs GRN (Simulated)")
        df_area = filtered_df.copy()
        df_area["Cumulative PO"] = df_area["PO Qty"].cumsum()
        df_area["Cumulative GRN"] = df_area["GRN Qty"].cumsum()
        fig_area = px.area(df_area, x="Category", y=["Cumulative PO", "Cumulative GRN"])
        st.plotly_chart(fig_area, use_container_width=True)

        st.subheader("🏷️ Vendor and Warehouse Breakdown")
        if 'Vendor' in filtered_df.columns and 'Warehouse' in filtered_df.columns:
            group_df = filtered_df.groupby(['Category', 'Vendor', 'Warehouse']).agg({
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

        st.download_button("📁 Download Filtered Data (Excel)", to_excel(filtered_df), "filtered_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("🖨️ Export Bar Chart as PDF")
        try:
            import plotly.io as pio
            fig_bar.update_layout(title_text="Fill Rate by Category")
            pdf_bytes = pio.to_image(fig_bar, format="pdf", engine="kaleido")
            b64_pdf = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="fill_rate_chart.pdf">📄 Download Bar Chart PDF</a>'
            st.markdown(href, unsafe_allow_html=True)
        except Exception as e:
            st.warning(f"PDF export failed: {e}")

    except Exception as e:
        st.error(f"⚠️ Error loading file: {e}")
else:
    st.info("⬆️ Upload Excel or CSV file to begin.")
