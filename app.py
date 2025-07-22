import streamlit as st
import pandas as pd
import plotly.express as px
from xhtml2pdf import pisa
import matplotlib.pyplot as plt
import io
import base64
from PIL import Image

st.set_page_config(page_title="Universal Fill Rate Dashboard", layout="wide")

# Sidebar uploader
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel or CSV File", type=["xlsx", "xls", "csv"])

# Load and cache data
@st.cache_data
def load_data(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        excel_file = pd.ExcelFile(file)
        sheet_name = "Base Data" if "Base Data" in excel_file.sheet_names else excel_file.sheet_names[0]
        df = pd.read_excel(file, sheet_name=sheet_name)
    # Convert fill rate columns if exist
    for col in ["sku_level_fill_rate", "overall_po_fill_rate"]:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].str.rstrip('%').astype(float)
    return df

if not uploaded_file:
    st.warning("Upload a file to continue.")
    st.stop()

df = load_data(uploaded_file)

# Filters with memory
if "filters" not in st.session_state:
    st.session_state.filters = {
        "manufacturer": [], "category": [],
        "subcategory": [], "location": []
    }

# Dynamic filter setup
def create_filter(label, col_name):
    options = df[col_name].dropna().unique().tolist() if col_name in df.columns else []
    default = [v for v in st.session_state.filters.get(label.lower(), []) if v in options]
    selected = st.sidebar.multiselect(label, options, default=default)
    st.session_state.filters[label.lower()] = selected
    return selected

manufacturer = create_filter("Manufacturer", "manufacturer_name")
category = create_filter("Category", "category_name")
subcategory = create_filter("Subcategory", "subcategory_name")
location = create_filter("Location", "wh_name")

# Apply filters
filtered_df = df.copy()
for col, selected in zip(
    ["manufacturer_name", "category_name", "subcategory_name", "wh_name"],
    [manufacturer, category, subcategory, location]):
    if selected and col in filtered_df.columns:
        filtered_df = filtered_df[filtered_df[col].isin(selected)]

# Define expected columns
expected_cols = [
    "sku_po_qty", "sku_grn_qty", "sku_level_fill_rate", "sku_po_line",
    "sku_grn_line", "overall_po_fill_rate", "po_amount", "grn_amount",
    "Vendor loss A/c"
]
agg_cols = [col for col in expected_cols if col in filtered_df.columns]
group_cols = [col for col in ["manufacturer_name", "category_name", "subcategory_name", "wh_name"] if col in filtered_df.columns]

# Validation
if not group_cols:
    st.error("Required grouping columns are missing.")
    st.stop()

summary_df = filtered_df.groupby(group_cols)[agg_cols].sum().reset_index()

# KPI Metrics
def safe_mean(col):
    return filtered_df[col].mean() if col in filtered_df.columns else 0.0

st.title("📊 Fill Rate Dashboard")

col1, col2, col3 = st.columns(3)
col1.metric("Total SKU PO Qty", f"{filtered_df.get('sku_po_qty', pd.Series([0])).sum():,.0f}")
col2.metric("Total SKU GRN Qty", f"{filtered_df.get('sku_grn_qty', pd.Series([0])).sum():,.0f}")
col3.metric("Vendor Loss A/c", f"{filtered_df.get('Vendor loss A/c', pd.Series([0])).sum():,.0f}")

col4, col5, col6 = st.columns(3)
col4.metric("Total PO Lines", f"{filtered_df.get('sku_po_line', pd.Series([0])).sum():,.0f}")
col5.metric("Total GRN Lines", f"{filtered_df.get('sku_grn_line', pd.Series([0])).sum():,.0f}")
col6.metric("PO Amount", f"{filtered_df.get('po_amount', pd.Series([0])).sum():,.0f}")

col7, col8 = st.columns(2)
col7.metric("QFR (%)", f"{safe_mean('sku_level_fill_rate'):.2f}%")
col8.metric("LFR (%)", f"{safe_mean('overall_po_fill_rate'):.2f}%")

# Charts
figs = {}

def save_plotly_fig(fig, name):
    buf = io.BytesIO()
    fig.write_image(buf, format="png")
    buf.seek(0)
    figs[name] = buf.read()

def show_and_capture_chart(title, fig_obj, key):
    st.subheader(title)
    st.plotly_chart(fig_obj, use_container_width=True)
    save_plotly_fig(fig_obj, key)

if "category_name" in summary_df.columns and "sku_level_fill_rate" in summary_df.columns:
    fig = px.bar(summary_df, x="category_name", y="sku_level_fill_rate", color="category_name")
    show_and_capture_chart("📦 QFR by Category", fig, "category")

if "subcategory_name" in summary_df.columns and "sku_level_fill_rate" in summary_df.columns:
    fig = px.bar(summary_df, x="subcategory_name", y="sku_level_fill_rate", color="subcategory_name")
    show_and_capture_chart("🔍 QFR by Subcategory", fig, "subcategory")

if "manufacturer_name" in summary_df.columns and "overall_po_fill_rate" in summary_df.columns:
    fig = px.bar(summary_df, x="manufacturer_name", y="overall_po_fill_rate", color="manufacturer_name")
    show_and_capture_chart("🏭 LFR by Manufacturer", fig, "manufacturer")

# Download CSV
st.subheader("📥 Download Data")
st.download_button("Download Aggregated Data as CSV",
                   data=summary_df.to_csv(index=False),
                   file_name="aggregated_fill_rate.csv",
                   mime="text/csv")

# PDF Generation
def generate_pdf_with_charts(metrics, charts):
    from fpdf import FPDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="📊 Fill Rate Summary", ln=True, align='C')
    pdf.ln(10)
    for label, value in metrics.items():
        pdf.cell(200, 10, txt=f"{label}: {value}", ln=True)

    for title, img_bytes in charts.items():
        pdf.add_page()
        pdf.image(io.BytesIO(img_bytes), w=180)

    pdf_output = io.BytesIO()
    pdf.output(pdf_output)
    return pdf_output.getvalue()

st.subheader("🧾 Generate PDF Summary")
if st.button("Create PDF Summary"):
    kpi_metrics = {
        "Total SKU PO Qty": f"{filtered_df.get('sku_po_qty', pd.Series([0])).sum():,.0f}",
        "Total SKU GRN Qty": f"{filtered_df.get('sku_grn_qty', pd.Series([0])).sum():,.0f}",
        "Vendor Loss A/c": f"{filtered_df.get('Vendor loss A/c', pd.Series([0])).sum():,.0f}",
        "Total PO Lines": f"{filtered_df.get('sku_po_line', pd.Series([0])).sum():,.0f}",
        "Total GRN Lines": f"{filtered_df.get('sku_grn_line', pd.Series([0])).sum():,.0f}",
        "PO Amount": f"{filtered_df.get('po_amount', pd.Series([0])).sum():,.0f}",
        "QFR (%)": f"{safe_mean('sku_level_fill_rate'):.2f}",
        "LFR (%)": f"{safe_mean('overall_po_fill_rate'):.2f}",
    }
    try:
        pdf_data = generate_pdf_with_charts(kpi_metrics, figs)
        st.download_button("Download Full PDF Report", pdf_data, file_name="full_fill_rate_dashboard.pdf")
    except Exception as e:
        st.error(f"PDF generation failed: {e}")
        
