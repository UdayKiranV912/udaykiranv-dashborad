import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

st.set_page_config(page_title="Universal Fill Rate Dashboard", layout="wide")

# File upload
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel File", type=["xlsx", "xls"])

# Cache Excel loader
@st.cache_data
def load_excel(file):
    excel = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=excel.sheet_names[0])

    # Convert percentage columns
    for col in ["sku_level_fill_rate", "overall_po_fill_rate"]:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].str.rstrip('%').astype(float)

    return df, excel.sheet_names

if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

df, sheet_names = load_excel(uploaded_file)

# Sidebar sheet selection (if multiple sheets exist)
if len(sheet_names) > 1:
    sheet_selected = st.sidebar.selectbox("📑 Select Sheet", sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_selected)

# Filter section
filter_columns = {
    "Manufacturer": "manufacturer_name",
    "Category": "category_name",
    "Subcategory": "subcategory_name",
    "Location": "wh_name"
}

st.sidebar.markdown("### 🔍 Filters")
for label, col in filter_columns.items():
    if col in df.columns:
        selected = st.sidebar.multiselect(label, df[col].dropna().unique().tolist())
        if selected:
            df = df[df[col].isin(selected)]

# Title
st.title("📊 Excel Power BI–Style Fill Rate Dashboard")

# KPI Calculation
kpi_cols = {
    "sku_po_qty": "Total SKU PO Qty",
    "sku_grn_qty": "Total SKU GRN Qty",
    "sku_po_line": "Total PO Lines",
    "sku_grn_line": "Total GRN Lines",
    "po_amount": "PO Amount",
    "grn_amount": "GRN Amount",
    "Vendor loss A/c": "Vendor Loss A/c",
}

metrics = {}
col1, col2, col3 = st.columns(3)
for i, (col, label) in enumerate(kpi_cols.items()):
    if col in df.columns:
        value = df[col].sum()
        metrics[label] = value
        [col1, col2, col3][i % 3].metric(label, f"{value:,.0f}")

# Fill Rate KPIs
qfr = df["sku_level_fill_rate"].mean() if "sku_level_fill_rate" in df.columns else None
lfr = df["overall_po_fill_rate"].mean() if "overall_po_fill_rate" in df.columns else None

col1, col2 = st.columns(2)
if qfr is not None:
    col1.metric("QFR (%)", f"{qfr:.2f}%")
    metrics["QFR (%)"] = qfr
if lfr is not None:
    col2.metric("LFR (%)", f"{lfr:.2f}%")
    metrics["LFR (%)"] = lfr

# Charts
def plot_chart(x_col, y_col, title):
    fig = px.bar(df, x=x_col, y=y_col, color=x_col)
    st.subheader(title)
    st.plotly_chart(fig, use_container_width=True)
    return fig

chart_images = {}

if "sku_level_fill_rate" in df.columns:
    for group_col in ["category_name", "subcategory_name"]:
        if group_col in df.columns:
            chart_images[group_col] = plot_chart(group_col, "sku_level_fill_rate", f"📦 QFR by {group_col.replace('_name','').capitalize()}")

if "overall_po_fill_rate" in df.columns and "manufacturer_name" in df.columns:
    chart_images["manufacturer_name"] = plot_chart("manufacturer_name", "overall_po_fill_rate", "🏭 LFR by Manufacturer")

# Download filtered data
st.subheader("📥 Download Filtered Data")
st.download_button("Download CSV", df.to_csv(index=False), "filtered_data.csv", "text/csv")

# PDF summary
def generate_pdf(metrics, charts):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, "📊 Fill Rate Summary", ln=True, align="C")
    pdf.set_font("Arial", "", 12)

    for label, value in metrics.items():
        if isinstance(value, float):
            pdf.cell(190, 10, f"{label}: {value:,.2f}", ln=True)
        else:
            pdf.cell(190, 10, f"{label}: {value:,}", ln=True)

    for name, fig in charts.items():
        img_bytes = fig.to_image(format="png")
        img_path = f"/tmp/{name}.png"
        with open(img_path, "wb") as f:
            f.write(img_bytes)
        pdf.add_page()
        pdf.image(img_path, x=10, y=30, w=180)

    output = io.BytesIO()
    pdf.output(output)
    return output.getvalue()

st.subheader("🧾 Download PDF Report")
if st.button("Generate PDF Summary"):
    try:
        pdf_bytes = generate_pdf(metrics, chart_images)
        st.download_button("📄 Download PDF", pdf_bytes, "dashboard_summary.pdf")
    except Exception as e:
        st.error(f"PDF generation failed: {e}")
