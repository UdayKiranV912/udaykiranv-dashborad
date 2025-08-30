import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

st.set_page_config(page_title="Excel Power BI Dashboard", layout="wide")

# Upload Excel
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel File", type=["xlsx", "xls"])

@st.cache_data
def load_excel(file):
    excel = pd.ExcelFile(file, engine="openpyxl")
    df = pd.read_excel(file, sheet_name=excel.sheet_names[0], engine="openpyxl")

    # Convert percentage strings to numbers
    for col in ["sku_level_fill_rate", "overall_po_fill_rate"]:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].str.rstrip('%').astype(float)

    return df, excel.sheet_names

if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

df, sheet_names = load_excel(uploaded_file)

# --- Report Section ---
st.title("📊 Power BI–Style Fill Rate Dashboard")

# KPI Summary
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

# --- Charts ---
def plot_chart(x_col, y_col, title):
    fig = px.bar(df, x=x_col, y=y_col, color=x_col, text_auto=".2f")
    st.subheader(title)
    st.plotly_chart(fig, use_container_width=True)
    return fig

chart_images = {}

if "sku_level_fill_rate" in df.columns:
    if "category_name" in df.columns:
        chart_images["category"] = plot_chart("category_name", "sku_level_fill_rate", "📦 QFR by Category")
    if "subcategory_name" in df.columns:
        chart_images["subcategory"] = plot_chart("subcategory_name", "sku_level_fill_rate", "📦 QFR by Subcategory")

if "overall_po_fill_rate" in df.columns and "manufacturer_name" in df.columns:
    chart_images["manufacturer"] = plot_chart("manufacturer_name", "overall_po_fill_rate", "🏭 LFR by Manufacturer")

# --- Filter Section (AFTER Report) ---
st.sidebar.markdown("## 🔍 Filter Data")

filter_columns = {
    "Manufacturer": "manufacturer_name",
    "Category": "category_name",
    "Subcategory": "subcategory_name",
    "Location": "wh_name"
}

filtered_df = df.copy()
for label, col in filter_columns.items():
    if col in df.columns:
        selected = st.sidebar.selectbox(f"{label}", ["All"] + df[col].dropna().unique().tolist())
        if selected != "All":
            filtered_df = filtered_df[filtered_df[col] == selected]

# --- Export Section ---
st.subheader("📥 Export Options")

# CSV Download
st.download_button("⬇️ Download Filtered Data (CSV)", filtered_df.to_csv(index=False), "filtered_data.csv", "text/csv")

# PDF Generation
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

if st.button("🧾 Generate PDF Report"):
    try:
        pdf_bytes = generate_pdf(metrics, chart_images)
        st.download_button("📄 Download PDF", pdf_bytes, "dashboard_summary.pdf")
    except Exception as e:
        st.error(f"PDF generation failed: {e}")
