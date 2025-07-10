import streamlit as st
import pandas as pd
import plotly.express as px
from xhtml2pdf import pisa
import io

st.set_page_config(page_title="Fill Rate Dashboard", layout="wide")

# Upload Excel
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel File", type=["xlsx"])

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="Base Data")
    df['sku_level_fill_rate'] = df['sku_level_fill_rate'].str.rstrip('%').astype(float)
    df['overall_po_fill_rate'] = df['overall_po_fill_rate'].str.rstrip('%').astype(float)
    return df

if uploaded_file:
    df = load_data(uploaded_file)
else:
    st.warning("Please upload 'Fill Rate_2025-06-30.xlsx'")
    st.stop()

# Filter memory
if "filters" not in st.session_state:
    st.session_state.filters = {
        "manufacturer": [], "category": [],
        "subcategory": [], "location": []
    }

# Sidebar Filters
manufacturer = st.sidebar.multiselect("Manufacturer", df["brand_name"].dropna().unique(),
                                      default=st.session_state.filters["manufacturer"])
category = st.sidebar.multiselect("Category", df["category_name"].dropna().unique(),
                                  default=st.session_state.filters["category"])
subcategory = st.sidebar.multiselect("Subcategory", df["subcategory_name"].dropna().unique(),
                                     default=st.session_state.filters["subcategory"])
location = st.sidebar.multiselect("Location", df["vendorcode"].dropna().unique(),
                                  default=st.session_state.filters["location"])

st.session_state.filters = {
    "manufacturer": manufacturer, "category": category,
    "subcategory": subcategory, "location": location
}

# Apply Filters
filtered_df = df.copy()
if manufacturer:
    filtered_df = filtered_df[filtered_df["brand_name"].isin(manufacturer)]
if category:
    filtered_df = filtered_df[filtered_df["category_name"].isin(category)]
if subcategory:
    filtered_df = filtered_df[filtered_df["subcategory_name"].isin(subcategory)]
if location:
    filtered_df = filtered_df[filtered_df["vendorcode"].isin(location)]

# KPIs
st.title("📊 Fill Rate Dashboard")
col1, col2 = st.columns(2)
col1.metric("Average QFR", f"{filtered_df['sku_level_fill_rate'].mean():.2f}%")
col2.metric("Average LFR", f"{filtered_df['overall_po_fill_rate'].mean():.2f}%")

# Charts
st.subheader("📦 Fill Rate by Category")
cat_fig = px.bar(filtered_df, x="category_name", y="sku_level_fill_rate", color="category_name")
st.plotly_chart(cat_fig, use_container_width=True)

st.subheader("🔍 Fill Rate by Subcategory")
sub_fig = px.bar(filtered_df, x="subcategory_name", y="sku_level_fill_rate", color="subcategory_name")
st.plotly_chart(sub_fig, use_container_width=True)

st.subheader("🏭 Fill Rate by Manufacturer")
man_fig = px.bar(filtered_df, x="brand_name", y="overall_po_fill_rate", color="brand_name")
st.plotly_chart(man_fig, use_container_width=True)

# CSV Export
st.subheader("📥 Download Data")
st.download_button("Download Filtered Data as CSV",
                   data=filtered_df.to_csv(index=False),
                   file_name="filtered_data.csv",
                   mime="text/csv")

# PDF Summary using xhtml2pdf
def convert_html_to_pdf(source_html):
    output = io.BytesIO()
    pisa_status = pisa.CreatePDF(source_html, dest=output)
    return output if not pisa_status.err else None

st.subheader("🧾 Generate PDF Summary")
if st.button("Create PDF Summary"):
    html = f"""
    <html><body>
    <h2>📊 Fill Rate Summary</h2>
    <p><strong>Average QFR:</strong> {filtered_df['sku_level_fill_rate'].mean():.2f}%</p>
    <p><strong>Average LFR:</strong> {filtered_df['overall_po_fill_rate'].mean():.2f}%</p>
    </body></html>
    """
    pdf_file = convert_html_to_pdf(html)
    if pdf_file:
        st.download_button("Download Summary PDF", pdf_file, file_name="fill_rate_summary.pdf")
    else:
        st.error("Failed to generate PDF")
