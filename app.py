import streamlit as st
import pandas as pd
import plotly.express as px
from xhtml2pdf import pisa
import io
st.set_page_config(page_title="Fill Rate Dashboard", layout="wide")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])
def load_data(file):
    df = pd.read_excel(file, sheet_name="Base Data")
    df['sku_level_fill_rate'] = df['sku_level_fill_rate'].str.rstrip('%').astype(float)
    df['overall_po_fill_rate'] = df['overall_po_fill_rate'].str.rstrip('%').astype(float)
    return df

if uploaded_file:
    df = load_data(uploaded_file)
else:
    st.warning("Please upload the file to continue.")
    st.stop()

if "filters" not in st.session_state:
    st.session_state.filters = {
        "manufacturer": [], "category": [],
        "subcategory": [], "location": []
    }

manufacturer = st.sidebar.multiselect("Manufacturer", df["manufacturer_name"].dropna().unique(),
                                      default=st.session_state.filters["manufacturer"])
category = st.sidebar.multiselect("Category", df["category_name"].dropna().unique(),
                                  default=st.session_state.filters["category"])
subcategory = st.sidebar.multiselect("Subcategory", df["subcategory_name"].dropna().unique(),
                                     default=st.session_state.filters["subcategory"])
location = st.sidebar.multiselect("Location", df["wh_name"].dropna().unique(),
                                  default=st.session_state.filters["location"])

st.session_state.filters = {
    "manufacturer": manufacturer, "category": category,
    "subcategory": subcategory, "location": location
}

filtered_df = df.copy()
if manufacturer:
    filtered_df = filtered_df[filtered_df["manufacturer_name"].isin(manufacturer)]
if category:
    filtered_df = filtered_df[filtered_df["category_name"].isin(category)]
if subcategory:
    filtered_df = filtered_df[filtered_df["subcategory_name"].isin(subcategory)]
if location:
    filtered_df = filtered_df[filtered_df["wh_name"].isin(location)]

agg_cols = [
    "sku_po_qty", "sku_grn_qty", "sku_level_fill_rate", "sku_po_line",
    "sku_grn_line", "overall_po_fill_rate", "po_amount", "grn_amount",
    "Vendor loss A/c"
]
summary_df = filtered_df.groupby(["manufacturer_name", "category_name", "subcategory_name", "wh_name"])[agg_cols].sum().reset_index()

st.title("📊 Fill Rate Dashboard")
col1, col2 = st.columns(2)
col1.metric("Total QFR", f"{filtered_df['sku_level_fill_rate'].sum():,.2f}")
col2.metric("Total LFR", f"{filtered_df['overall_po_fill_rate'].sum():,.2f}")

st.subheader("📦 QFR by Category")
cat_fig = px.bar(summary_df, x="category_name", y="sku_level_fill_rate", color="category_name")
st.plotly_chart(cat_fig, use_container_width=True)

st.subheader("🔍 QFR by Subcategory")
sub_fig = px.bar(summary_df, x="subcategory_name", y="sku_level_fill_rate", color="subcategory_name")
st.plotly_chart(sub_fig, use_container_width=True)

st.subheader("🏭 LFR by Manufacturer")
man_fig = px.bar(summary_df, x="manufacturer_name", y="overall_po_fill_rate", color="manufacturer_name")
st.plotly_chart(man_fig, use_container_width=True)

st.subheader("📥 Download Data")
st.download_button("Download Aggregated Data as CSV",
                   data=summary_df.to_csv(index=False),
                   file_name="aggregated_fill_rate.csv",
                   mime="text/csv")

def convert_html_to_pdf(source_html):
    output = io.BytesIO()
    pisa_status = pisa.CreatePDF(source_html, dest=output)
    return output if not pisa_status.err else None

st.subheader("🧾 Generate PDF Summary")
if st.button("Create PDF Summary"):
    html = f"""
    <html><body>
    <h2>📊 Fill Rate Summary</h2>
    <p><strong>Total QFR:</strong> {filtered_df['sku_level_fill_rate'].sum():,.2f}</p>
    <p><strong>Total LFR:</strong> {filtered_df['overall_po_fill_rate'].sum():,.2f}</p>
    </body></html>
    """
    pdf_file = convert_html_to_pdf(html)
    if pdf_file:
        st.download_button("Download Summary PDF", pdf_file, file_name="fill_rate_summary.pdf")
    else:
        st.error("Failed to generate PDF")
