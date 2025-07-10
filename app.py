import streamlit as st
import pandas as pd
import plotly.express as px
from xhtml2pdf import pisa
import io

st.set_page_config(page_title="Fill Rate Dashboard", layout="wide")
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
    st.warning("Please upload the file to continue.")
    st.stop()

if "filters" not in st.session_state:
    st.session_state.filters = {
        "manufacturer": [], "category": [],
        "subcategory": [], "location": []
    }

manufacturer_options = df["manufacturer_name"].dropna().unique().tolist()
manufacturer_default = [m for m in st.session_state.filters["manufacturer"] if m in manufacturer_options]
manufacturer = st.sidebar.multiselect("Manufacturer", manufacturer_options, default=manufacturer_default)

category_options = df["category_name"].dropna().unique().tolist()
category_default = [c for c in st.session_state.filters["category"] if c in category_options]
category = st.sidebar.multiselect("Category", category_options, default=category_default)

subcategory_options = df["subcategory_name"].dropna().unique().tolist()
subcategory_default = [s for s in st.session_state.filters["subcategory"] if s in subcategory_options]
subcategory = st.sidebar.multiselect("Subcategory", subcategory_options, default=subcategory_default)

location_options = df["wh_name"].dropna().unique().tolist()
location_default = [l for l in st.session_state.filters["location"] if l in location_options]
location = st.sidebar.multiselect("Location", location_options, default=location_default)

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

expected_cols = [
    "sku_po_qty", "sku_grn_qty", "sku_level_fill_rate", "sku_po_line",
    "sku_grn_line", "overall_po_fill_rate", "po_amount", "grn_amount",
    "Vendor loss A/c"
]
agg_cols = [col for col in expected_cols if col in filtered_df.columns]
group_cols = ["manufacturer_name", "category_name", "subcategory_name", "wh_name"]
missing_cols = [col for col in group_cols if col not in filtered_df.columns]

if missing_cols:
    st.error(f"The following required columns are missing: {', '.join(missing_cols)}")
    st.stop()

summary_df = filtered_df.groupby(group_cols)[agg_cols].sum().reset_index()

total_sku_po_qty = filtered_df["sku_po_qty"].sum()
total_sku_grn_qty = filtered_df["sku_grn_qty"].sum()
total_sku_po_line = filtered_df["sku_po_line"].sum()
total_sku_grn_line = filtered_df["sku_grn_line"].sum()
total_po_amount = filtered_df["po_amount"].sum()
total_grn_amount = filtered_df["grn_amount"].sum()
total_vendor_loss = filtered_df["Vendor loss A/c"].sum() if "Vendor loss A/c" in filtered_df.columns else 0.0
total_qfr = filtered_df["sku_level_fill_rate"].mean()
total_lfr = filtered_df["overall_po_fill_rate"].mean()

st.title("📊 Fill Rate Dashboard")
c1, c2, c3 = st.columns(3)
c1.metric("Total SKU PO Qty", f"{total_sku_po_qty:,.0f}")
c2.metric("Total SKU GRN Qty", f"{total_sku_grn_qty:,.0f}")
c3.metric("Vendor Loss A/c", f"{total_vendor_loss:,.0f}")

c4, c5, c6 = st.columns(3)
c4.metric("Total PO Lines", f"{total_sku_po_line:,.0f}")
c5.metric("Total GRN Lines", f"{total_sku_grn_line:,.0f}")
c6.metric("PO Amount", f"{total_po_amount:,.0f}")

c7, c8 = st.columns(2)
c7.metric("QFR (%)", f"{total_qfr:.2f}%", delta=None)
c8.metric("LFR (%)", f"{total_lfr:.2f}%", delta=None)

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
    <p><strong>Total QFR (%):</strong> {total_qfr:.2f}</p>
    <p><strong>Total LFR (%):</strong> {total_lfr:.2f}</p>
    <p><strong>Total SKU PO Qty:</strong> {total_sku_po_qty:,.0f}</p>
    <p><strong>Total SKU GRN Qty:</strong> {total_sku_grn_qty:,.0f}</p>
    <p><strong>Total PO Lines:</strong> {total_sku_po_line:,.0f}</p>
    <p><strong>Total GRN Lines:</strong> {total_sku_grn_line:,.0f}</p>
    <p><strong>Total PO Amount:</strong> {total_po_amount:,.0f}</p>
    <p><strong>Total GRN Amount:</strong> {total_grn_amount:,.0f}</p>
    <p><strong>Vendor Loss A/c:</strong> {total_vendor_loss:,.0f}</p>
    </body></html>
    """
    pdf_file = convert_html_to_pdf(html)
    if pdf_file:
        st.download_button("Download Summary PDF", pdf_file, file_name="fill_rate_summary.pdf")
    else:
        st.error("Failed to generate PDF")
