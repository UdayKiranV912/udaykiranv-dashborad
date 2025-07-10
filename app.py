import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
import pdfkit
import yagmail

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

# Filters
if "filters" not in st.session_state:
    st.session_state.filters = {
        "manufacturer": [], "category": [],
        "subcategory": [], "location": []
    }

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

# Export chart
st.subheader("📤 Export Chart")
chart_map = {
    "Category Chart": cat_fig,
    "Subcategory Chart": sub_fig,
    "Manufacturer Chart": man_fig
}
chart_choice = st.selectbox("Choose chart", list(chart_map.keys()))
format_choice = st.selectbox("Format", ["PNG", "PDF"])
if st.button("Export Chart"):
    img_bytes = pio.to_image(chart_map[chart_choice], format=format_choice.lower(), width=1000, height=600)
    st.download_button(f"Download {chart_choice}", data=img_bytes,
                       file_name=f"{chart_choice.replace(' ', '_')}.{format_choice.lower()}")

# PDF Summary
st.subheader("🧾 Generate PDF Summary")
if st.button("Create PDF"):
    html = f"""
    <html><body>
    <h2>Fill Rate Summary</h2>
    <p><strong>Avg QFR:</strong> {filtered_df['sku_level_fill_rate'].mean():.2f}%</p>
    <p><strong>Avg LFR:</strong> {filtered_df['overall_po_fill_rate'].mean():.2f}%</p>
    </body></html>
    """
    with open("report.html", "w") as f:
        f.write(html)
    pdfkit.from_file("report.html", "summary.pdf")
    with open("summary.pdf", "rb") as f:
        st.download_button("Download PDF Summary", f, file_name="summary.pdf")