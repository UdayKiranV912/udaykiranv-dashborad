import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io

st.set_page_config(page_title="Excel Power BI Report", layout="wide")

# Upload Excel
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel File", type=["xlsx", "xls"])

@st.cache_data
def load_excel(file):
    excel = pd.ExcelFile(file, engine="openpyxl")
    return excel

if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

excel = load_excel(uploaded_file)

# Select sheet
sheet_name = st.sidebar.selectbox("📑 Select Sheet", excel.sheet_names)
df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")

# Clean dataframe: drop fully empty columns
df = df.dropna(axis=1, how="all")

st.title(f"📊 Power BI–Style Dashboard ({sheet_name})")

# --- Filters ---
filterable_cols = [col for col in df.columns if df[col].nunique() < 25 and df[col].dtype == object]
filtered_df = df.copy()

st.sidebar.markdown("## 🔍 Filters")
for col in filterable_cols:
    options = ["All"] + sorted(df[col].dropna().unique().tolist())
    choice = st.sidebar.selectbox(f"{col}", options)
    if choice != "All":
        filtered_df = filtered_df[filtered_df[col] == choice]

# --- KPIs ---
st.subheader("📌 Key Metrics")
num_cols = filtered_df.select_dtypes(include="number").columns.tolist()

if len(num_cols) > 0:
    col1, col2, col3 = st.columns(3)
    metrics = {}
    for i, col in enumerate(num_cols[:6]):  # show first 6 metrics
        value = filtered_df[col].sum()
        metrics[col] = value
        [col1, col2, col3][i % 3].metric(col, f"{value:,.2f}")
else:
    st.info("No numeric columns found to show KPIs.")

# --- Charts ---
st.subheader("📊 Visualizations")

chart_images = {}

# Bar chart (numeric by category columns)
cat_cols = [c for c in df.columns if df[c].dtype == object and df[c].nunique() < 20]

for cat in cat_cols[:3]:  # limit to 3 for clarity
    for num in num_cols[:1]:  # first numeric col
        fig = px.bar(filtered_df, x=cat, y=num, color=cat, text_auto=".2f")
        st.plotly_chart(fig, use_container_width=True)
        chart_images[f"{cat}_{num}_bar"] = fig

# Pie chart
if len(cat_cols) > 0 and len(num_cols) > 0:
    fig = px.pie(filtered_df, names=cat_cols[0], values=num_cols[0], title=f"{num_cols[0]} by {cat_cols[0]}")
    st.plotly_chart(fig, use_container_width=True)
    chart_images[f"{cat_cols[0]}_{num_cols[0]}_pie"] = fig

# --- Export Section ---
st.subheader("📥 Export Options")

st.download_button("⬇️ Download Filtered Data (CSV)", filtered_df.to_csv(index=False), "filtered_data.csv", "text/csv")

# PDF Summary
def generate_pdf(metrics, charts):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, "📊 Excel Report Summary", ln=True, align="C")
    pdf.set_font("Arial", "", 12)

    for label, value in metrics.items():
        pdf.cell(190, 10, f"{label}: {value:,.2f}", ln=True)

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
