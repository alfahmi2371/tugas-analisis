import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from openpyxl import Workbook

# ============= LOAD DATA =============
df = pd.read_csv("hotel_bookings_cleaned.csv")

month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# ============= SIDEBAR FILTERS =============
st.sidebar.title("🔎 Filter Data")

hotel_filter = st.sidebar.multiselect(
    "Pilih Tipe Hotel",
    options=df['hotel'].unique(),
    default=df['hotel'].unique()
)

country_filter = st.sidebar.multiselect(
    "Pilih Negara",
    options=df['country'].unique(),
    default=df['country'].unique()
)

month_filter = st.sidebar.multiselect(
    "Pilih Bulan",
    options=month_order,
    default=month_order
)

df_filtered = df[
    (df['hotel'].isin(hotel_filter)) &
    (df['country'].isin(country_filter)) &
    (df['arrival_date_month'].isin(month_filter))
]

# ============= HEADER + LOGO =============
st.image("logo.png", width=190)
st.title("🏨 Dashboard Hotel Booking Demand")
st.write("Dashboard interaktif analisis pemesanan hotel menggunakan Streamlit.")
st.markdown("---")

# ============= DATA TABLE =============
st.subheader("📄 Data Table")
st.dataframe(df_filtered, use_container_width=True)
st.markdown("---")

# ============= STATISTIK DESKRIPTIF =============
st.subheader("📊 Statistik Data (Mean – Median – Modus – Std Dev)")

numeric_df = df_filtered.select_dtypes(include=['int64', 'float64', 'number'])

if numeric_df.empty:
    st.warning("⚠ Tidak ada kolom numerik untuk dihitung statistik.")
else:
    stats = pd.DataFrame({
        "Mean": numeric_df.mean(),
        "Median": numeric_df.median(),
        "Modus": numeric_df.mode().iloc[0],
        "Std Dev": numeric_df.std()
    }).round(2)

    stats.reset_index(inplace=True)
    stats.rename(columns={"index": "Variabel"}, inplace=True)

    st.dataframe(stats, use_container_width=True)

st.markdown("---")

# ============= VISUALISASI DATA =============

# 1. Bar chart
st.subheader("📊 Bar Chart — Jumlah Booking per Hotel")
bar_data = df_filtered['hotel'].value_counts().reset_index()
bar_data.columns = ['hotel_type', 'count']
bar_fig = px.bar(bar_data, x='hotel_type', y='count', color='hotel_type')
st.plotly_chart(bar_fig)

# 2. Line chart
st.subheader("📈 Line Chart — Tren Booking per Bulan")
monthly_booking = df_filtered.groupby("arrival_date_month").size().reindex(month_order)
line_fig = px.line(x=month_order, y=monthly_booking.values, markers=True)
st.plotly_chart(line_fig)

# 3. Pie chart
st.subheader("🥧 Pie Chart — Customer Type")
pie_fig = px.pie(df_filtered, names='customer_type')
st.plotly_chart(pie_fig)

# 4. Scatter plot
st.subheader("📍 Scatter Plot — Lead Time vs ADR")
scatter_fig = px.scatter(df_filtered, x='lead_time', y='adr', color='hotel')
st.plotly_chart(scatter_fig)

# 5. Heatmap
st.subheader("🔥 Heatmap — Korelasi Variabel Numerik")
plt.figure(figsize=(8,5))
sns.heatmap(df_filtered.select_dtypes(include="number").corr(), annot=True, cmap="coolwarm")
plt.tight_layout()
st.pyplot(plt)

st.markdown("---")

# ============= EXPORT EXCEL =============
st.subheader("⬇ Download Semua Data sebagai Excel")

def export_to_excel():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Filtered Data"

    for i, col in enumerate(df_filtered.columns, 1):
        ws1.cell(row=1, column=i).value = col
    for r, row in enumerate(df_filtered.values, 2):
        for c, v in enumerate(row, 1):
            ws1.cell(row=r, column=c).value = v

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

excel_file = export_to_excel()

st.download_button(
    label="📥 Download Data",
    data=excel_file,
    file_name="dashboard_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")
st.write("Dashboard dibuat untuk Tugas UAS Analisis & Visualisasi Data 🎓")
