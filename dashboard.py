import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image

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
st.image("logo.png", width=180)
st.title("🏨 Dashboard Hotel Booking Demand")
st.write("Dashboard interaktif analisis pemesanan hotel menggunakan Streamlit.")
st.markdown("---")

# ============= DATA TABLE =============
st.subheader("📄 Data Table")
st.dataframe(df_filtered, use_container_width=True)
st.markdown("---")

# ============= STATISTIK DESKRIPTIF =============
st.subheader("📊 Statistik Data (Mean – Median – Modus – Std Dev)")

if df_filtered.empty:
    st.warning("⚠ Data kosong karena filter. Silakan ubah filter.")
else:
    numeric_df = df_filtered.select_dtypes(include=['number'])
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
st.subheader("📊 Bar Chart — Jumlah Booking per Hotel")
bar_data = df_filtered['hotel'].value_counts().reset_index()
bar_data.columns = ['hotel_type', 'count']
bar_fig = px.bar(bar_data, x='hotel_type', y='count', color='hotel_type')
st.plotly_chart(bar_fig)

st.subheader("📈 Line Chart — Tren Booking per Bulan")
monthly_booking = df_filtered.groupby("arrival_date_month").size().reindex(month_order)
line_fig = px.line(x=month_order, y=monthly_booking.values, markers=True)
st.plotly_chart(line_fig)

st.subheader("🥧 Pie Chart — Customer Type")
pie_fig = px.pie(df_filtered, names='customer_type')
st.plotly_chart(pie_fig)

st.subheader("📍 Scatter Plot — Lead Time vs ADR")
scatter_fig = px.scatter(df_filtered, x='lead_time', y='adr', color='hotel')
st.plotly_chart(scatter_fig)

st.subheader("🔥 Heatmap — Korelasi Variabel Numerik")
plt.figure(figsize=(8,5))
sns.heatmap(df_filtered.select_dtypes(include="number").corr(), annot=True, cmap="coolwarm")
st.pyplot(plt)
plt.close()

st.markdown("---")

# ============= DOWNLOAD DATA =============
st.download_button(
    label="📥 Download Filtered Data sebagai Excel",
    data=df_filtered.to_excel("filtered_data.xlsx", index=False),
    file_name="filtered_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")
st.write("Dashboard dibuat untuk Tugas UAS Analisis & Visualisasi Data 🎓")
