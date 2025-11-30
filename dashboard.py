import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# ============ LOAD DATA ============
df = pd.read_csv("hotel_bookings_cleaned.csv")

month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# ============ SIDEBAR FILTER ============
st.sidebar.title("🔎 Filter Data")

hotel_filter = st.sidebar.multiselect("Pilih Tipe Hotel", df['hotel'].unique(), df['hotel'].unique())
country_filter = st.sidebar.multiselect("Pilih Negara", df['country'].unique(), df['country'].unique())
month_filter = st.sidebar.multiselect("Pilih Bulan", month_order, month_order)

df_filtered = df[
    (df['hotel'].isin(hotel_filter)) &
    (df['country'].isin(country_filter)) &
    (df['arrival_date_month'].isin(month_filter))
]

# ============ HEADER + LOGO ============
st.image("logo.png", width=180)
st.title("🏨 Dashboard Hotel Booking Demand")
st.write("Dashboard interaktif analisis pemesanan hotel menggunakan Streamlit.")
st.markdown("---")

# ============ DATA TABLE ============
st.subheader("📄 Data Table")
st.dataframe(df_filtered, width="stretch")
st.markdown("---")

# ============ STATISTIK ============
st.subheader("📌 Statistik Data (Mean – Median – Modus – Std Dev)")

numeric_df = df_filtered.select_dtypes(include=['int64', 'float64'])
stats = pd.DataFrame({
    "Mean": numeric_df.mean(),
    "Median": numeric_df.median(),
    "Modus": numeric_df.mode().iloc[0],
    "Std Dev": numeric_df.std()
}).round(2)

stats.reset_index(inplace=True)
stats.rename(columns={"index": "Variabel"}, inplace=True)
st.table(stats)

st.markdown("---")

# ============ VISUALISASI ============
st.subheader("📊 Bar Chart — Jumlah Booking per Hotel")
bar_data = df_filtered['hotel'].value_counts().reset_index()
bar_data.columns = ['hotel_type', 'count']
st.plotly_chart(px.bar(bar_data, x='hotel_type', y='count', color='hotel_type'), use_container_width=True)

st.subheader("📈 Line Chart — Tren Booking per Bulan")
monthly_booking = df_filtered.groupby("arrival_date_month").size().reindex(month_order)
st.plotly_chart(px.line(x=month_order, y=monthly_booking.values, markers=True), use_container_width=True)

st.subheader("🥧 Pie Chart — Customer Type")
st.plotly_chart(px.pie(df_filtered, names='customer_type'), use_container_width=True)

st.subheader("📍 Scatter Plot — Lead Time vs ADR")
st.plotly_chart(px.scatter(df_filtered, x='lead_time', y='adr', color='hotel'), use_container_width=True)

st.subheader("🔥 Heatmap — Korelasi Variabel Numerik")
plt.figure(figsize=(8,5))
sns.heatmap(df_filtered.select_dtypes(include="number").corr(), annot=True, cmap="coolwarm")
st.pyplot(plt)

st.markdown("---")
st.write("Dashboard dibuat untuk Tugas UAS Analisis & Visualisasi Data 🎓")
