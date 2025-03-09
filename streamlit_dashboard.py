import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
from typing import Dict, List, Optional
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Set page config
st.set_page_config(
    page_title="Content Creator Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #1F4E78;
        padding: 1rem;
        border-radius: 0.5rem;
    }
    .stMetric > div {
        color: white !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Constants
DEFAULT_INPUT_FILES = {
    "kpi_data": "car_kpi_data.json",
    "kpi_values": "car_kpi_values_2024.json",
    "historical_data": "car_historical_data.json"
}

def load_data() -> tuple[Dict, Dict, Optional[Dict]]:
    """Load data from JSON files"""
    try:
        with open(DEFAULT_INPUT_FILES["kpi_data"], "r") as f:
            kpi_data = json.load(f)
        with open(DEFAULT_INPUT_FILES["kpi_values"], "r") as f:
            values_data = json.load(f)
        try:
            with open(DEFAULT_INPUT_FILES["historical_data"], "r") as f:
                historical_data = json.load(f)
        except FileNotFoundError:
            historical_data = None
        return kpi_data, values_data, historical_data
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return {}, {}, None

def create_performance_metrics(values_data: Dict, cars: List[str], kpis: List[str]):
    """Create key performance metrics section"""
    st.header("🎯 Key Performance Metrics")
    cols = st.columns(len(kpis))
    
    for idx, kpi in enumerate(kpis):
        with cols[idx]:
            values = [values_data[car][kpi] for car in cars if kpi in values_data[car]]
            if values:
                avg_value = sum(values) / len(values)
                st.metric(
                    label=kpi,
                    value=f"{avg_value:.1f}",
                    delta=f"{(max(values) - avg_value):.1f} (Best)",
                )

def create_comparison_chart(values_data: Dict, cars: List[str], selected_kpi: str):
    """Create interactive comparison chart"""
    st.header("📊 Performance Comparison")
    
    data = []
    for car in cars:
        if selected_kpi in values_data[car]:
            data.append({"Car": car, "Value": values_data[car][selected_kpi]})
    
    if data:
        df = pd.DataFrame(data)
        fig = px.bar(
            df,
            x="Car",
            y="Value",
            title=f"{selected_kpi} Comparison",
            color="Value",
            color_continuous_scale="viridis"
        )
        fig.update_layout(
            height=400,
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig, use_container_width=True)

def create_trend_analysis(historical_data: Optional[Dict], cars: List[str], selected_kpi: str):
    """Create trend analysis section"""
    if historical_data and "years" in historical_data and "cars" in historical_data:
        st.header("📈 Trend Analysis")
        
        fig = go.Figure()
        for car in cars:
            if car in historical_data["cars"] and selected_kpi in historical_data["cars"][car]:
                fig.add_trace(go.Scatter(
                    x=historical_data["years"],
                    y=historical_data["cars"][car][selected_kpi],
                    name=car,
                    mode="lines+markers"
                ))
        
        fig.update_layout(
            title=f"{selected_kpi} Trend Over Time",
            height=400,
            showlegend=True,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig, use_container_width=True)

def main():
    # Load data
    kpi_data, values_data, historical_data = load_data()
    
    if not kpi_data or not values_data:
        st.error("Failed to load required data files.")
        return
    
    # Sidebar
    st.sidebar.title("Dashboard Controls")
    cars = kpi_data.get("top_cars_US_2024", [])
    kpis = kpi_data.get("top_5_KPIs", [])
    
    selected_kpi = st.sidebar.selectbox(
        "Select KPI to Display",
        kpis,
        index=0
    )
    
    # Main content
    st.title("Content Creator Analytics Dashboard")
    st.markdown("### Real-time Performance Tracking")
    
    # Create dashboard sections
    create_performance_metrics(values_data, cars, kpis)
    create_comparison_chart(values_data, cars, selected_kpi)
    create_trend_analysis(historical_data, cars, selected_kpi)
    
    # Additional insights
    st.header("💡 Key Insights")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Top Performers")
        for kpi in kpis[:3]:  # Show top 3 KPIs
            top_car = max(cars, key=lambda x: values_data[x].get(kpi, 0))
            st.markdown(f"- **{kpi}**: {top_car} ({values_data[top_car][kpi]})")
    
    with col2:
        st.markdown("#### Areas for Improvement")
        for kpi in kpis[:3]:  # Show bottom 3 KPIs
            bottom_car = min(cars, key=lambda x: values_data[x].get(kpi, 0))
            st.markdown(f"- **{kpi}**: {bottom_car} ({values_data[bottom_car][kpi]})")

if __name__ == "__main__":
    main() 