import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from attribute_analyzer import (
    find_similar_attributes,
    find_differences,
    export_to_excel
)
import os
from datetime import datetime

st.set_page_config(page_title="Attribute Similarity Analyzer", layout="wide")

# Initialize session state
if 'similarity_threshold' not in st.session_state:
    st.session_state.similarity_threshold = 90
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = None
if 'original_file' not in st.session_state:
    st.session_state.original_file = None

st.title("Attribute Similarity Analyzer")
st.write("Upload an Excel file to analyze similar attributes and standardize them.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Save the file temporarily
    with open("temp_upload.xlsx", "wb") as f:
        f.write(uploaded_file.getvalue())
    st.session_state.original_file = "temp_upload.xlsx"
    
    # Load the data
    df = pd.read_excel(uploaded_file)
    
    # Dataset Overview
    with st.expander("Dataset Overview", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Rows", len(df))
            st.metric("Total Columns", len(df.columns))
        with col2:
            st.metric("Total Attributes", len(df.iloc[:, 0].unique()))
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024**2:.2f} MB")
    
    # Analysis Settings
    with st.expander("Analysis Settings", expanded=True):
        st.slider(
            "Similarity Threshold (%)",
            min_value=0,
            max_value=100,
            value=st.session_state.similarity_threshold,
            key="similarity_threshold"
        )
        
        if st.button("Run Analysis", type="primary"):
            with st.spinner("Analyzing attributes..."):
                # Run the analysis
                similar_groups = find_similar_attributes(
                    st.session_state.original_file,
                    similarity_threshold=st.session_state.similarity_threshold
                )
                st.session_state.analysis_results = similar_groups
                
                # Generate Excel report
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = export_to_excel(
                    similar_groups,
                    st.session_state.similarity_threshold,
                    st.session_state.original_file
                )
                
                # Provide download link
                with open(output_file, 'rb') as f:
                    st.download_button(
                        label="📥 Download Analysis Report",
                        data=f,
                        file_name=f"similarity_analysis_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
    # Results Display
    if st.session_state.analysis_results:
        with st.expander("Analysis Results", expanded=True):
            # Convert results to a more streamlit-friendly format
            results_data = []
            for base_attr, matches in st.session_state.analysis_results.items():
                for match, score in matches:
                    diffs = find_differences(base_attr, match)
                    diff_text = " vs ".join([f"{d1} → {d2}" for d1, d2 in diffs if d1 or d2])
                    results_data.append({
                        "Base Attribute": base_attr,
                        "Similar Attribute": match,
                        "Similarity %": score,
                        "Differences": diff_text
                    })
            
            if results_data:
                results_df = pd.DataFrame(results_data)
                
                # Display results in an interactive table
                st.dataframe(
                    results_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Base Attribute": st.column_config.TextColumn("Base Attribute"),
                        "Similar Attribute": st.column_config.TextColumn("Similar Attribute"),
                        "Similarity %": st.column_config.NumberColumn(
                            "Similarity %",
                            format="%.1f%%"
                        ),
                        "Differences": st.column_config.TextColumn(
                            "Differences",
                            width="large",
                        ),
                    }
                )
                
                # Add visualizations
                st.subheader("Similarity Distribution")
                similarities = results_df["Similarity %"].tolist()
                fig = px.histogram(
                    x=similarities,
                    nbins=20,
                    labels={'x': 'Similarity %', 'y': 'Count'},
                    title='Distribution of Similarity Scores'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Summary statistics
                st.subheader("Summary Statistics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Similar Pairs", len(results_df))
                with col2:
                    avg_similarity = np.mean(similarities) if similarities else 0
                    st.metric("Average Similarity", f"{avg_similarity:.1f}%")
                with col3:
                    st.metric("100% Matches", len([s for s in similarities if s == 100]))
            else:
                st.info("No similar attributes found with the current threshold.")

# Cleanup temporary files
if os.path.exists("temp_upload.xlsx"):
    os.remove("temp_upload.xlsx")
