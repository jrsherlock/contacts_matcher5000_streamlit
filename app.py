import streamlit as st
import pandas as pd
import json
import os
from fuzzywuzzy import process, fuzz
import plotly.express as px
import io
import base64

# Set page configuration
st.set_page_config(
    page_title="Contacts Matcher 5000",
    page_icon="📇",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Add logo to sidebar
with st.sidebar:
    st.image("assets/ProCircular-Logo-Black.png", width=200)
    st.divider()
    
    # File uploads with clearer sections
    st.markdown("### 📋 Upload Your Contact Lists")
    
    st.markdown("""
    #### 1️⃣ Ideal Contact List
    Upload your master/target contact list (e.g., from your CRM)
    """)
    ideal_file = st.file_uploader("Choose your ideal contacts CSV file", type=["csv"])
    
    st.markdown("""
    #### 2️⃣ Source Contact Lists
    Upload one or more contact lists to compare against your ideal list
    """)
    source_files = st.file_uploader("Choose source contacts CSV file(s)", type=["csv"], accept_multiple_files=True)

# App title and description
st.title("Contacts Matcher 5000")
st.markdown("""
A powerful tool for analyzing and comparing contact databases across different sources.
Identify overlapping company relationships and their associated contacts with ease.
""")

# Add welcome message and instructions
st.markdown("""
## Welcome to Contacts Matcher 5000!

This app allows you to compare contact databases across different sources and identify overlapping company relationships.

### Getting Started
1. Upload your ideal contact list CSV file
2. Upload one or more source contact list CSV files
3. Configure the matching settings
4. View and download the results

**Note:** This app processes data only in your browser. No data is stored on our servers.

### Expected CSV Format
Your CSV files should contain columns for:
- Company Name
- First Name
- Last Name
- Email Address
- Job Title/Position
- Department (optional)
- LinkedIn URL (optional)

Don't worry if your column names are different - you'll be able to map them in the interface.
""")

# Default settings
DEFAULT_SETTINGS = {
    "thresholds": {
        "company_name": 85,
        "person_name": 85,
        "email": 100,
        "title": 70,
        "department": 70
    },
    "column_mapping": {
        "First Name": "First Name",
        "Last Name": "Last Name",
        "URL": "LinkedIn Contact Profile URL",
        "Email Address": "Email Address",
        "Company": "Company Name",
        "Position": "Job Title",
        "Department": "Job Function"
    }
}

# Function to normalize company names
def normalize_company_name(name):
    if not isinstance(name, str):
        return ""
    
    # Convert to lowercase
    name = name.lower()
    
    # Remove common suffixes
    suffixes = [" inc", " inc.", " incorporated", " llc", " ltd", " limited", " corp", " corp.", " corporation"]
    for suffix in suffixes:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
    
    # Remove punctuation and extra whitespace
    import re
    name = re.sub(r'[^\w\s]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    
    return name

# Function to try reading CSV files with different encodings and delimiters
@st.cache_data
def try_read_csv(uploaded_file):
    """Try to read a CSV file with different encodings and delimiters"""
    encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252', 'macroman']
    delimiters = [',', ';', '\t']
    
    # Get the file content
    content = uploaded_file.getvalue()
    
    for encoding in encodings:
        for delimiter in delimiters:
            try:
                df = pd.read_csv(io.BytesIO(content), encoding=encoding, delimiter=delimiter, dtype=str)
                
                # Clean up column names
                df.columns = [col.strip() for col in df.columns]
                
                # Drop unnamed columns
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                
                return df
            except Exception as e:
                continue
    
    st.error(f"Failed to read {uploaded_file.name} with all encodings and delimiters")
    return None

# Function to match companies
def match_companies(ideal_df, source_df, company_threshold, column_mapping):
    """Match companies between ideal and source dataframes"""
    matches = []
    
    # Debug prints
    st.write("Ideal DataFrame columns:", ideal_df.columns.tolist())
    st.write("Source DataFrame columns:", source_df.columns.tolist())
    st.write("Column mapping:", column_mapping)
    
    # Get company name columns - look for mapped columns first, then try defaults
    source_company_col = next((col for col in source_df.columns if col in ["Company", "Company Name"]), None)
    if not source_company_col:
        st.error("Could not find a company name column in the source file. Please ensure it contains 'Company' or 'Company Name'.")
        return []
    
    ideal_company_col = next((col for col in ideal_df.columns if col in ["Company", "Company Name"]), None)
    if not ideal_company_col:
        st.error("Could not find a company name column in the ideal file. Please ensure it contains 'Company' or 'Company Name'.")
        return []
    
    st.write("Using columns:", ideal_company_col, "and", source_company_col)
    
    # Normalize company names
    ideal_companies = {idx: normalize_company_name(name) for idx, name in enumerate(ideal_df[ideal_company_col])}
    source_companies = {idx: normalize_company_name(name) for idx, name in enumerate(source_df[source_company_col])}
    
    # Create a list of normalized ideal company names for fuzzy matching
    ideal_company_names = list(ideal_companies.values())
    
    # For each source company, find the best match in the ideal list
    for source_idx, source_company in source_companies.items():
        if not source_company:
            continue
            
        # Find the best match
        best_match, score = process.extractOne(source_company, ideal_company_names, scorer=fuzz.token_sort_ratio)
        
        if score >= company_threshold:
            # Find the index of the best match in the ideal list
            ideal_idx = [idx for idx, name in ideal_companies.items() if name == best_match][0]
            
            # Get the original company names
            original_ideal_name = ideal_df.iloc[ideal_idx][ideal_company_col]
            original_source_name = source_df.iloc[source_idx][source_company_col]
            
            matches.append({
                "ideal_idx": ideal_idx,
                "source_idx": source_idx,
                "ideal_company": original_ideal_name,
                "source_company": original_source_name,
                "score": score
            })
    
    return matches

# Function to generate a download link for a dataframe
def get_download_link(df, filename, link_text):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">{link_text}</a>'
    return href

# Sidebar for file uploads and settings
with st.sidebar:
    st.header("Matching Settings")
    
    # Thresholds
    st.subheader("Matching Thresholds")
    company_threshold = st.slider("Company Name Matching Threshold", 50, 100, DEFAULT_SETTINGS["thresholds"]["company_name"], 
                                help="Higher values require closer matches. 85 is recommended for most cases.")
    person_threshold = st.slider("Person Name Matching Threshold", 50, 100, DEFAULT_SETTINGS["thresholds"]["person_name"],
                               help="Higher values require closer matches for person names.")
    
    # Advanced settings expander
    with st.expander("Advanced Settings"):
        email_threshold = st.slider("Email Matching Threshold", 50, 100, DEFAULT_SETTINGS["thresholds"]["email"],
                                  help="Recommended to keep at 100 for exact email matching.")
        title_threshold = st.slider("Job Title Matching Threshold", 50, 100, DEFAULT_SETTINGS["thresholds"]["title"],
                                  help="Lower values allow matching similar job titles.")
        department_threshold = st.slider("Department Matching Threshold", 50, 100, DEFAULT_SETTINGS["thresholds"]["department"],
                                       help="Lower values allow matching similar department names.")

# Main content
if ideal_file is not None and len(source_files) > 0:
    # Read the ideal file
    ideal_df = try_read_csv(ideal_file)
    
    if ideal_df is not None:
        st.subheader("Ideal Contact List")
        st.write(f"Found {len(ideal_df)} contacts in the ideal list")
        
        # Display ideal file columns
        st.write("Columns in ideal file:", ideal_df.columns.tolist())
        
        # Process each source file
        all_matches = {}
        
        for source_file in source_files:
            source_df = try_read_csv(source_file)
            
            if source_df is not None:
                st.subheader(f"Source: {source_file.name}")
                st.write(f"Found {len(source_df)} contacts in {source_file.name}")
                
                # Display source file columns
                st.write("Columns in source file:", source_df.columns.tolist())
                
                # Column mapping
                st.write("Column Mapping")
                column_mapping = {}
                
                # Create two columns for mapping
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("Source Column")
                
                with col2:
                    st.write("Maps to Ideal Column")
                
                # For each important field, create a mapping
                for source_col, default_ideal_col in DEFAULT_SETTINGS["column_mapping"].items():
                    if source_col in source_df.columns:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write(source_col)
                        
                        with col2:
                            ideal_col = st.selectbox(
                                f"Map {source_col} to",
                                options=ideal_df.columns.tolist(),
                                index=ideal_df.columns.tolist().index(default_ideal_col) if default_ideal_col in ideal_df.columns else 0,
                                key=f"{source_file.name}_{source_col}"
                            )
                            column_mapping[source_col] = ideal_col
                
                # Match button
                if st.button(f"Match with {source_file.name}", key=f"match_{source_file.name}"):
                    # Show progress
                    with st.spinner(f"Matching companies in {source_file.name}..."):
                        # Perform matching
                        matches = match_companies(ideal_df, source_df, company_threshold, column_mapping)
                        all_matches[source_file.name] = matches
                    
                    # Display results
                    st.write(f"Found {len(matches)} matching companies")
                    
                    if len(matches) > 0:
                        # Create a dataframe of matches
                        matches_df = pd.DataFrame(matches)
                        
                        # Display matches
                        st.dataframe(matches_df)
                        
                        # Create a download link for the matches
                        st.markdown(
                            get_download_link(matches_df, f"matches_{source_file.name.split('.')[0]}.csv", "Download Matches CSV"),
                            unsafe_allow_html=True
                        )
                        
                        # Visualize match scores
                        fig = px.histogram(matches_df, x="score", nbins=10, title="Match Score Distribution")
                        st.plotly_chart(fig)
                        
                        # Display detailed matches
                        st.subheader("Detailed Matches")
                        
                        for match in matches:
                            with st.expander(f"{match['ideal_company']} ↔ {match['source_company']} (Score: {match['score']})"):
                                # Get the contacts from both dataframes
                                ideal_contact = ideal_df.iloc[match['ideal_idx']]
                                source_contact = source_df.iloc[match['source_idx']]
                                
                                # Display the contacts side by side
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.write("Ideal Contact")
                                    st.json(ideal_contact.to_dict())
                                
                                with col2:
                                    st.write("Source Contact")
                                    st.json(source_contact.to_dict())
else:
    # Display sample data section
    st.subheader("Sample Data")
    st.markdown("""
    To use this app, you'll need to upload:
    
    1. An **Ideal Contact List** - Your primary contact database (e.g., from your CRM)
    2. One or more **Source Contact Lists** - Other contact databases to compare against your ideal list
    
    ### Sample CSV Format
    
    Here's an example of what your CSV files might look like:
    
    **Ideal Contact List:**
    ```
    Company Name,First Name,Last Name,Email Address,Job Title,Department
    Acme Corporation,John,Doe,john.doe@acme.com,CTO,Technology
    Globex Inc,Jane,Smith,jane.smith@globex.com,CEO,Executive
    ```
    
    **Source Contact List:**
    ```
    Company,First Name,Last Name,Email Address,Position
    ACME Corp,John,Doe,john.doe@acme.com,Chief Technology Officer
    Globex,Jane,Smith,jane.smith@globex.com,Chief Executive Officer
    ```
    
    The app will help you map these different column names during the matching process.
    """)
    
    # Example usage
    st.subheader("How to Use")
    st.markdown("""
    1. Upload your ideal contact list CSV file
    2. Upload one or more source contact list CSV files
    3. Adjust the matching thresholds as needed
    4. Map the columns from your source files to the ideal file
    5. Click the "Match" button to find overlapping companies
    6. View and download the results
    """)
    
    # Features
    st.subheader("Features")
    st.markdown("""
    - **Fuzzy Matching**: Intelligently matches company names even with slight variations
    - **Customizable Thresholds**: Adjust matching sensitivity for different fields
    - **Interactive Results**: Explore matches with detailed information
    - **Data Visualization**: See match score distributions
    - **Export Options**: Download match results for further analysis
    """)
    
    # Add a footer with GitHub link
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center">
        <p>Created by Jim Sherlock | <a href="https://github.com/jrsherlock/contacts_matcher5000" target="_blank">GitHub Repository</a></p>
    </div>
    """, unsafe_allow_html=True) 