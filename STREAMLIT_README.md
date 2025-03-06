# Contacts Matcher 5000 - Streamlit App

A web-based interface for the Contacts Matcher 5000 tool, built with Streamlit. This application allows you to analyze and compare contact databases across different sources, identifying overlapping company relationships and their associated contacts.

## Features

- **User-friendly Interface**: Upload CSV files and configure matching settings through an intuitive web interface
- **Fuzzy Matching**: Intelligently matches company names even with slight variations
- **Customizable Thresholds**: Adjust matching sensitivity for different fields
- **Interactive Results**: Explore matches with detailed information
- **Data Visualization**: See match score distributions
- **Export Options**: Download match results for further analysis
- **Light Theme**: Clean, professional interface optimized for readability

## Getting Started

### Local Development

1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd contacts_matcher5000
   ```

2. Create and activate a Python 3.11 virtual environment:
   ```bash
   python -m venv venv-py311
   source venv-py311/bin/activate  # On Windows: venv-py311\Scripts\activate
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

5. Open your browser and navigate to `http://localhost:8501`

### Deployment on Streamlit Cloud

1. Push your code to GitHub
2. Visit [Streamlit Cloud](https://streamlit.io/cloud)
3. Click "New app"
4. Select your repository and branch
5. Set the main file path to `app.py`
6. Click "Deploy!"

## Usage

1. **Upload Your Contact Lists**:
   - Upload your ideal/master contact list (e.g., from your CRM)
   - Upload one or more source contact lists to compare against your ideal list

2. **Configure Settings**:
   - Adjust the matching thresholds as needed
   - Advanced settings available for fine-tuning

3. **View Results**:
   - See matching companies and their confidence scores
   - Download results for further analysis

## Input File Format

The app expects CSV files with contact information. Required columns:
- Company Name
- First Name
- Last Name
- Email Address
- Job Title/Position

Optional columns:
- Department
- LinkedIn URL

Don't worry about exact column names - you can map them in the interface.

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin feature-name`
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details. 