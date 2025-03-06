# Contacts Matcher 5000 - Streamlit App

A powerful web application for matching company names and finding potential duplicates in your contact lists. This Streamlit app provides an intuitive interface for uploading contact data and identifying matching companies based on advanced string matching algorithms.

## Features

- Easy-to-use web interface for uploading contact data
- Support for CSV file uploads
- Advanced name matching algorithms
- Configurable matching settings
- Interactive results display
- Export functionality for matched results

## Live Demo

Visit the live app at: [Contacts Matcher 5000](https://contacts-matcher5000.streamlit.app)

## Usage

1. Visit the app URL
2. Upload your CSV file containing company names
3. Configure matching settings (optional)
4. Review and download the results

## Local Development

To run this app locally:

1. Clone the repository:
```bash
git clone https://github.com/jrsherlock/contacts_matcher5000_streamlit.git
cd contacts_matcher5000_streamlit
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv-py311
source venv-py311/bin/activate  # On Windows: .\venv-py311\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the app:
```bash
streamlit run app.py
```

## Requirements

- Python 3.11
- Streamlit
- Other dependencies as listed in requirements.txt

## File Structure

- `app.py`: Main Streamlit application file
- `requirements.txt`: Python dependencies
- `runtime.txt`: Python version specification for deployment
- `.streamlit/`: Streamlit configuration directory

## Contributing

Feel free to submit issues and enhancement requests!