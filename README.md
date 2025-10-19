# Excel Station Analyzer

Analyzes Excel files containing station data for TUS and CT stations.

## Required Excel Columns
- Station_ID (must contain 'TUS' or 'CT')
- PCode
- Date_Time
- Result

## Local Setup
1. Create virtual environment: `python -m venv venv`
2. Activate: `source venv/bin/activate` (Linux/Mac) or `venv\Scripts\activate` (Windows)
3. Install dependencies: `pip install -r requirements.txt`
4. Run: `python app.py`
5. Open browser: `http://localhost:5000`

## Deployment
See deployment instructions for Render or PythonAnywhere.