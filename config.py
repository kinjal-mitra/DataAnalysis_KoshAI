import os
from dotenv import load_dotenv
load_dotenv()

class Config:
    """Flask application configuration"""
    
    # Secret key for session management
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    
    # File upload settings
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    UPLOAD_FOLDER = 'uploads'
    ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
    
    # Station IDs - only these two values are valid
    ALLOWED_STATIONS = ['TUS', 'CT']
    
    # Optional: Station names for display purposes
    STATION_NAMES = {
        'TUS': 'TUS Station',
        'CT': 'CT Station'
    }