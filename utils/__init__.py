"""
Utils package for Excel Analyzer application.

This package contains utility functions for analyzing Excel files,
specifically for processing station data and reorganizing it by date and position codes.
"""

from .analyzer import (
    analyze_excel_file,
    process_excel_for_web,
    get_available_stations,
    pos_extractor
)

__all__ = [
    'analyze_excel_file',
    'process_excel_for_web',
    'get_available_stations',
    'pos_extractor'
]

__version__ = '1.0.0'