"""
Debug script to test the analyzer function with a sample Excel file.
Run this to verify your functions work correctly before testing in Flask.
"""

import pandas as pd
from io import BytesIO
from utils.analyzer import process_excel_for_web, analyze_excel_file

# Create a sample Excel file for testing
def create_sample_excel():
    """Create a sample Excel file with test data"""
    data = {
        'Station_ID': ['TUS', 'TUS', 'TUS', 'CT', 'CT', 'CT'],
        'PCode': ['P01', 'P02', 'P03', 'P01', 'P02', 'P03'],
        'Date_Time': pd.to_datetime(['2024-01-01', '2024-01-01', '2024-01-01', 
                                      '2024-01-01', '2024-01-01', '2024-01-01']),
        'Result': [10.5, 20.3, 15.7, 12.1, 18.9, 22.4]
    }
    
    df = pd.DataFrame(data)
    
    # Save to BytesIO (simulating uploaded file)
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_buffer.seek(0)
    
    return excel_buffer

# Test the function
if __name__ == "__main__":
    print("Creating sample Excel file...")
    test_file = create_sample_excel()
    
    print("Testing with TUS station...")
    try:
        result = process_excel_for_web(test_file, 'TUS')
        print("✓ TUS analysis completed successfully!")
        
        # Read the result to verify
        result.seek(0)
        result_df = pd.read_excel(result)
        print("\nResult preview:")
        print(result_df)
        
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()
    
    # Test with CT
    print("\n" + "="*50)
    print("Testing with CT station...")
    test_file.seek(0)  # Reset file pointer
    try:
        result = process_excel_for_web(test_file, 'CT')
        print("✓ CT analysis completed successfully!")
        
        result.seek(0)
        result_df = pd.read_excel(result)
        print("\nResult preview:")
        print(result_df)
        
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()