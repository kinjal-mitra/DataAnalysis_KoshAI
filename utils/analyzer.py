import pandas as pd
import numpy as np
from io import BytesIO


def pos_extractor(pcode):
    """
    Extract position code from PCode by taking last 2 characters.
    
    Args:
        pcode: Position code value from the PCode column
    
    Returns:
        int: Extracted position code number from last 2 characters
    """
    return int(pcode[-2:])


def analyze_excel_file(raw_df, station_id):
    """
    Analyze Excel data for a specific station ID and reorganize data by date and position.
    
    Args:
        raw_df: pandas DataFrame containing the raw data with columns:
                - Station_ID: Station identifier
                - PCode: Position code
                - Date_Time: Date and time information
                - Result: Result values to be organized
        station_id: The specific station ID to filter and analyze
    
    Returns:
        pandas DataFrame: Analyzed data organized by date and position codes
    """
    # Create dataframe on basis of Station_ID
    station_id_df = raw_df[raw_df["Station_ID"] == station_id].copy()
    
    # Apply pos_extractor to PCode column
    station_id_df["PositionCode"] = station_id_df["PCode"].apply(pos_extractor)
    station_id_df = station_id_df.drop(columns=["PCode"])
    
    # Get Unique dates
    unique_dates = station_id_df["Date_Time"].unique()
    
    # Create output Dataframe column names
    column_names = ["Station", "Dates"]
    for i in range(1, station_id_df["PositionCode"].max() + 1):
        name = "Data " + str(i)
        column_names.append(name)
    
    # Create Output Dataframe structure
    final_df = pd.DataFrame(columns=column_names)
    
    # Creating data-point for each date
    for date in unique_dates:
        req_data = station_id_df[station_id_df["Date_Time"] == date].drop(columns="Date_Time")
        req_data = req_data.sort_values(by="PositionCode", ascending=True)  # Sort DataFrame on Basis of PositionCode
        
        data_point = [station_id, date.strftime('%d-%m-%Y')]  # Create data point for each specific date
        data_point.extend([np.nan for i in range(1, station_id_df["PositionCode"].max() + 1)])  # Fill all Data values with np.nan
        
        for i in range(len(req_data)):
            for column in range(i, station_id_df["PositionCode"].max() + 1):
                if req_data.iloc[i, :]["PositionCode"] == column:
                    data_point[req_data.iloc[i, :]["PositionCode"] + 1] = req_data.iloc[i, :]["Result"]
        
        final_df.loc[len(final_df)] = data_point
    
    return final_df


def process_excel_for_web(input_file, station_id):
    """
    Wrapper function to process Excel file for web application.
    Reads the uploaded file, performs analysis, and returns result as downloadable Excel.
    
    Args:
        input_file: File object from Flask request
        station_id: Station ID to analyze (string: 'TUS' or 'CT')
    
    Returns:
        BytesIO: Excel file in memory ready for download
    
    Raises:
        ValueError: If required columns are missing or data is invalid
        Exception: For other processing errors
    """
    try:
        # Read the input Excel file into DataFrame
        raw_df = pd.read_excel(input_file)
        
        # Validate required columns
        required_columns = ["Station_ID", "PCode", "Date_Time", "Result"]
        missing_columns = [col for col in required_columns if col not in raw_df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Convert Date_Time to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(raw_df["Date_Time"]):
            raw_df["Date_Time"] = pd.to_datetime(raw_df["Date_Time"])
        
        # Validate Station ID (must be 'TUS' or 'CT')
        if station_id not in ['TUS', 'CT']:
            raise ValueError(f"Invalid Station ID: '{station_id}'. Must be either 'TUS' or 'CT'")
        
        # Check if station exists in data
        if station_id not in raw_df["Station_ID"].values:
            available_stations = raw_df["Station_ID"].unique().tolist()
            raise ValueError(f"Station ID '{station_id}' not found in the uploaded file. Available stations in file: {available_stations}")
        
        # Perform analysis - pass the DataFrame, not the file object
        result_df = analyze_excel_file(raw_df, station_id)
        
        # Create output Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Analysis')
        output.seek(0)
        
        return output
    
    except ValueError as ve:
        # Re-raise ValueError with original message
        raise ve
    except Exception as e:
        # Wrap other exceptions with more context
        raise Exception(f"Error processing Excel file: {str(e)}")


def get_available_stations(input_file):
    """
    Helper function to get list of available station IDs from uploaded file.
    Useful for displaying options to user.
    
    Args:
        input_file: File object from Flask request
    
    Returns:
        list: List of unique station IDs
    """
    try:
        raw_df = pd.read_excel(input_file)
        if "Station_ID" in raw_df.columns:
            return sorted(raw_df["Station_ID"].unique().tolist())
        else:
            return []
    except Exception:
        return []