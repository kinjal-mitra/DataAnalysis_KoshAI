import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend for server environments


def analyze_excel_file(raw_df, station_id):
    """
    Analyze Excel data for a specific station ID and reorganize data by date and position.
    
    Args:
        raw_df: pandas DataFrame containing the raw data with columns:
                - Station_ID: Station identifier
                - PCode: Data Name/Code
                - Date_Time: Date and time information
                - Result: Result values to be organized
        station_id: The specific station ID to filter and analyze
    
    Returns:
        pandas DataFrame: Analyzed data organized by date and position codes
    """
    # Create dataframe on basis of Station_ID
    station_id_df = raw_df[raw_df["Station_ID"] == station_id].copy()
    
    # Get Unique dates
    unique_dates = station_id_df["Date_Time"].unique()
    
    # Create output Dataframe column names
    column_names = ["Station", "Dates"]
    unique_PCode = station_id_df["PCode"].unique().tolist()
    unique_PCode.sort()
    column_names.extend(unique_PCode)

    # Create Output Dataframe structure
    final_df = pd.DataFrame(columns=column_names)
    
    # Creating data-point for each date
    for date in unique_dates:
        req_data = station_id_df[station_id_df["Date_Time"]==date].drop(columns="Date_Time")
        req_data = req_data.sort_values(by="PCode", ascending=True)  # Sort DataFrame on Basis of PositionCode
        
        data_point = [station_id, date.strftime('%d-%m-%Y')]  # Create data point for each specific date
        data_point.extend([np.nan for i in range (1, len(station_id_df["PCode"].unique())+1)])  # Fill all Data values with np.nan
        
        for i in range (0, len(req_data)):
            for column in unique_PCode:
                if req_data.iloc[i, :]["PCode"]==column:
                    data_point[unique_PCode.index(column)+2]=req_data.iloc[i, :]["Result"]
        final_df.loc[len(final_df)] = data_point
    
    return final_df


def get_available_pcodes(input_file, station_id):
    """
    Get list of unique PCode values for a specific station from uploaded file.
    
    Args:
        input_file: File object from Flask request
        station_id: Station ID to filter ('TUS' or 'CT')
    
    Returns:
        list: Sorted list of unique PCode values
    """
    try:
        raw_df = pd.read_excel(input_file)
        
        if "PCode" not in raw_df.columns or "Station_ID" not in raw_df.columns:
            return []
        
        # Filter by station_id and get unique PCodes
        station_df = raw_df[raw_df["Station_ID"] == station_id]
        unique_pcodes = sorted(station_df["PCode"].unique().tolist())
        
        return unique_pcodes
    except Exception:
        return []


def create_graph(final_df, data_number):
    """
    Create a graph/chart from the analyzed dataframe.
    Plots specified data column values (y-axis) vs Dates (x-axis).
    
    Args:
        final_df: pandas DataFrame with analyzed data
                 Columns: Station, Dates, PCode1, PCode2, ...
        data_number: String indicating which data column to plot (PCode value)
    
    Returns:
        BytesIO: Image buffer containing the graph as PNG
    """
    
    # Create figure and axis
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Create numeric x-axis positions
    x_positions = range(len(final_df))
    
    # Plot specified data column vs numeric positions
    ax.plot(x_positions, final_df[data_number], 
            marker='o', 
            linewidth=2, 
            markersize=8,
            color='#667eea',
            label=data_number)
    
    # Customize the plot
    ax.set_xlabel('Dates', fontsize=12, fontweight='normal')
    ax.set_ylabel(f'{data_number}', fontsize=12, fontweight='normal')
    ax.set_title(f'{final_df["Station"].iloc[0]} - {data_number} Analysis', 
                 fontsize=16, fontweight='bold')
    ax.legend(loc='best', fontsize=10)
    ax.grid(True, alpha=0.3, linestyle='--')

    # Create evenly spaced ticks including start and end
    num_ticks = min(10, len(final_df))  # Don't exceed number of data points
    dates = final_df['Dates'].reset_index(drop=True)
    tick_positions = np.linspace(0, len(final_df) - 1, num_ticks, dtype=int)
    
    # Set the x-ticks at the calculated positions with corresponding date labels
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(dates.iloc[tick_positions], rotation=45, ha='right')
    
    # Adjust layout to prevent label cutoff
    plt.tight_layout()
    
    # Save to buffer
    img_buffer = BytesIO()
    fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
    img_buffer.seek(0)
    
    # Close the figure to free memory
    plt.close(fig)
    
    return img_buffer


def process_excel_for_web(input_file, station_id, pcode1=None, pcode2=None):
    """
    Wrapper function to process Excel file for web application.
    Reads the uploaded file, performs analysis, and returns result as downloadable Excel.
    
    Args:
        input_file: File object from Flask request
        station_id: Station ID to analyze (string: 'TUS' or 'CT')
        pcode1: First PCode to generate chart (optional)
        pcode2: Second PCode to generate chart (optional)
    
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

        # Use XlsxWriter engine to support image insertion
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # Write the dataframe to first sheet
            result_df.to_excel(writer, index=False, sheet_name=f'{station_id} station')
            
            # Get workbook object
            workbook = writer.book
            worksheet = writer.sheets[f'{station_id} station']

            # === Add font formatting ===
            header_format = workbook.add_format({
                'font_name': 'Arial', 
                'font_size': 12, 
                'bold': True, 
                'align': 'center', 
                'valign': 'vcenter'
            })

            cell_format = workbook.add_format({
                'font_name': 'Times New Roman',
                'font_size': 12,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Apply header format
            for col_num, value in enumerate(result_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Apply data cell format
            worksheet.set_column(0, len(result_df.columns) - 1, 15, cell_format)
            
            # === Add graphs if PCodes are provided ===
            if pcode1 and pcode1 in result_df.columns:
                img_buffer = create_graph(result_df, pcode1)
                graph_worksheet = workbook.add_worksheet(f'{pcode1} Chart')
                graph_worksheet.insert_image('B2', f'{pcode1}_Chart.png', {'image_data': img_buffer})
            
            if pcode2 and pcode2 in result_df.columns:
                img_buffer = create_graph(result_df, pcode2)
                graph_worksheet = workbook.add_worksheet(f'{pcode2} Chart')
                graph_worksheet.insert_image('B2', f'{pcode2}_Chart.png', {'image_data': img_buffer})

        output.seek(0)
        
        return output
    
    except ValueError as ve:
        raise ve
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")


def get_available_stations(input_file):
    """
    Helper function to get list of available station IDs from uploaded file.
    
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