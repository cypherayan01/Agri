import pandas as pd
import json
import sys
from pathlib import Path

def read_excel_sheet1(file_path):
    """Read Sheet1 from the Excel file"""
    try:
        # Read with first row as header
        df = pd.read_excel(file_path, sheet_name='Sheet 1', header=0)
        
        # If columns are still unnamed, try using the first row as column names
        if 'Unnamed' in str(df.columns[1]):
            # The actual headers are in the first row
            new_columns = df.iloc[0].tolist()
            df.columns = new_columns
            df = df.iloc[1:].reset_index(drop=True)
        
        return df
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_sheet2_data(sheet1_df):
    """Generate Sheet2 data based on Sheet1 data with calculated columns"""
    
    if sheet1_df is None:
        return None
    
    sheet2_data = []
    current_date = pd.Timestamp.now()
    
    for index, row in sheet1_df.iterrows():
        # Skip if state is NaN or empty
        if pd.isna(row['State']) or str(row['State']).strip() == '':
            continue
            
        # Map columns from Sheet1 to Sheet2 structure
        # a = State
        # b = No. of Plots Targeted for Rabi DCS  
        # c = Total Plots Surveyed (assume 0 for now, you can update this)
        # d = Rabi DCS Activity Start Date
        # e = Rabi DCS Activity End Date
        
        state = row['State']
        plots_targeted = row['No. of Plots Targeted for Rabi DCS']
        start_date = row['Rabi DCS Activity Start Date']
        end_date = row['Rabi DCS Activity End Date']
        
        # Convert plots_targeted to numeric
        try:
            plots_targeted = pd.to_numeric(str(plots_targeted).replace(',', ''), errors='coerce')
            if pd.isna(plots_targeted):
                plots_targeted = 0
        except:
            plots_targeted = 0
        
        # Parse dates with proper format
        try:
            start_date_parsed = pd.to_datetime(start_date, dayfirst=True) if not pd.isna(start_date) else None
            end_date_parsed = pd.to_datetime(end_date, dayfirst=True) if not pd.isna(end_date) else None
        except:
            start_date_parsed = None
            end_date_parsed = None
        
        # Calculate derived fields
        total_plots_surveyed = 0  # You can update this with actual survey data
        
        # Calculate days
        if start_date_parsed and end_date_parsed:
            total_days = (end_date_parsed - start_date_parsed).days
            days_elapsed = (current_date - start_date_parsed).days
            days_elapsed = max(0, min(days_elapsed, total_days))  # Clamp between 0 and total_days
        else:
            total_days = 0
            days_elapsed = 0
        
        # Calculate required per day
        required_per_day = plots_targeted / total_days if total_days > 0 else 0
        
        # Calculate actual per day
        actual_per_day = total_plots_surveyed / days_elapsed if days_elapsed > 0 else 0
        
        # Calculate required percentage (time elapsed)
        required_percentage = (days_elapsed / total_days * 100) if total_days > 0 else 0
        
        # Calculate actual percentage (survey done)
        actual_percentage = (total_plots_surveyed / plots_targeted * 100) if plots_targeted > 0 else 0
        
        # Calculate gap
        gap = actual_percentage - required_percentage
        
        # Format gap with emoji
        if gap >= 0:
            gap_formatted = f"✅ +{gap:.1f}%"
        else:
            gap_formatted = f"🔻{gap:.1f}%"
        
        # Create Sheet2 row
        sheet2_row = {
            'a': state,  # State
            'b': plots_targeted,  # No. of Plots Targeted for Rabi DCS
            'c': total_plots_surveyed,  # Total Plots Surveyed
            'd': start_date,  # Rabi DCS Activity Start Date
            'e': end_date,  # Rabi DCS Activity End Date
            'days': days_elapsed,
            'Total Days': total_days,
            'Days Elapsed': days_elapsed,
            'Required /day': round(required_per_day, 1),
            'Actual /day': round(actual_per_day, 1),
            'Required % (Time elapsed)': round(required_percentage, 1),
            'Actual % (Survey done)': round(actual_percentage, 1),
            'Gap': gap_formatted
        }
        
        sheet2_data.append(sheet2_row)
    
    return sheet2_data

def convert_to_json(data):
    """Convert data to JSON format with proper handling of pandas/numpy types"""
    if data is None:
        return None
    
    # Convert pandas/numpy types to Python native types for JSON serialization
    def convert_types(obj):
        if pd.isna(obj):
            return None
        elif isinstance(obj, pd.Timestamp):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        elif isinstance(obj, (pd.Int64Dtype, int)):
            return int(obj) if not pd.isna(obj) else None
        elif isinstance(obj, float):
            return float(obj) if not pd.isna(obj) else None
        else:
            return str(obj)
    
    # Apply conversion to all data
    converted_data = []
    for row in data:
        converted_row = {k: convert_types(v) for k, v in row.items()}
        converted_data.append(converted_row)
    
    return json.dumps(converted_data, indent=2, ensure_ascii=False)

def main():
    # File path for the Excel file
    excel_file = "Rabi_Plan_Sheet1.xlsx"
    
    # Check if file exists
    if not Path(excel_file).exists():
        print(f"Error: '{excel_file}' not found in current directory.")
        print("Please ensure the file is in the same directory as this script.")
        return
    
    print(f"Reading Sheet1 from {excel_file}...")
    
    # Read Sheet1 data
    sheet1_df = read_excel_sheet1(excel_file)
    
    if sheet1_df is None:
        return
    
    print(f"Sheet1 loaded successfully with {len(sheet1_df)} rows and {len(sheet1_df.columns)} columns.")
    
    # Generate Sheet2 data with calculations
    print("Generating Sheet2 data with calculations...")
    sheet2_data = generate_sheet2_data(sheet1_df)
    
    if sheet2_data is None:
        print("Failed to generate Sheet2 data.")
        return
    
    print(f"Generated {len(sheet2_data)} rows for Sheet2.")
    
    # Convert to JSON
    print("Converting to JSON format...")
    json_output = convert_to_json(sheet2_data)
    
    if json_output is None:
        print("Failed to convert to JSON.")
        return
    
    # Save JSON to file
    output_file = "sheet2_data.json"
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_output)
        print(f"Sheet2 data saved to {output_file}")
    except Exception as e:
        print(f"Error saving JSON file: {e}")
    
    # Display column mapping
    print("\nColumn Mapping:")
    print("a = State")
    print("b = No. of Plots Targeted for Rabi DCS")
    print("c = Total Plots Surveyed (currently set to 0)")
    print("d = Rabi DCS Activity Start Date") 
    print("e = Rabi DCS Activity End Date")
    
    # Also print to console (first few records)
    print(f"\nFirst 3 records of Sheet2 data in JSON format:")
    try:
        data_list = json.loads(json_output)
        preview = data_list[:3] if len(data_list) > 3 else data_list
        print(json.dumps(preview, indent=2, ensure_ascii=False))
    except Exception as e:
        print(f"Error displaying preview: {e}")

if __name__ == "__main__":
    main()