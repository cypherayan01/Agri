import pandas as pd
import numpy as np

def analyze_survey_data(excel_file_path):
    """
    Analyze survey data from Excel file's daily summary report sheet
    
    Args:
        excel_file_path (str): Path to the Excel file
    
    Returns:
        dict: Analysis results containing total surveys and district/village counts
    """
    
    try:
        
        df = pd.read_excel(excel_file_path, sheet_name='Daily_Survey Summary Report')
        
        print("Excel file loaded successfully!")
        print(f"Shape of data: {df.shape}")
        print(f"Columns in the sheet: {list(df.columns)}")
        print("\nFirst few rows:")
        print(df.head())
        
        results = {}
        
        
        survey_count_col = 'Total Survey Completed By Surveyor'
        district_col = 'District Name'
        village_col = 'Village Name'
        
        
        missing_cols = []
        if survey_count_col not in df.columns:
            missing_cols.append(survey_count_col)
        if district_col not in df.columns:
            missing_cols.append(district_col)
        if village_col not in df.columns:
            missing_cols.append(village_col)
        
        if missing_cols:
            print("\nAvailable columns:")
            for i, col in enumerate(df.columns):
                print(f"{i}: {col}")
            raise ValueError(f"Could not find required columns: {missing_cols}")
        
        print(f"\nUsing columns:")
        print(f"  Survey count: '{survey_count_col}'")
        print(f"  District: '{district_col}'")
        print(f"  Village: '{village_col}'")
        
        
        total_surveys = df[survey_count_col].sum()
        results['total_surveys_today'] = total_surveys
        
        
        villages_with_surveys = len(df[df[survey_count_col] > 0])
        results['villages_with_surveys'] = villages_with_surveys
        total_villages = len(df)
        results['total_villages'] = total_villages
        
        
        districts_with_surveys_df = df[df[survey_count_col] > 0]
        districts_with_surveys = districts_with_surveys_df[district_col].nunique()
        results['districts_with_surveys'] = districts_with_surveys
        total_districts = df[district_col].nunique()
        results['total_districts'] = total_districts
        
        return results
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None

def display_results(results):
    """Display the analysis results in a formatted way"""
    
    if results is None:
        print("No results to display due to error in processing.")
        return
    
    print("\n" + "="*50)
    print("SURVEY ANALYSIS RESULTS")
    print("="*50)
    
    print(f"\n Total surveys completed today: {results['total_surveys_today']:,}")
    
    if results['districts_with_surveys'] != "Not found":
        print(f"\n  Districts:")
        print(f"   • Districts with surveys: {results['districts_with_surveys']}")
        print(f"   • Total districts: {results['total_districts']}")
        if results['total_districts'] > 0:
            percentage = (results['districts_with_surveys'] / results['total_districts']) * 100
            print(f"   • Coverage: {percentage:.1f}%")
    
    if results['villages_with_surveys'] != "Not found":
        print(f"\n  Villages:")
        print(f"   • Villages with surveys: {results['villages_with_surveys']}")
        print(f"   • Total villages: {results['total_villages']}")
        if results['total_villages'] > 0:
            percentage = (results['villages_with_surveys'] / results['total_villages']) * 100
            print(f"   • Coverage: {percentage:.1f}%")

if __name__ == "__main__":
    excel_file = "DAILY_SURVEY_SUMMARY_REPORT_1210832_2026-01-13_00_42_51.816.xlsx"
    
    results = analyze_survey_data(excel_file)
    
    display_results(results)