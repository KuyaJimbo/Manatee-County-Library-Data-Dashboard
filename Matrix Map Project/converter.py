import csv
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

def parse_events_calendar_csv(csv_file_path, output_excel_path):
    """
    Convert EventsCalendar2024.csv to Power BI optimized Excel workbook
    """
    
    # Initialize workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Read the CSV file
    with open(csv_file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Split content into sections
    sections = content.split('\n\n')
    
    # Process each section
    for section in sections:
        if not section.strip():
            continue
            
        lines = section.strip().split('\n')
        if not lines:
            continue
            
        # Get section header
        header = lines[0].strip()
        
        # Skip if no data rows
        if len(lines) < 2:
            continue
            
        # Parse different section types
        if header == "Summary":
            create_summary_sheet(wb, lines[1:])
        elif header == "Monthly Distribution":
            create_monthly_distribution_sheet(wb, lines[1:])
        elif header == "Day of the Week Distribution":
            create_day_of_week_sheet(wb, lines[1:])
        elif header == "Hour of the Day Distribution":
            create_hour_distribution_sheet(wb, lines[1:])
        elif header == "Daily/Hourly Distribution":
            create_daily_hourly_sheet(wb, lines[1:])
        elif "Library Branch Distribution" in header:
            create_library_branch_sheet(wb, lines[1:])
        elif header == "Audience Distribution":
            create_audience_sheet(wb, lines[1:])
        elif header == "Category Distribution":
            create_category_sheet(wb, lines[1:])
        elif "Individual Calendars Monthly Breakdown" in header:
            process_individual_calendars(wb, section)
    
    # Save the workbook
    wb.save(output_excel_path)
    print(f"Excel file saved to: {output_excel_path}")

def create_summary_sheet(wb, data_lines):
    """Create Summary worksheet"""
    ws = wb.create_sheet("Summary")
    
    # Headers
    ws.append(["Metric", "Registration_Type", "Total", "In_Person", "Online"])
    
    for line in data_lines:
        if line.strip():
            parts = [part.strip() for part in line.split(',')]
            if len(parts) >= 4:
                ws.append(parts)

def create_monthly_distribution_sheet(wb, data_lines):
    """Create Monthly Distribution worksheet in normalized format"""
    ws = wb.create_sheet("Monthly_Distribution")
    
    # Headers for normalized table
    ws.append(["Month_Year", "Metric", "Value"])
    
    # Parse header row to get months
    header_line = data_lines[0]
    months = [month.strip() for month in header_line.split(',')[1:]]
    
    # Process each metric row
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip() for part in line.split(',')]
            metric = parts[0]
            values = parts[1:]
            
            # Create normalized rows
            for i, value in enumerate(values):
                if i < len(months) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([months[i], metric, numeric_value])
                    except:
                        ws.append([months[i], metric, value])

def create_day_of_week_sheet(wb, data_lines):
    """Create Day of Week Distribution worksheet"""
    ws = wb.create_sheet("Day_of_Week_Distribution")
    
    ws.append(["Day_of_Week", "Metric", "Value"])
    
    # Get days from header
    header_line = data_lines[0]
    days = [day.strip() for day in header_line.split(',')[1:]]
    
    # Process each metric
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip() for part in line.split(',')]
            metric = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(days) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([days[i], metric, numeric_value])
                    except:
                        ws.append([days[i], metric, value])

def create_hour_distribution_sheet(wb, data_lines):
    """Create Hour Distribution worksheet"""
    ws = wb.create_sheet("Hour_Distribution")
    
    ws.append(["Hour", "Metric", "Value"])
    
    # Get hours from header
    header_line = data_lines[0]
    hours = [hour.strip() for hour in header_line.split(',')[1:]]
    
    # Process each metric
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip() for part in line.split(',')]
            metric = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(hours) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([hours[i], metric, numeric_value])
                    except:
                        ws.append([hours[i], metric, value])

def create_daily_hourly_sheet(wb, data_lines):
    """Create Daily/Hourly Distribution worksheet"""
    ws = wb.create_sheet("Daily_Hourly_Distribution")
    
    ws.append(["Day_of_Week", "Hour", "Value"])
    
    # Get hours from header
    header_line = data_lines[0]
    hours = [hour.strip() for hour in header_line.split(',')[1:]]
    
    # Process each day
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip() for part in line.split(',')]
            day = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(hours) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        if numeric_value > 0:  # Only include non-zero values
                            ws.append([day, hours[i], numeric_value])
                    except:
                        continue

def create_library_branch_sheet(wb, data_lines):
    """Create Library Branch Distribution worksheet"""
    ws = wb.create_sheet("Library_Branch_Distribution")
    
    ws.append(["Month_Year", "Library_Branch", "Events"])
    
    # Get months from header
    header_line = data_lines[0]
    months = [month.strip().strip('"') for month in header_line.split(',')[1:]]
    
    # Process each branch
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip().strip('"') for part in line.split(',')]
            branch = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(months) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([months[i], branch, numeric_value])
                    except:
                        ws.append([months[i], branch, value])

def create_audience_sheet(wb, data_lines):
    """Create Audience Distribution worksheet"""
    ws = wb.create_sheet("Audience_Distribution")
    
    ws.append(["Month_Year", "Audience_Type", "Events"])
    
    # Get months from header
    header_line = data_lines[0]
    months = [month.strip().strip('"') for month in header_line.split(',')[1:]]
    
    # Process each audience type
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip().strip('"') for part in line.split(',')]
            audience = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(months) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([months[i], audience, numeric_value])
                    except:
                        ws.append([months[i], audience, value])

def create_category_sheet(wb, data_lines):
    """Create Category Distribution worksheet"""
    ws = wb.create_sheet("Category_Distribution")
    
    ws.append(["Month_Year", "Category", "Events"])
    
    # Get months from header
    header_line = data_lines[0]
    months = [month.strip().strip('"') for month in header_line.split(',')[1:]]
    
    # Process each category
    for line in data_lines[1:]:
        if line.strip():
            parts = [part.strip().strip('"') for part in line.split(',')]
            category = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(months) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        ws.append([months[i], category, numeric_value])
                    except:
                        ws.append([months[i], category, value])

def process_individual_calendars(wb, section_content):
    """Process Individual Calendars section with multiple sub-tables"""
    lines = section_content.strip().split('\n')
    
    current_table = None
    header_found = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check if this is a table header
        if 'Monthly Distribution' in line and line.startswith('"') and line.endswith('"'):
            table_name = line.strip('"').replace(' ', '_').replace('/', '_')
            # Clean up table name
            table_name = re.sub(r'[^\w\s-]', '', table_name)
            table_name = re.sub(r'\s+', '_', table_name)
            
            # Create new worksheet for this sub-table
            current_table = wb.create_sheet(f"IndCal_{table_name}")
            current_table.append(["Month_Year", "Metric", "Value"])
            header_found = False
            continue
        
        # Check if this is the month header row
        if line.startswith('Month/Year') and current_table is not None:
            months = [month.strip() for month in line.split(',')[1:]]
            header_found = True
            continue
        
        # Process data rows
        if header_found and current_table is not None and ',' in line:
            parts = [part.strip() for part in line.split(',')]
            metric = parts[0]
            values = parts[1:]
            
            for i, value in enumerate(values):
                if i < len(months) and value:
                    try:
                        numeric_value = int(value) if value.isdigit() else 0
                        current_table.append([months[i], metric, numeric_value])
                    except:
                        current_table.append([months[i], metric, value])

def main():
    """Main function to run the conversion"""
    # Define file paths
    input_csv = "EventsCalendar2024.csv"
    output_excel = "EventsCalendar2024_PowerBI.xlsx"
    
    # Check if input file exists
    if not os.path.exists(input_csv):
        print(f"Error: Input file '{input_csv}' not found.")
        print("Please ensure the CSV file is in the same directory as this script.")
        return
    
    try:
        # Convert the CSV to Excel
        parse_events_calendar_csv(input_csv, output_excel)
        
        # Print summary
        if os.path.exists(output_excel):
            wb = load_workbook(output_excel)
            print(f"\nConversion completed successfully!")
            print(f"Created {len(wb.sheetnames)} worksheets:")
            for sheet_name in wb.sheetnames:
                print(f"  - {sheet_name}")
            wb.close()
            
            print(f"\nThe file '{output_excel}' is now ready for Power BI import.")
            print("Each worksheet represents a different data dimension for analysis.")
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        print("Please check your CSV file format and try again.")

if __name__ == "__main__":
    main()