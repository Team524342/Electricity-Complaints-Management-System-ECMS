# excel_handler.py - Excel integration functionality (continued)
import pandas as pd
import os
from datetime import datetime
import io
from flask import send_file

def generate_report_excel(complaints_df):
    """
    Generate Excel report from complaints data
    Returns a BytesIO object containing the Excel file
    """
    # Create a BytesIO object to store the Excel file
    output = io.BytesIO()
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write the complaints DataFrame to a sheet
        complaints_df.to_excel(writer, sheet_name='All Complaints', index=False)
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['All Complaints']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D9D9D9',
            'border': 1
        })
        
        # Apply the header format to the first row
        for col_num, value in enumerate(complaints_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Adjust column widths
        for i, col in enumerate(complaints_df.columns):
            column_width = max(len(str(col)), 
                              complaints_df[col].astype(str).map(len).max())
            worksheet.set_column(i, i, column_width + 2)
        
        # Create a status summary sheet
        status_counts = complaints_df['status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        status_counts.to_excel(writer, sheet_name='Status Summary', index=False)
        
        # Create a category summary sheet
        category_counts = complaints_df['category'].value_counts().reset_index()
        category_counts.columns = ['Category', 'Count']
        category_counts.to_excel(writer, sheet_name='Category Summary', index=False)
        
        # Create charts in the summary sheets
        # Status chart
        status_chart = workbook.add_chart({'type': 'pie'})
        status_chart.add_series({
            'name': 'Complaint Status',
            'categories': ['Status Summary', 1, 0, len(status_counts), 0],
            'values': ['Status Summary', 1, 1, len(status_counts), 1],
            'data_labels': {'percentage': True}
        })
        status_chart.set_title({'name': 'Complaints by Status'})
        status_chart.set_style(10)
        worksheet_status = writer.sheets['Status Summary']
        worksheet_status.insert_chart('D2', status_chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Category chart
        category_chart = workbook.add_chart({'type': 'column'})
        category_chart.add_series({
            'name': 'Complaint Categories',
            'categories': ['Category Summary', 1, 0, len(category_counts), 0],
            'values': ['Category Summary', 1, 1, len(category_counts), 1],
            'data_labels': {'value': True}
        })
        category_chart.set_title({'name': 'Complaints by Category'})
        category_chart.set_x_axis({'name': 'Category'})
        category_chart.set_y_axis({'name': 'Number of Complaints'})
        category_chart.set_style(11)
        worksheet_category = writer.sheets['Category Summary']
        worksheet_category.insert_chart('D2', category_chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Monthly trends sheet (if we have enough data)
        if not complaints_df.empty:
            # Convert submission_date to datetime
            complaints_df['submission_date'] = pd.to_datetime(complaints_df['submission_date'])
            
            # Group by month and count
            monthly_data = complaints_df.groupby(pd.Grouper(key='submission_date', freq='M')).size().reset_index()
            monthly_data.columns = ['Month', 'Count']
            monthly_data['Month'] = monthly_data['Month'].dt.strftime('%Y-%m')
            
            # Write to sheet
            monthly_data.to_excel(writer, sheet_name='Monthly Trends', index=False)
            
            # Create a line chart
            trend_chart = workbook.add_chart({'type': 'line'})
            trend_chart.add_series({
                'name': 'Monthly Complaints',
                'categories': ['Monthly Trends', 1, 0, len(monthly_data), 0],
                'values': ['Monthly Trends', 1, 1, len(monthly_data), 1],
                'marker': {'type': 'circle', 'size': 8},
                'data_labels': {'value': True}
            })
            trend_chart.set_title({'name': 'Monthly Complaint Trends'})
            trend_chart.set_x_axis({'name': 'Month'})
            trend_chart.set_y_axis({'name': 'Number of Complaints'})
            trend_chart.set_style(12)
            worksheet_trends = writer.sheets['Monthly Trends']
            worksheet_trends.insert_chart('D2', trend_chart, {'x_scale': 1.5, 'y_scale': 1.5})
    
    # Reset the pointer to the beginning of the BytesIO object
    output.seek(0)
    
    return output

def export_complaints_excel(complaints_df):
    """
    Export complaints data to Excel file and return it as a downloadable response
    """
    output = generate_report_excel(complaints_df)
    
    # Create timestamp for filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"complaints_report_{timestamp}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

def import_complaints_from_excel(file_path):
    """
    Import complaints from Excel file
    Returns a DataFrame with the imported data
    """
    try:
        df = pd.read_excel(file_path)
        required_columns = [
            'complaint_id', 'user_id', 'category', 'description', 
            'location', 'submission_date', 'status'
        ]
        
        # Check if required columns exist
        for col in required_columns:
            if col not in df.columns:
                return None, f"Required column '{col}' not found in Excel file"
        
        return df, "Import successful"
    except Exception as e:
        return None, f"Error importing Excel file: {str(e)}"

def backup_database():
    """
    Create a backup of the database files
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_dir = 'backups'
    os.makedirs(backup_dir, exist_ok=True)
    
    # Backup complaints file
    if os.path.exists('data/complaints.xlsx'):
        complaints_df = pd.read_excel('data/complaints.xlsx')
        complaints_df.to_excel(f"{backup_dir}/complaints_backup_{timestamp}.xlsx", index=False)
    
    # Backup users file
    if os.path.exists('data/users.xlsx'):
        users_df = pd.read_excel('data/users.xlsx')
        users_df.to_excel(f"{backup_dir}/users_backup_{timestamp}.xlsx", index=False)
    
    return f"Backup created at {timestamp}"