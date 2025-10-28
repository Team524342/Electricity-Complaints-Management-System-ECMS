
#     app.register_blueprint(excel_bp, url_prefix=url_prefix)
from flask import Blueprint, render_template, request, jsonify, current_app
import pandas as pd
import os
import json

# Create a blueprint factory function instead of a direct blueprint
def create_excel_editor_blueprint(name, excel_file, sheet_name):
    # Create a new blueprint instance with a unique name
    excel_bp = Blueprint(f'excel_editor_{name}', __name__, template_folder='templates')
    
    # Helper function to read Excel file
    def read_excel():
        # Create a sample file if it doesn't exist
        if not os.path.exists(excel_file):
            create_sample_excel(excel_file, sheet_name)
        return pd.read_excel(excel_file, sheet_name=sheet_name)

    # Helper function to write to Excel file
    def write_excel(df):
        df.to_excel(excel_file, index=False, sheet_name=sheet_name)
    
    # Create a sample Excel file if it doesn't exist
    def create_sample_excel(file_path, sheet):
        if not os.path.exists(file_path):
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
            
            # Create sample data based on the file name
            if 'employee' in file_path.lower():
                data = {
                    'id': [1, 2, 3],
                    'name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
                    'position': ['Manager', 'Developer', 'Designer'],
                    'department': ['HR', 'IT', 'Marketing'],
                    'salary': [75000, 85000, 65000]
                }
            elif 'product' in file_path.lower():
                data = {
                    'id': [1, 2, 3],
                    'name': ['Laptop', 'Smartphone', 'Tablet'],
                    'price': [1200, 800, 500],
                    'stock': [45, 120, 75],
                    'category': ['Electronics', 'Electronics', 'Electronics']
                }
            elif 'customer' in file_path.lower():
                data = {
                    'id': [1, 2, 3],
                    'name': ['Acme Corp', 'Globex Inc', 'Initech LLC'],
                    'contact': ['Jane Smith', 'John Brown', 'Mike Wilson'],
                    'email': ['jsmith@acme.com', 'jbrown@globex.com', 'mwilson@initech.com'],
                    'phone': ['555-1234', '555-5678', '555-9012']
                }
            else:
                data = {
                    'id': [1, 2, 3],
                    'column1': ['Value 1', 'Value 2', 'Value 3'],
                    'column2': ['Data 1', 'Data 2', 'Data 3']
                }
            
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, sheet_name=sheet)

    @excel_bp.route('/')
    def index():
        # Ensure we have an Excel file
        if not os.path.exists(excel_file):
            create_sample_excel(excel_file, sheet_name)
            
        # Extract just the filename without path for display
        filename = os.path.basename(excel_file)
        
        return render_template(
            'excel_editor.html', 
            excel_file=filename,
            sheet_name=sheet_name
        )

    @excel_bp.route('/data')
    def get_data():
        df = read_excel()
        # Convert to list of dictionaries for JSON response
        records = df.to_dict('records')
        return jsonify({"data": records})

    @excel_bp.route('/add', methods=['POST'])
    def add_record():
        try:
            df = read_excel()
            
            # Get new record data from form
            new_record = {}
            for column in df.columns:
                new_record[column] = request.form.get(column)
            
            # Determine new ID if 'id' is a column
            if 'id' in df.columns:
                if df['id'].dtype == 'int64':
                    new_record['id'] = int(df['id'].max() + 1) if not df.empty else 1
            
            # Append new record to dataframe
            df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
            write_excel(df)
            
            return jsonify({"success": True, "message": "Record added successfully"})
        except Exception as e:
            return jsonify({"success": False, "message": str(e)})

    @excel_bp.route('/update', methods=['POST'])
    def update_record():
        try:
            df = read_excel()
            data = json.loads(request.data)
            
            # Extract record ID and updated data
            record_id = data.get('id')
            updated_data = data.get('data')
            
            # Find the row with matching ID and update it
            if 'id' in df.columns:
                idx = df.index[df['id'] == int(record_id)].tolist()
                if idx:
                    for key, value in updated_data.items():
                        df.at[idx[0], key] = value
                    write_excel(df)
                    return jsonify({"success": True, "message": "Record updated successfully"})
                else:
                    return jsonify({"success": False, "message": "Record not found"})
            else:
                return jsonify({"success": False, "message": "ID column not found in Excel file"})
        except Exception as e:
            return jsonify({"success": False, "message": str(e)})

    @excel_bp.route('/delete', methods=['POST'])
    def delete_record():
        try:
            df = read_excel()
            data = json.loads(request.data)
            
            # Get record ID to delete
            record_id = data.get('id')
            
            # Delete the row with matching ID
            if 'id' in df.columns:
                df = df[df['id'] != int(record_id)]
                write_excel(df)
                return jsonify({"success": True, "message": "Record deleted successfully"})
            else:
                return jsonify({"success": False, "message": "ID column not found in Excel file"})
        except Exception as e:
            return jsonify({"success": False, "message": str(e)})

    @excel_bp.route('/columns')
    def get_columns():
        df = read_excel()
        columns = list(df.columns)
        return jsonify({"columns": columns})

    return excel_bp

# Helper function to register excel editors more easily
def register_excel_editors(app, editors_config):
    """
    Register multiple Excel editors with the Flask app
    
    Args:
        app: Flask application
        editors_config: List of dictionaries with configuration for each editor
                        Each dict should have: name, url_prefix, excel_file, sheet_name
    """
    for config in editors_config:
        blueprint = create_excel_editor_blueprint(
            name=config['name'],
            excel_file=config['excel_file'],
            sheet_name=config['sheet_name']
        )
        app.register_blueprint(blueprint, url_prefix=config['url_prefix'])
