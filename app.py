from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, jsonify
import pandas as pd
import os
from datetime import datetime, timedelta
import uuid
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from excel_handler import export_complaints_excel, backup_database, import_complaints_from_excel
from biil import check_payment_status
from flask import current_app
import logging
import speech_recognition as sr
from pydub import AudioSegment
import tempfile
import wave
import io
from excel_editor_multi import register_excel_editors
from send_email import send_email_smtp ;
import re
from voice22 import main


# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "default_secret_key")  # Use environment variable for secret key

# File paths
UPLOAD_FOLDER = 'uploads'
COMPLAINT_FILE = 'data/complaints.xlsx'
USER_FILE = 'data/users.xlsx'
TECHNICIAN_FILE = "data/technician.xlsx"

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('data', exist_ok=True)

# Initialize Excel files if they don't exist
if not os.path.exists(COMPLAINT_FILE):
    complaints_df = pd.DataFrame(columns=[
        'complaint_id', 'user_id', 'category', 'description',
        'location', 'submission_date', 'status', 'assigned_to',
        'attachment_path', 'resolution_notes', 'resolution_date'
    ])
    complaints_df.to_excel(COMPLAINT_FILE, index=False)

def load_technician():
    """Load technicians from Excel file"""
    if os.path.exists(TECHNICIAN_FILE):
        technicians_df = pd.read_excel(TECHNICIAN_FILE)
        technicians_df['technician_id'] = technicians_df['technician_id'].astype(str)  # Ensure technician_id is a string
        return technicians_df
    return pd.DataFrame(columns=[
        'technician_id', 'fullName', 'aadhar', 'email', 'phone',
        'address', 'password', 'role'
    ])

if not os.path.exists(TECHNICIAN_FILE):
    technician_df = pd.DataFrame(columns=[
        'technician_id', 'fullName', 'aadhar', 'email', 'phone',
        'address', 'password', 'role'
    ])
    technician_data = {
        'technician_id': str(uuid.uuid4()),
        'fullName': 'Technician Name',
        'aadhar': '123456789012',
        'email': 'technician@example.com',
        'phone': '9876543210',
        'address': 'Technician Address',
        'password': generate_password_hash('password123'),  # Hash the password
        'role': 'technician'
    }
    technician_df = pd.concat([technician_df, pd.DataFrame([technician_data])], ignore_index=True)
    technician_df.to_excel(TECHNICIAN_FILE, index=False)

if not os.path.exists(USER_FILE):
    users_df = pd.DataFrame(columns=[
        'user_id', 'fullName', 'aadhar', 'email', 'phone',
        'address', 'password', 'role', 'registration_date'
    ])
    admin_data = {
        'user_id': str(uuid.uuid4()),
        'fullName': '1Admin User',
        'aadhar': '000000000000',
        'email': 'admin@123.com',
        'phone': '1234567891',
        'address': 'Admin Office',
        'password': generate_password_hash('admin123'),  # Hash the password
        'role': 'admin',
        'registration_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    users_df = pd.concat([users_df, pd.DataFrame([admin_data])], ignore_index=True)
    users_df.to_excel(USER_FILE, index=False)

# Ensure all technician passwords are hashed
technicians_df = load_technician()
technicians_df['password'] = technicians_df['password'].apply(
    lambda x: generate_password_hash(str(x)) if not isinstance(x, str) else x
)
technicians_df.to_excel(TECHNICIAN_FILE, index=False)

# Helper functions
# def load_complaints():
#     """Load complaints from Excel file"""
#     if os.path.exists(COMPLAINT_FILE):
#         return pd.read_excel(COMPLAINT_FILE)
#     return pd.DataFrame()

def load_complaints():
    """Load complaints from Excel file"""
    if os.path.exists(COMPLAINT_FILE):
        df = pd.read_excel(COMPLAINT_FILE)
        # Ensure complaint_id is string and clean
        df['complaint_id'] = df['complaint_id'].astype(str).str.strip()
        # Ensure user_id is string and clean
        df['user_id'] = df['user_id'].astype(str).str.strip()
        return df
    return pd.DataFrame()

def load_users():
    """Load users from Excel file"""
    if os.path.exists(USER_FILE):
        return pd.read_excel(USER_FILE)
    return pd.DataFrame(columns=[
        'user_id', 'fullName', 'aadhar', 'email', 'phone',
        'address', 'password', 'role', 'registration_date'
    ])



def save_complaint(complaint_data):
    """Save a new complaint to Excel file"""
    complaints_df = load_complaints()
    new_complaint = pd.DataFrame([complaint_data])
    updated_df = pd.concat([complaints_df, new_complaint], ignore_index=True)
    updated_df.to_excel(COMPLAINT_FILE, index=False)

def save_user(user_data):
    """Save a new user to Excel file"""
    users_df = load_users()
    new_user = pd.DataFrame([user_data])
    updated_df = pd.concat([users_df, new_user], ignore_index=True)
    updated_df.to_excel(USER_FILE, index=False)

def update_complaint_status(complaint_id, status, notes=None):
    """Update complaint status and notes"""
    complaints_df = load_complaints()
    idx = complaints_df.index[complaints_df['complaint_id'] == complaint_id].tolist()
    if idx:
        complaints_df.at[idx[0], 'status'] = status
        if notes:
            complaints_df.at[idx[0], 'resolution_notes'] = notes
            complaints_df.at[idx[0], 'resolution_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        complaints_df.to_excel(COMPLAINT_FILE, index=False)
        return True
    return False

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        
        if not username or not password:
            return render_template('login.html', error_message="Please fill in all fields")
        
        
        users_df = load_users()
        
        # Check login credentials
        user = users_df[
            ((users_df['email'] == username) | 
             (users_df['aadhar'] == username) | 
             (users_df['phone'] == username))
        ]
        
        if not user.empty and check_password_hash(user.iloc[0]['password'], password):
            user_data = user.iloc[0].to_dict()
            session['user_id'] = user_data['user_id']
            session['username'] = user_data['email']
            session['role'] = user_data['role']
            session.permanent = True
            app.permanent_session_lifetime = timedelta(minutes=30)  # Session timeout
            
            flash('Login successful!', 'success')
            
            if user_data['role'] == 'admin':
                return redirect(url_for('admin_dashboard'))
            
            else:
                return redirect(url_for('user_dashboard'))
            
        else:
            # flash('Invalid credentials. Please try again.', 'danger')

             return render_template('login.html', 
                                 error_message="Invalid credentials. Please try again.",
                                 username=username)  # Preserve username input
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out', 'info')
    return redirect(url_for('index'))

@app.route('/submit_complaint', methods=['GET', 'POST'])
def submit_complaint():
    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
    # Check if the user has paid their previous bill
    if check_payment_status(session.get('username')) != 'paid':
        flash('You must pay your pending bills before submitting a complaint.', 'danger')
        return render_template('unpaid.html')
    
    # Initialize variables for voice transcript
    voice_transcript = session.pop('voice_transcript', None)
    # main()

    complaints_df = load_complaints()
    if complaints_df.empty:
        new_complaint_id = "CID0001"
    else:
        last_complaint_id = complaints_df['complaint_id'].iloc[-1]
        last_complaint_number = int(re.findall(r'\d+', last_complaint_id)[0])
        new_complaint_number = last_complaint_number + 1 
        new_complaint_id = f"CID{str(new_complaint_number).zfill(4)}"   
    
    if request.method == 'POST':
        category = request.form['category']
        description = request.form['description']
        location = request.form['location']
        
        # Handle file upload
        attachment_path = ''
        if 'attachment' in request.files and request.files['attachment'].filename:
            file = request.files['attachment']
            filename = secure_filename(file.filename)
            unique_filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
            file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
            file.save(file_path)
            attachment_path = file_path
        
        # Create complaint data
        complaint_data = {
            'complaint_id': new_complaint_id,
            'user_id': str(session['user_id']),
            'category': category,
            'description': description,
            'location': location,
            'submission_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'status': 'Open',
            'assigned_to': '',
            'attachment_path': attachment_path,
            'resolution_notes': '',
            'resolution_date': '',
            'voice_complaint': True if 'voice_used' in request.form else False
        }
        
        save_complaint(complaint_data)
        flash('Complaint submitted successfully!', complaint_data['complaint_id'])
        return redirect(url_for('user_dashboard'))
    
    return render_template('submit_complaint.html', voice_transcript=voice_transcript)

@app.route('/register', methods=['GET', 'POST'])
def register():
    # Generate new user ID
    users_df = load_users()
    if users_df.empty:
        new_user_id = "UID0001"
    else:
        last_user_id = users_df['user_id'].iloc[-1]
        last_user_number = int(re.findall(r'\d+', last_user_id)[0])
        new_user_number = last_user_number + 1 
        new_user_id = f"UID{str(new_user_number).zfill(4)}"   

    if request.method == 'POST':
        fullName = request.form['fullName']
        aadhar = request.form['aadhar']
        email = request.form['email']
        phone = request.form['phone']
        address = request.form['address']
        password = request.form['password']
        confirm_password = request.form['confirm_password']
        
        # Validate Aadhar and Phone Number
        if not aadhar.isdigit() or len(aadhar) != 12:
            flash('Aadhar must be a 12-digit number.', 'danger')
            return render_template('register.html')
        if not phone.isdigit() or len(phone) != 10:
            flash('Phone number must be a 10-digit number.', 'danger')
            return render_template('register.html')
        
        # Validate password confirmation
        if password != confirm_password:
            flash('Passwords do not match!', 'danger')
            return render_template('register.html')
        
        users_df = load_users()
        
        # Check for unique constraints
        if not users_df[users_df['aadhar'] == aadhar].empty:
            flash('Aadhar Card number already exists!', 'danger')
            return render_template('register.html')
        if not users_df[users_df['email'] == email].empty:
            flash('Email ID already exists!', 'danger')
            return render_template('register.html')
        if not users_df[users_df['phone'] == phone].empty:
            flash('Phone number already exists!', 'danger')
            return render_template('register.html')
        
        # Register new user
        user_data = {
            'user_id': new_user_id,
            'fullName': fullName,
            'aadhar': aadhar,
            'email': email,
            'phone': phone,
            'address': address,
            'password': generate_password_hash(password),  # Hash the password
            'role': 'customer',
            'registration_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        save_user(user_data)
        flash('Registration successful! Please login.', 'success')
        return redirect(url_for('login'))
    
    return render_template('register.html')

@app.route('/add_technician', methods=['GET', 'POST'])
def add_technician():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    # Generate new technician ID
    technicians_df = load_technician()
    if technicians_df.empty:
        new_technician_id = "TID0001"
    else:
        last_technician_id = technicians_df['technician_id'].iloc[-1]
        last_technician_number = int(re.findall(r'\d+', last_technician_id)[0])
        new_technician_number = last_technician_number + 1 
        new_technician_id = f"TID{str(new_technician_number).zfill(4)}"
    
    if request.method == 'POST':
        fullName = request.form['fullName']
        aadhar = request.form['aadhar']
        email = request.form['email']
        phone = request.form['phone']
        address = request.form['address']
        password = request.form['password']
        
        # Validate Aadhar and Phone Number
        if not aadhar.isdigit() or len(aadhar) != 12:
            flash('Aadhar must be a 12-digit number.', 'danger')
            return redirect(url_for('add_technician'))
        if not phone.isdigit() or len(phone) != 10:
            flash('Phone number must be a 10-digit number.', 'danger')
            return redirect(url_for('add_technician'))
        
        # Load technicians data
        technicians_df = load_technician()
        
        # Check for unique constraints
        if not technicians_df[technicians_df['aadhar'] == aadhar].empty:
            flash('Aadhar Card number already exists!', 'danger')
            return redirect(url_for('add_technician'))
        if not technicians_df[technicians_df['email'] == email].empty:
            flash('Email ID already exists!', 'danger')
            return redirect(url_for('add_technician'))
        
        # Create new technician
        technician_data = {
            'technician_id': new_technician_id,
            'fullName': fullName,
            'aadhar': aadhar,
            'email': email,
            'phone': phone,
            'address': address,
            'password': generate_password_hash(password),  # Hash the password
            'role': 'technician'
        }
        
        # Add new technician to the DataFrame
        new_technician = pd.DataFrame([technician_data])
        updated_df = pd.concat([technicians_df, new_technician], ignore_index=True)
        updated_df.to_excel(TECHNICIAN_FILE, index=False)
        
        flash('Technician added successfully!', 'success')
        return redirect(url_for('admin_dashboard'))
    
    return render_template('add_technician.html')

@app.route('/user_dashboard')
def user_dashboard():
    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
    # Get user's complaints
    complaints_df = load_complaints()
    user_complaints = complaints_df[complaints_df['user_id'] == session['user_id']]
    
    return render_template('user_dashboard.html', complaints=user_complaints.to_dict('records'))


@app.route('/admin_dashboard')
def admin_dashboard():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    # Get all complaints
    complaints_df = load_complaints()
    complaints_df=complaints_df[::-1]
    
    # Get all technicians for assignment
    technicians_df = load_technician()
    
    # Statistics
    total_complaints = len(complaints_df)
    open_complaints = len(complaints_df[complaints_df['status'] == 'Open'])
    in_progress = len(complaints_df[complaints_df['status'] == 'In Progress'])
    resolved = len(complaints_df[complaints_df['status'] == 'Resolved'])
    
    stats = {
        'total': total_complaints,
        'open': open_complaints,
        'in_progress': in_progress,
        'resolved': resolved
    }
    
    # Convert DataFrames to dictionaries for template rendering
    complaints_list = complaints_df.to_dict('records')
    technicians_list = technicians_df.to_dict('records')

    show_count=3
    visible=complaints_list[:show_count]
    hidden=complaints_list[show_count:]
    
    # Monthly data for visualization
    monthly_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
    monthly_data = [12, 19, 3, 5, 2, 3]
    
    return render_template(
        'admin_dashboard.html',
        complaints=complaints_list,
        technicians=technicians_list,
        visible_rows=visible,
        hidden_rows=hidden,
        stats=stats,
        now=datetime.now(),
        monthly_labels=monthly_labels,
        monthly_data=monthly_data
    )
   
@app.route('/admin_dashboard/excelto')
def excelto():
    return render_template('excelto.html')

# Configure multiple Excel editors
editors = [
    {
        'name': 'Technician',
        'url_prefix': '/Technician',
        'excel_file': 'data/technician.xlsx',
        'sheet_name': 'Sheet1'
    },
    {
        'name': 'users',
        'url_prefix': '/Customer',
        'excel_file': 'data/users.xlsx',
        'sheet_name': 'Sheet1'
    },
    {
        'name': 'bills',
        'url_prefix': '/bill_records',
        'excel_file': 'data/Electricity_Bills_3Months.xlsx',
        'sheet_name': 'Sheet1'
    },{
        'name': 'complaints',
        'url_prefix': '/Complaints',
        'excel_file': 'data/complaints.xlsx',
        'sheet_name': 'Sheet1'
    },
]

 # Register all the Excel editors
register_excel_editors(app, editors)
# Add this new route to handle voice recording uploads
@app.route('/process_voice_complaint', methods=['POST'])
def process_voice_complaint():
    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
    if 'audio_data' not in request.files:
        flash('No audio file found', 'danger')
        return redirect(url_for('submit_complaint'))
    
    audio_file = request.files['audio_data']
    
    try:
        # Create a temporary file to store the audio
        with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_audio:
            audio_file.save(temp_audio.name)
        
        # Transcribe the audio file
        recognizer = sr.Recognizer()
        with sr.AudioFile(temp_audio.name) as source:
            audio_data = recognizer.record(source)
            transcript = recognizer.recognize_google(audio_data)
        
        # Clean up temporary file
        os.unlink(temp_audio.name)
        
        # Store the transcription in session for use in the complaint form
        session['voice_transcript'] = transcript
        
        return jsonify({
            'success': True,
            'transcript': transcript
        })
    
    except sr.UnknownValueError:
        return jsonify({
            'success': False,
            'error': 'Could not understand audio. Please try again.'
        })
    
    except sr.RequestError as e:
        return jsonify({
            'success': False,
            'error': f'Speech recognition service error: {str(e)}'
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Error processing audio: {str(e)}'
        })

@app.route('/view_complaint/<complaint_id>')
def view_complaint(complaint_id):
    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
    complaints_df = load_complaints()
    complaint = complaints_df[complaints_df['complaint_id'] == complaint_id]

    # Get all technicians for assignment
    technicians_df = load_technician()
    technicians_list = technicians_df.to_dict('records')
    
    if complaint.empty:
        flash('Complaint not found', 'danger')
        if session['role'] == 'admin':
            return redirect(url_for('admin_dashboard'))
        else:
            return redirect(url_for('user_dashboard'))
    
    return render_template('view_complaint.html', 
                         complaint=complaint.iloc[0].to_dict(), 
                         technicians=technicians_list)

@app.route('/assign_technician/<complaint_id>', methods=['POST'])
def assign_technician(complaint_id):
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    # Get the selected technician ID from the form
    technician_id = request.form.get('technician_id')
    
    if not technician_id:
        flash('No technician selected', 'danger')
        return redirect(url_for('view_complaint', complaint_id=complaint_id))
    
    try:
        # Load complaints and technician data
        complaints_df = load_complaints()
        technicians_df = load_technician()
        
        # Find complaint by ID
        complaint_idx = complaints_df.index[complaints_df['complaint_id'] == complaint_id].tolist()
        
        if not complaint_idx:
            flash('Complaint not found', 'danger')
            return redirect(url_for('admin_dashboard'))
        
        # Find technician by ID
        technician = technicians_df[technicians_df['technician_id'] == technician_id]
        
        if technician.empty:
            flash('Technician not found', 'danger')
            return redirect(url_for('view_complaint', complaint_id=complaint_id))
        
        # Update complaint with technician assignment
        complaints_df.at[complaint_idx[0], 'assigned_to'] = technician_id
        complaints_df.at[complaint_idx[0], 'technician_name'] = technician.iloc[0]['fullName']
        
        # Update status to "In Progress" if it's currently "Open"
        if complaints_df.at[complaint_idx[0], 'status'] == 'Open':
            complaints_df.at[complaint_idx[0], 'status'] = 'In Progress'
        
        # Save changes to Excel file
        complaints_df.to_excel(COMPLAINT_FILE, index=False)
        
        flash(f'Complaint assigned to {technician.iloc[0]["fullName"]}', 'success')
        return redirect(url_for('admin_dashboard', complaint_id=complaint_id))
        
    except Exception as e:
        flash(f'Error assigning technician: {str(e)}', 'danger')
        return redirect(url_for('view_complaint', complaint_id=complaint_id))
@app.route('/update_complaint/<complaint_id>', methods=['POST'])
def update_complaint(complaint_id):
    if 'user_id' not in session or session['role'] != 'admin' or session['role'] != 'technician':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    status = request.form['status']
    notes = request.form['notes']
    
    if update_complaint_status(complaint_id, status, notes):
        # --- Email notification logic ---
        # Load complaint and user info
        complaints_df = load_complaints()
        users_df = load_users()
        complaint = complaints_df[complaints_df['complaint_id'] == complaint_id]
        if not complaint.empty:
            complaint_row = complaint.iloc[0]
            user = users_df[users_df['user_id'] == complaint_row['user_id']]
            if not user.empty:
                user_row = user.iloc[0]
                receiver_email = user_row['email']
                customer_name = user_row['fullName']
                complaint_date = complaint_row['submission_date']
                status = complaint_row['status']
                resolution_time = complaint_row.get('resolution_date', '')
                support_contact = "1800-123-456"
                complaint_id_str = complaint_row['complaint_id']

                subject = "Electricity Complaint Status Update"
                message = (
                    f"Dear {customer_name},\n\n"
                    f"Thank you for contacting the Electricity Board.\n"
                    f"Your complaint (ID: {complaint_id_str}) registered on {complaint_date} has been updated.\n\n"
                    f"Current Status: {status}\n"
                    f"Resolution Time: {resolution_time}\n\n"
                    f"If you have further issues, please contact our support at {support_contact}.\n\n"
                    f"Thank you,\n"
                    f"Electricity Board Support Team"
                )
                html_message = f"""
                <html>
                <body>
                  <h2>Electricity Complaint Status Update</h2>
                  <p>Dear <strong>{customer_name}</strong>,</p>
                  <p>Thank you for contacting the Electricity Board.<br>
                  Your complaint details are as follows:</p>
                  <div>
                    <p><strong>Complaint ID:</strong> {complaint_id_str}<br>
                    <strong>Date Registered:</strong> {complaint_date}<br>
                    <strong>Status:</strong> {status}<br>
                    <strong>Resolution Time:</strong> {resolution_time}</p>
                  </div>
                  <p>If you have further issues, please contact our support at <strong>{support_contact}</strong>.</p>
                  <div>
                    Thank you,<br>
                    Electricity Board Support Team
                  </div>
                </body>
                </html>
                """
                # Use your sender email and app password here
                sender_email = "shoaib.@gmail"
                password = ""  # Use your app password

                send_email_smtp(
                    sender_email,
                    receiver_email,
                    subject,
                    message,
                    password,
                    html_message=html_message
                )
        flash('Complaint updated successfully and user notified!', 'success')
    else:
        flash('Failed to update complaint', 'danger')
    
    return redirect(url_for('admin_dashboard'))

@app.route('/generate_report')
def generate_report():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    complaints_df = load_complaints()
    
    # Create a summary report
    report_data = {
        'total_complaints': len(complaints_df),
        'open_complaints': len(complaints_df[complaints_df['status'] == 'Open']),
        'in_progress': len(complaints_df[complaints_df['status'] == 'In Progress']),
        'resolved': len(complaints_df[complaints_df['status'] == 'Resolved']),
        'category_counts': complaints_df['category'].value_counts().to_dict(),
        'recent_complaints': complaints_df.sort_values('submission_date', ascending=False).head(5).to_dict('records')
    }
    
    return render_template('report.html', report=report_data)

@app.route('/export_report_excel')
def export_report_excel():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    complaints_df = load_complaints()
    
    # Call the export function from excel_handler.py
    return export_complaints_excel(complaints_df)

@app.route('/backup_database')
def backup_db():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    message = backup_database()
    flash(message, 'success')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin_tools')
def admin_tools():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    return render_template('admin_tools.html')
 
@app.route('/import_complaints', methods=['POST'])
def import_complaints():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    if 'excel_file' not in request.files:
        flash('No file part', 'danger')
        return redirect(url_for('admin_tools'))
    
    file = request.files['excel_file']
    
    if file.filename == '':
        flash('No selected file', 'danger')
        return redirect(url_for('admin_tools'))
    
    if file:
        try:
            filename = secure_filename(file.filename)
            file_path = os.path.join('uploads', filename)
            file.save(file_path)
            
            imported_df, message = import_complaints_from_excel(file_path)
            
            if imported_df is not None:
                # Backup current data before import
                backup_database()
                
                # Merge with existing complaints
                complaints_df = load_complaints()
                # Avoid duplicates based on complaint_id
                existing_ids = set(complaints_df['complaint_id'])
                new_complaints = imported_df[~imported_df['complaint_id'].isin(existing_ids)]
                
                if not new_complaints.empty:
                    updated_df = pd.concat([complaints_df, new_complaints], ignore_index=True)
                    updated_df.to_excel(COMPLAINT_FILE, index=False)
                    flash(f"Successfully imported {len(new_complaints)} new complaints", 'success')
                else:
                    flash("No new complaints to import", 'info')
            else:
                flash(f"Import failed: {message}", 'danger')
            
            # Remove the temporary file
            os.remove(file_path)
        except Exception as e:
            flash(f"An error occurred during import: {str(e)}", 'danger')
        
    return redirect(url_for('admin_tools'))

@app.route('/update_profile', methods=['POST'])
def update_profile():
    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
    email = request.form['email']
    phone = request.form['phone']
    address = request.form['address']
    current_password = request.form['current_password']
    new_password = request.form['new_password']
    
    # Load users data
    users_df = load_users()
    
    # Find user by ID
    user_idx = users_df.index[users_df['user_id'] == session['user_id']].tolist()
    
    if not user_idx:
        flash('User not found', 'danger')
        return redirect(url_for('profile'))
    
    # Verify current password
    if not check_password_hash(users_df.at[user_idx[0], 'password'], current_password):
        flash('Current password is incorrect', 'danger')
        return redirect(url_for('profile'))
    
    # Update user information
    users_df.at[user_idx[0], 'email'] = email
    users_df.at[user_idx[0], 'phone'] = phone
    users_df.at[user_idx[0], 'address'] = address
    
    # Update password if provided
    if new_password:
        users_df.at[user_idx[0], 'password'] = generate_password_hash(new_password)
    
    # Save changes
    users_df.to_excel(USER_FILE, index=False)
    
    flash('Profile updated successfully', 'success')
    return redirect(url_for('profile'))

@app.route('/profile')
def profile():
    def get_initial():
                    names=users_df['fullName'].split()
                    first_name_initial=names[0][0].upper()
                    last_name_initial=names[-1][0].upper()
                    return f"{ first_name_initial}{last_name_initial}"


    if 'user_id' not in session:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    
   
       
    # Load users data
    users_df = load_users()
   
    # Find user by ID
    user = users_df[users_df['user_id'] == session['user_id']]
    
    if user.empty:
     flash('User not found', 'danger')
     return redirect(url_for('index')) 

    return render_template('profile.html', user=user.iloc[0].to_dict())

@app.route('/help')
def help_page():
    return render_template('help.html')

@app.route('/technicianLogin', methods=['GET', 'POST'])
def technicianLogin():
    if request.method == 'POST':
        login_identifier = request.form['login_identifier']
        password = request.form['password']
        
        technicians_df = load_technician()
        
        # Check login credentials
        technician = technicians_df[
            ((technicians_df['email'] == login_identifier) | 
             (technicians_df['technician_id'] == login_identifier))
        ]
        
        if not technician.empty:
            technician_data = technician.iloc[0].to_dict()
            
            # Ensure the password field is a string
            if isinstance(technician_data['password'], str):
                if check_password_hash(technician_data['password'], password):
                    session['user_id'] = technician_data['technician_id']
                    session['username'] = technician_data['email']
                    session['role'] = 'technician'
                    session.permanent = True
                    app.permanent_session_lifetime = timedelta(minutes=30)  # Session timeout
                    
                    flash('Technician login successful!', 'success')
                    return redirect(url_for('technician_dashboard'))
                else:
                    flash('Invalid credentials. Please try again.', 'danger')
            else:
                flash('Invalid password format in the database.', 'danger')
        else:
            flash('Invalid credentials. Please try again.', 'danger')
    
    return render_template('technicianLogin.html')

@app.route('/track_complaint ',methods=['GET', 'POST'])
def track_complaint():
    if request.method == 'POST':
        complaint_id = request.form['complaint_id']
        
        # Load complaints data
        complaints_df = load_complaints()
        
        # Find the complaint by ID
        complaint = complaints_df[complaints_df['complaint_id'] == complaint_id]
        
        if not complaint.empty:
            return render_template('track_complaint.html', complaint=complaint.iloc[0].to_dict())
        else:
            flash('Complaint not found', 'danger')
            return redirect(url_for('track_complaint'))
    
    return render_template('track_complaint.html')

@app.route('/technician_dashboard')
def technician_dashboard():
    if 'user_id' not in session or session['role'] != 'technician':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('technicianLogin'))
    
    # Load complaints assigned to the logged-in technician
    complaints_df = load_complaints()
    technician_complaints = complaints_df[complaints_df['assigned_to'] == session['user_id']]
    
    total_complaint=len(technician_complaints)
    open_complaint=len(technician_complaints[technician_complaints['status']=='open'])
    inProgress_complaint=len(technician_complaints[technician_complaints['status']=='Inprogress'])
    resolved_complaint=len(technician_complaints[technician_complaints['status']=='Resolved'])

    stats={
        'total':total_complaint,
        'open':open_complaint,
        'InProgress':inProgress_complaint,
        'Resolved':resolved_complaint,

    }


    return render_template('technician_dashboard.html',
                            complaints=technician_complaints.to_dict('records'),
                            stats=stats)
# Add this new route for technician_profile after the technician_dashboard route
@app.route('/technician_profile')
def technician_profile():
    if 'user_id' not in session or session['role'] != 'technician':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('technicianLogin'))
    
    # Load technician data
    technicians_df = load_technician()
    technician = technicians_df[technicians_df['technician_id'] == session['user_id']]
    
    if technician.empty:
        flash('Technician not found', 'danger')
        return redirect(url_for('technician_dashboard'))
    
    return render_template('technician_profile.html', technician=technician.iloc[0].to_dict())

# Add this new route for updating technician profile
@app.route('/update_technician_profile', methods=['POST'])
def update_technician_profile():
    if 'user_id' not in session or session['role'] != 'technician':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('technicianLogin'))
    
    email = request.form['email']
    phone = request.form['phone']
    address = request.form['address']
    current_password = request.form['current_password']
    new_password = request.form['new_password']
    
    # Load technicians data
    technicians_df = load_technician()
    
    # Find technician by ID
    tech_idx = technicians_df.index[technicians_df['technician_id'] == session['user_id']].tolist()
    
    if not tech_idx:
        flash('Technician not found', 'danger')
        return redirect(url_for('technician_profile'))
    
    # Verify current password
    if not check_password_hash(technicians_df.at[tech_idx[0], 'password'], current_password):
        flash('Current password is incorrect', 'danger')
        return redirect(url_for('technician_profile'))
    
    # Update technician information
    technicians_df.at[tech_idx[0], 'email'] = email
    technicians_df.at[tech_idx[0], 'phone'] = phone
    technicians_df.at[tech_idx[0], 'address'] = address
    
    # Update password if provided
    if new_password:
        technicians_df.at[tech_idx[0], 'password'] = generate_password_hash(new_password)
    
    # Save changes
    technicians_df.to_excel(TECHNICIAN_FILE, index=False)
    
    flash('Profile updated successfully', 'success')
    return redirect(url_for('technician_profile'))

@app.route('/update_technician_complaint/<complaint_id>', methods=['POST'])
def update_technician_complaint(complaint_id):
    if 'user_id' not in session or session['role'] != 'technician':
        flash('Please login first', 'warning')
        return redirect(url_for('technicianLogin'))
    
    status = request.form['status']
    notes = request.form['notes']
    
    # Load complaints
    complaints_df = load_complaints()
    
    # Check if this complaint is assigned to the logged-in technician
    complaint = complaints_df[complaints_df['complaint_id'] == complaint_id]
    if complaint.empty:
        flash('Complaint not found', 'danger')
        return redirect(url_for('technician_dashboard'))
    
    if complaint.iloc[0]['assigned_to'] != session['user_id']:
        flash('You are not authorized to update this complaint', 'danger')
        return redirect(url_for('technician_dashboard'))
    
    # Update complaint status and resolution notes
    if update_complaint_status(complaint_id, status, notes):
        flash('Complaint updated successfully!', 'success')
    else:
        flash('Failed to update complaint', 'danger')
    
    return redirect(url_for('technician_dashboard'))


# Route to manage technicians
@app.route('/manage_technicians')
def manage_technicians():
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    # Load technicians data
    technicians_df = load_technician()
    technicians_count = len(technicians_df)
    return render_template('manage_technicians.html', technicians=technicians_df.to_dict('records'),technicians_count=technicians_count)

@app.route('/edit_technician/<technician_id>', methods=['GET', 'POST'])
def edit_technician(technician_id):
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    technicians_df = load_technician()
    
    # Find technician by ID
    technician = technicians_df[technicians_df['technician_id'] == technician_id]
    
    if technician.empty:
        flash('Technician not found', 'danger')
        return redirect(url_for('manage_technicians'))
    
    if request.method == 'POST':
        fullName = request.form['fullName']
        aadhar = request.form['aadhar']
        email = request.form['email']
        phone = request.form['phone']
        address = request.form['address']
        
        # Validate Aadhar and Phone Number
        if not aadhar.isdigit() or len(aadhar) != 12:
            flash('Aadhar must be a 12-digit number.', 'danger')
            return redirect(url_for('edit_technician', technician_id=technician_id))
        if not phone.isdigit() or len(phone) != 10:
            flash('Phone number must be a 10-digit number.', 'danger')
            return redirect(url_for('edit_technician', technician_id=technician_id))
        
        # Update technician details
        tech_idx = technicians_df.index[technicians_df['technician_id'] == technician_id].tolist()[0]
        technicians_df.at[tech_idx, 'fullName'] = fullName
        technicians_df.at[tech_idx, 'aadhar'] = aadhar
        technicians_df.at[tech_idx, 'email'] = email
        technicians_df.at[tech_idx, 'phone'] = phone
        technicians_df.at[tech_idx, 'address'] = address
        
        # Save changes
        technicians_df.to_excel(TECHNICIAN_FILE, index=False)
        
        flash('Technician updated successfully!', 'success')
        return redirect(url_for('manage_technicians'))
    
    return render_template('edit_technician.html', technician=technician.iloc[0].to_dict())

@app.route('/delete_technician/<technician_id>', methods=['POST'])
def delete_technician(technician_id):
    if 'user_id' not in session or session['role'] != 'admin':
        flash('Unauthorized access', 'danger')
        return redirect(url_for('login'))
    
    technicians_df = load_technician()
    
    # Find technician by ID
    tech_idx = technicians_df.index[technicians_df['technician_id'] == technician_id].tolist()
    
    if not tech_idx:
        flash('Technician not found', 'danger')
        return redirect(url_for('manage_technicians'))
    
    # Remove technician from the DataFrame
    technicians_df = technicians_df.drop(tech_idx[0])
    technicians_df.to_excel(TECHNICIAN_FILE, index=False)
    
    flash('Technician deleted successfully!', 'success')
    return redirect(url_for('manage_technicians'))


@app.route('/technician_dashboard/assign_complaint')
def assigned_complaint():
    return f"assigned no complaint"
@app.route('/reports')
def reports():
    # Example: get or build your report data here
    report = {
        "total_complaints": 0,
        "open_complaints": 0,
        "in_progress": 0,
        "resolved": 0,
        "category_counts": {},
        "recent_complaints": []
    }
    # Replace the above with your actual report data logic
    return render_template('report.html', report=report)

@app.route('/admin_profile')
def admin_profile():
    return render_template('admin_profile.html')


# Error handlers
@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html')

@app.errorhandler(500)
def server_error(e):
    return render_template('500.html')

@app.route('/start_voice_interaction')
def start_voice_interaction():
    # This will speak and listen using microphone
    main()
    return  redirect('user_dashboard') # Or

@app.route('/about')
def about():
    return render_template('about.html')

if __name__ == '__main__':
    app.run(debug=True)
