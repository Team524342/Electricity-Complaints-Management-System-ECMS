import speech_recognition as sr
import openpyxl
from datetime import datetime
import time  # For adding pauses

# Define the Excel file name and sheet name
EXCEL_FILE = "complaints.xlsx"
SHEET_NAME = "ComplaintsData"

def listen_for_complaint():
    """Provides introductory messages and then listens for voice input."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Welcome to the Complaint Recording System.")
        time.sleep(1)  # Pause for 1 second
        print("Please state your complaint clearly after the beep.")
        time.sleep(1)
        print("\007")  # Beep sound (may not work on all systems/terminals)
        r.adjust_for_ambient_noise(source)  # Adjust for background noise
        try:
            audio = r.listen(source)
            print("Processing...")
            complaint_text = r.recognize_google(audio)  # Use Google Speech Recognition
            print(f"You said: {complaint_text}")
            return complaint_text
        except sr.UnknownValueError:
            print("Could not understand audio.")
            print("Please try again.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results from speech recognition service; {e}")
            print("Please check your internet connection.")
            return None

def save_complaint_to_excel(complaint):
    """Saves the complaint details to an Excel file."""
    if complaint:
        try:
            # Load the workbook or create a new one if it doesn't exist
            try:
                workbook = openpyxl.load_workbook(EXCEL_FILE)
                sheet = workbook[SHEET_NAME]
            except FileNotFoundError:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.title = SHEET_NAME
                # Add headers to the sheet
                sheet.append(["Timestamp", "Complaint"])

            # Get the current timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Append the complaint and timestamp as a new row
            sheet.append([timestamp, complaint])

            # Save the workbook
            workbook.save(EXCEL_FILE)
            print(f"\nComplaint saved to '{EXCEL_FILE}', sheet '{SHEET_NAME}'.")
        except Exception as e:
            print(f"Error saving to Excel: {e}")
    else:
        print("No complaint to save.")

if __name__ == "__main__":
    complaint = listen_for_complaint()
    if complaint:  # Only save if a complaint was successfully transcribed
        save_complaint_to_excel(complaint)
