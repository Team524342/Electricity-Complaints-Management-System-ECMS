import speech_recognition as sr
import pyttsx3
import datetime
import json
import os
import re
from dataclasses import dataclass, asdict
from typing import List, Dict
import threading
import time
import pandas as pd


VOICE_COMPLAINT_FILE = "data/voiceComplaint.xlsx"
os.makedirs('data', exist_ok=True)    

# Initialize Excel file if it doesn't exist
if not os.path.exists(VOICE_COMPLAINT_FILE):
    voice_complaint_df = pd.DataFrame(columns=[
        'complaint_id', 'customer_name', 'phone_number', 'address', 'complaint_type',
        'description', 'timestamp', 'priority', 'status'
    ])
    voice_complaint_df.to_excel(VOICE_COMPLAINT_FILE, index=False)

def load_voice_complaints():
    """Load complaints from Excel file"""
    if os.path.exists(VOICE_COMPLAINT_FILE):
        try:
            voice_complaint_df = pd.read_excel(VOICE_COMPLAINT_FILE)
            return voice_complaint_df
        except Exception as e:
            print(f"Error loading complaints: {e}")
            return pd.DataFrame(columns=[
                'complaint_id', 'customer_name', 'phone_number', 'address', 'complaint_type',
                'description', 'timestamp', 'priority', 'status'
            ])
    return pd.DataFrame(columns=[
        'complaint_id', 'customer_name', 'phone_number', 'address', 'complaint_type',
        'description', 'timestamp', 'priority', 'status'
    ])

def save_complaint_to_excel(complaint):
    """Save a single complaint to Excel file"""
    try:
        # Load existing complaints
        complaints_df = load_voice_complaints()
        
        # Convert complaint object to dictionary
        complaint_dict = asdict(complaint)
        
        # Create new row DataFrame
        new_complaint_df = pd.DataFrame([complaint_dict])
        
        # Concatenate with existing data
        updated_df = pd.concat([complaints_df, new_complaint_df], ignore_index=True)
        
        # Save to Excel
        updated_df.to_excel(VOICE_COMPLAINT_FILE, index=False)
        print(f"Complaint {complaint.complaint_id} saved to Excel successfully!")
        
    except Exception as e:
        print(f"Error saving complaint to Excel: {e}")

@dataclass
class Complaint:
    complaint_id: str
    customer_name: str
    phone_number: str
    address: str
    complaint_type: str
    description: str
    timestamp: str
    priority: str = "Medium"
    status: str = "Open"

class ElectricityComplaintSystem:
    def __init__(self):
        # Initialize speech recognition and text-to-speech
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.tts_engine = pyttsx3.init()
        
        # Configure TTS settings
        self.tts_engine.setProperty('rate', 150)  # Speed of speech
        self.tts_engine.setProperty('volume', 0.8)  # Volume level
        
        # Complaint storage (also maintain JSON backup)
        self.complaints_file = "complaints.json"
        self.complaints = self.load_complaints()
        
        # Complaint types and keywords
        self.complaint_types = {
            "power outage": ["outage", "blackout", "no power", "electricity gone", "power cut"],
            "voltage fluctuation": ["voltage", "fluctuation", "high voltage", "low voltage", "unstable"],
            "billing issue": ["bill", "billing", "overcharge", "payment", "meter reading"],
            "equipment fault": ["pole", "wire", "transformer", "meter", "equipment", "damaged"],
            "street light": ["street light", "lamp", "lighting", "dark", "bulb"],
            "new connection": ["new connection", "connection", "supply", "installation"]
        }
        
        print("Voice-Based Electricity Complaint System Initialized")
        self.speak("Welcome to the Electricity Complaint System. How can I help you today?")
    
    def speak(self, text: str):
        """Convert text to speech"""
        print(f"System: {text}")
        self.tts_engine.say(text)
        self.tts_engine.runAndWait()
    
    def listen(self, timeout=5, phrase_time_limit=10) -> str:
        """Listen for audio input and convert to text"""
        try:
            with self.microphone as source:
                print("Listening...")
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
            
            print("Processing speech...")
            text = self.recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.lower()
        
        except sr.WaitTimeoutError:
            return "timeout"
        except sr.UnknownValueError:
            return "unclear"
        except sr.RequestError as e:
            print(f"Speech recognition error: {e}")
            return "error"
    
    def get_voice_input(self, prompt: str, max_attempts=3) -> str:
        """Get voice input with retries"""
        self.speak(prompt)
        
        for attempt in range(max_attempts):
            response = self.listen()
            
            if response == "timeout":
                if attempt < max_attempts - 1:
                    self.speak("I didn't hear anything. Please try again.")
                else:
                    self.speak("No response received. Moving to next step.")
                    return ""
            elif response == "unclear":
                if attempt < max_attempts - 1:
                    self.speak("I couldn't understand. Please speak clearly.")
                else:
                    self.speak("Unable to understand. Please try again later.")
                    return ""
            elif response == "error":
                self.speak("There was an error processing your speech. Please try again later.")
                return ""
            else:
                return response
        
        return ""
    
    def classify_complaint(self, description: str) -> tuple:
        """Classify complaint type and priority based on description"""
        description_lower = description.lower()
        
        # Determine complaint type
        complaint_type = "general"
        for category, keywords in self.complaint_types.items():
            if any(keyword in description_lower for keyword in keywords):
                complaint_type = category
                break
        
        # Determine priority
        priority = "Medium"
        if any(word in description_lower for word in ["emergency", "urgent", "fire", "danger", "safety"]):
            priority = "High"
        elif any(word in description_lower for word in ["outage", "blackout", "no power"]):
            priority = "High"
        elif any(word in description_lower for word in ["billing", "bill", "payment"]):
            priority = "Low"
        
        return complaint_type, priority
    
    def generate_complaint_id(self) -> str:
        """Generate unique complaint ID"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        return f"EC{timestamp}"
    
    def register_complaint(self):
        """Register a new complaint through voice interaction"""
        self.speak("I'll help you register your electricity complaint. Let's start with your details.")
        
        # Get customer name
        name = self.get_voice_input("Please tell me your full name.")
        if not name:
            self.speak("Unable to get your name. Please try again later.")
            return
        
        # Get phone number
        phone = self.get_voice_input("Please tell me your phone number digit by digit.")
        if not phone:
            self.speak("Unable to get your phone number. Please try again later.")
            return
        
        # Clean phone number (extract digits)
        phone_digits = re.findall(r'\d+', phone)
        phone_number = ''.join(phone_digits) if phone_digits else phone
        
        # Get address
        address = self.get_voice_input("Please tell me your complete address.")
        if not address:
            self.speak("Unable to get your address. Please try again later.")
            return
        
        # Get complaint description
        self.speak("Now, please describe your electricity problem in detail.")
        description = self.get_voice_input("Please describe your complaint.", max_attempts=2)
        if not description:
            self.speak("Unable to get complaint description. Please try again later.")
            return
        
        # Classify complaint
        complaint_type, priority = self.classify_complaint(description)
        
        # Create complaint
        complaint_id = self.generate_complaint_id()
        complaint = Complaint(
            complaint_id=complaint_id,
            customer_name=name,
            phone_number=phone_number,
            address=address,
            complaint_type=complaint_type,
            description=description,
            timestamp=datetime.datetime.now().isoformat(),
            priority=priority,
            status="Open"
        )
        
        # Store complaint in memory
        self.complaints.append(complaint)
        
        # Save to JSON (backup)
        self.save_complaints_to_json()
        
        # Save to Excel
        save_complaint_to_excel(complaint)
       
        # Confirm registration
        self.speak(f"Your complaint has been registered successfully. Your complaint ID is {complaint_id}. "
                  f"The complaint type is {complaint_type} with {priority.lower()} priority. "
                  f"We will contact you at {phone_number} for updates.")
        
        print(f"\nComplaint Registered:")
        print(f"ID: {complaint_id}")
        print(f"Name: {name}")
        print(f"Phone: {phone_number}")
        print(f"Address: {address}")
        print(f"Type: {complaint_type}")
        print(f"Priority: {priority}")
        print(f"Description: {description}")
    
    def check_complaint_status(self):
        """Check status of existing complaint"""
        complaint_id = self.get_voice_input("Please tell me your complaint ID.")
        
        if not complaint_id:
            self.speak("Unable to get complaint ID. Please try again later.")
            return
        
        # Extract complaint ID from speech
        complaint_id = complaint_id.upper().replace(" ", "")
        
        # Check in Excel file first
        try:
            complaints_df = load_voice_complaints()
            matching_complaints = complaints_df[complaints_df['complaint_id'].str.contains(complaint_id, case=False, na=False)]
            
            if not matching_complaints.empty:
                complaint_row = matching_complaints.iloc[0]
                self.speak(f"Found your complaint. Complaint ID {complaint_row['complaint_id']}. "
                         f"Type: {complaint_row['complaint_type']}. "
                         f"Status: {complaint_row['status']}. "
                         f"Priority: {complaint_row['priority']}. "
                         f"Registered on {str(complaint_row['timestamp'])[:10]}.")
                return
        except Exception as e:
            print(f"Error checking Excel file: {e}")
        
        # Fallback to memory search
        complaint = None
        for c in self.complaints:
            if complaint_id in c.complaint_id.upper():
                complaint = c
                break
        
        if complaint:
            self.speak(f"Found your complaint. Complaint ID {complaint.complaint_id}. "
                     f"Type: {complaint.complaint_type}. "
                     f"Status: {complaint.status}. "
                     f"Priority: {complaint.priority}. "
                     f"Registered on {complaint.timestamp[:10]}.")
        else:
            self.speak("Sorry, I couldn't find a complaint with that ID. Please check and try again.")
    
    def load_complaints(self) -> List[Complaint]:
        """Load complaints from JSON file (backup)"""
        if os.path.exists(self.complaints_file):
            try:
                with open(self.complaints_file, 'r') as f:
                    data = json.load(f)
                return [Complaint(**item) for item in data]
            except Exception as e:
                print(f"Error loading complaints from JSON: {e}")
        return []
    
    def save_complaints_to_json(self):
        """Save complaints to JSON file (backup)"""
        try:
            with open(self.complaints_file, 'w') as f:
                json.dump([asdict(complaint) for complaint in self.complaints], f, indent=2)
        except Exception as e:
            print(f"Error saving complaints to JSON: {e}")
    
    def show_menu(self):
        """Display menu options"""
        menu_text = """
        Available options:
        1. Register new complaint
        2. Check complaint status
        3. View all complaints
        4. Exit
        
        Please say the number of your choice or say:
        - 'register' or 'new complaint' for option 1
        - 'status' or 'check status' for option 2
        - 'view all' or 'show all' for option 3
        - 'exit' or 'quit' to exit
        """
        print(menu_text)
        self.speak("What would you like to do? You can register a new complaint, check complaint status, view all complaints, or exit.")
    
    def view_all_complaints(self):
        """View all complaints from Excel file"""
        try:
            complaints_df = load_voice_complaints()
            if complaints_df.empty:
                self.speak("No complaints found in the system.")
                return
            
            print("\n" + "="*80)
            print("ALL COMPLAINTS")
            print("="*80)
            
            for index, row in complaints_df.iterrows():
                print(f"\nComplaint #{index + 1}")
                print(f"ID: {row['complaint_id']}")
                print(f"Customer: {row['customer_name']}")
                print(f"Phone: {row['phone_number']}")
                print(f"Address: {row['address']}")
                print(f"Type: {row['complaint_type']}")
                print(f"Priority: {row['priority']}")
                print(f"Status: {row['status']}")
                print(f"Description: {row['description']}")
                print(f"Timestamp: {row['timestamp']}")
                print("-" * 40)
            
            total_complaints = len(complaints_df)
            self.speak(f"Found {total_complaints} complaints in total. Details are displayed on screen.")
            
        except Exception as e:
            print(f"Error viewing complaints: {e}")
            self.speak("Sorry, there was an error retrieving the complaints.")
    
    def process_menu_choice(self, choice: str) -> bool:
        """Process menu choice and return False to exit"""
        choice = choice.lower().strip()
        
        if any(word in choice for word in ['1', 'register', 'new complaint', 'complaint']):
            self.register_complaint()
        elif any(word in choice for word in ['2', 'status', 'check status', 'check']):
            self.check_complaint_status()
        elif any(word in choice for word in ['3', 'view all', 'show all', 'all complaints']):
            self.view_all_complaints()
        elif any(word in choice for word in ['4', 'exit', 'quit', 'bye']):
            self.speak("Thank you for using the Electricity Complaint System. Have a good day!")
            return False
        else:
            self.speak("I didn't understand your choice. Please try again.")
        
        return True
    
    def run(self):
        """Main application loop"""
        try:
            while True:
                self.show_menu()
                choice = self.get_voice_input("What would you like to do?")
                
                if not choice:
                    self.speak("No input received. Please try again.")
                    continue
                
                if not self.process_menu_choice(choice):
                    break
                
                # Brief pause before showing menu again
                time.sleep(1)
                
        except KeyboardInterrupt:
            self.speak("System shutting down. Goodbye!")
        except Exception as e:
            print(f"System error: {e}")
            self.speak("Sorry, there was a system error. Please try again later.")

def main():
    """Main function to run the complaint system"""
    print("Starting Voice-Based Electricity Complaint System...")
    print("Make sure your microphone is connected and working.")
    print("Press Ctrl+C to exit at any time.")
    
    try:
        system = ElectricityComplaintSystem()
        system.run()
    except Exception as e:
        print(f"Failed to start system: {e}")
        print("Please ensure you have the required packages installed:")
        print("pip install speechrecognition pyttsx3 pyaudio pandas openpyxl")