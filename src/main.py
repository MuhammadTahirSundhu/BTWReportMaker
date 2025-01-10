# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main.ui'
##
## Created by: Qt User Interface Compiler version 6.8.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QProgressBar
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QListWidget

from PySide6.QtGui import (QCursor,QFont, QPixmap)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFrame, QHBoxLayout,
    QLabel, QListWidget, QListWidgetItem, QMainWindow,
    QProgressBar, QPushButton, QVBoxLayout,
    QWidget)
from PySide6.QtWidgets import QFileDialog
import fitz
import pdfplumber
import re
import random
import os
import sys
import json
from datetime import datetime, time
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import sys

# Set the standard output encoding to utf-8
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding='utf-8')
else:
    print("Warning: Unable to reconfigure sys.stdout encoding.")


CONFIG_FILE = "config.json"

# Basic Row class with minimal requirements
class Row:
    def __init__(self, date=None, time=None, status=None):
        self.date = date
        self.time = time  # Optional
        self.status = status

    def is_valid(self):
        """A row is valid if it has a date and status is Complete"""
        return self.date is not None and self.status == "Complete"

# Define row types and their coordinates

ROW_LOOKUP = [
    {"row": 1, "type": "dayfull", "y": 137},
    {"row": 2, "type": "nightfull", "y": 155},
    {"row": 3, "type": "dayfull", "y": 173},
    {"row": 4, "type": "dayfull", "y": 187},
    {"row": 5, "type": "nightfull", "y": 203},
    {"row": 6, "type": "dayhalf", "y": 219},
    {"row": 7, "type": "nighthalf", "y": 232},
    {"row": 8, "type": "dayfull", "y": 250},
    {"row": 9, "type": "dayfull", "y": 265},
    {"row": 10, "type": "dayfull", "y": 280},
    {"row": 11, "type": "nightfull", "y": 296},
    {"row": 12, "type": "dayfull", "y": 315},
    {"row": 13, "type": "dayfull", "y": 330},
    {"row": 14, "type": "nightfull", "y": 346},
    {"row": 15, "type": "dayhalf", "y": 362},
    {"row": 16, "type": "nighthalf", "y": 377},
    {"row": 17, "type": "dayfull", "y": 392},
    {"row": 18, "type": "nightfull", "y": 407},
    {"row": 19, "type": "dayfull", "y": 424},
    {"row": 20, "type": "dayfull", "y": 439},
    {"row": 21, "type": "dayfull", "y": 454},
    {"row": 22, "type": "nightfull", "y": 469},
    {"row": 23, "type": "dayfull", "y": 485},
    {"row": 24, "type": "dayfull", "y": 500},
    {"row": 25, "type": "dayfull", "y": 515},
    {"row": 26, "type": "daynightfull", "y": 530},
    {"row": 27, "type": "nightfull", "y": 546},
    {"row": 28, "type": "dayfirst", "y": 562},
    {"row": 29, "type": "dayfirst", "y": 576},
    {"row": 30, "type": "dayfirst", "y": 590},
    {"row": 31, "type": "daynightfull", "y": 605},
    {"row": 32, "type": "nightfull", "y": 620}
]

for row in ROW_LOOKUP:
    row["coordinates"] = {
        "date_x": 200,
        "date_y": row["y"],
        "time_x": 255,
        "time_y": row["y"],
        "img_x": 420,
        "img_y": row["y"] - 4,
        "text_x": 510,
        "text_y": row["y"],
    }

# Helper functions

def load_config():
    """Load configuration from file if it exists"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return None
    return None

def save_config(config):
    """Save configuration to file"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

def get_path(path, must_exist=True):
    """Validate path input from predefined variables"""
    path = path.strip().strip('"')  # Remove quotes if present
    if not path:
        raise ValueError("Path cannot be empty.")

    if must_exist and not os.path.exists(path):
        raise FileNotFoundError(f"Path does not exist: {path}")

    return path

def setup_paths(templatePdf,signaturefile,outputdir):
    """Set up required paths using predefined variables"""
    print("\nSetting up the required file paths using predefined variables...")

    # Predefined variables
    var1 = templatePdf  # Example path for template PDF
    var2 = signaturefile  # Example path for signature image
    var3 = outputdir  # Example path for output directory

    # Validate and assign paths
    template_path = get_path(var1)
    if not template_path.endswith('.pdf'):
        raise ValueError("Template must be a PDF file.")

    signature_path = get_path(var2)
    if not signature_path.lower().endswith(('.png', '.jpg', '.jpeg')):
        raise ValueError("Signature must be an image file (PNG, JPG).")

    output_dir = get_path(var3, must_exist=False)

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    config = {
        "template_path": template_path,
        "signature_path": signature_path,
        "output_dir": output_dir
    }

    save_config(config)
    return config

def extract_pdf_text(pdf_path):
    """Extract text from a PDF file using pdfplumber"""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def write_text_to_file(text, output_path):
    """Write text content to a file"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(text)

def parse_student_name(text):
    """Extract student name from the PDF text"""
    match = re.search(r'Student:\s*(.*?)\s*Address:', text)
    if match:
        return match.group(1).strip()
    return ""

def parse_license_number(text):
    """Extract license number from the PDF text"""
    match = re.search(r'LDL:\s*(\d+)', text)
    if match:
        return match.group(1).strip()
    return ""

def parse_appointments(text):
    """Extract appointments from the PDF text"""
    appointments = []
    lines = text.split('\n')
    for line in lines:
        if 'Teen BTW' in line:
            appointments.append(line.strip())
    return appointments

# Function to extract all text from the PDF
def extract_pdf_text(file_path):
    all_text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            all_text += page.extract_text() + "\n"
    return all_text

# Function to write extracted text to a file
def write_text_to_file(text, output_file):
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(text)

# Function to parse student name
def parse_student_name(text):
    match = re.search(r"Student:\s*(.+?)\s*DOI:", text)
    if match:
        return match.group(1).strip()
    return None

# Function to parse license number
def parse_license_number(text):
    match = re.search(r"LDL:\s*(\d+)", text)
    if match:
        return match.group(1).strip()
    return None

def test_parse_appointments(text):
    appointments = parse_appointments(text)
    print(f"\nFound {len(appointments)} valid appointments:")
    for i, app in enumerate(appointments, 1):
        print(f"\nAppointment {i}:")
        print(f"Date: {app.date}")
        print(f"Time: {app.time or 'No time specified'}")
        print(f"Status: {app.status}")
        print(f"Is Valid: {app.is_valid()}")
    return appointments

# Function to parse appointments
def parse_appointments(text):
    """Parse appointments from text, returning list of valid Row objects"""
    appointments = []
    lines = text.split('\n')
    
    for line in lines:
        if 'Teen BTW' in line:
            # Extract date
            date_match = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', line)
            date = date_match.group(1) if date_match else None
            
            # Extract time (optional)
            time_match = re.search(r'\b(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?)\b', line)
            time = time_match.group(1) if time_match else None
            
            # Extract status
            status = 'Complete' if 'Complete' in line else None
            
            # Create row only if we have required fields
            if date and status == 'Complete':
                row = Row(date=date, time=time, status=status)
                appointments.append(row)
    
    return appointments

# Function to extract date, time, and print debug info for appointments
def parse_time_debug(appointments):
    for day, details in appointments:
        match = re.search(r"\b(\d{2}:\d{2})\b", details)
        if match:
            time = match.group(1)
            print(f"Processing line: {day} {details}")
            print(f"Extracted Time: {time}")
        else:
            print(f"Warning: Could not extract time for line: {day} {details}")

def parse_detailed_appointments(appointments):
    """Process appointments, no need to unpack since they're already Row objects"""
    parsed_rows = []
    for appointment in appointments:
        # Add to parsed rows if it's a valid appointment
        if appointment.is_valid():
            parsed_rows.append(appointment)
            
            # Debugging output
            print(f"\nProcessed appointment:")
            print(f"Date: {appointment.date}")
            print(f"Time: {appointment.time or 'No time specified'}")
            print(f"Status: {appointment.status}")
    
    return parsed_rows

def analyze_parsed_rows(parsed_rows, student_name):
    """Analyze parsed rows and return a summary"""
    # Count total complete rows
    complete_rows = [row for row in parsed_rows if row.status == "Complete"]
    total_complete = len(complete_rows)
    
    # Get expressway lessons (rows 8-10)
    expressway_lessons = complete_rows[7:10] if len(complete_rows) > 7 else []
    
    # Find duplicates by comparing dates
    date_counts = {}
    for row in complete_rows:
        date_counts[row.date] = date_counts.get(row.date, 0) + 1
    duplicates = sum(count - 1 for count in date_counts.values() if count > 1)
    
    # Check if times are present
    rows_with_times = [row for row in complete_rows if row.time is not None]
    needs_random_times = len(rows_with_times) == 0
    
    summary = f"\nSummary for {student_name}:"
    summary += f"\n- Total complete lessons: {total_complete}"
    summary += f"\n- Expressway lessons: {len(expressway_lessons)}"
    if duplicates > 0:
        summary += f"\n- Duplicate dates found: {duplicates}"
    if needs_random_times:
        summary += "\n- No times found in log, random times will be assigned"
    
    return summary

def format_time_range(time_str, row_type, second_half=False):
    """
    Format time range based on row type
    Args:
        time_str: Time string
        row_type: Type of row
        second_half: If True, adds 30 minutes to start time for second half of pair
    """
    if "daynightfull" in row_type:
        return "5:30-6:30p"
        
    # If no time provided, generate random time
    if not time_str:
        hour = random.randint(12, 19)  # Defaulting to afternoon times for random generation
        minute = random.choice([0, 30])
        time_str = f"{hour}:{minute:02d}"
    
    try:
        # Parse the time
        if "PM" in time_str.upper() or "AM" in time_str.upper():
            time_parts = time_str.replace("PM", "").replace("AM", "").strip().split(":")
            hour = int(time_parts[0])
            minute = int(time_parts[1])
            if "PM" in time_str.upper() and hour != 12:
                hour += 12
        else:
            time_parts = time_str.split(":")
            hour = int(time_parts[0])
            minute = int(time_parts[1])
        
        # Adjust time if this is second half of a pair
        if second_half:
            minute += 30
            if minute >= 60:
                hour += 1
                minute -= 60
        
        # Convert to 12-hour format and determine if morning or afternoon
        is_morning = hour < 12
        output_hour = hour if hour <= 12 else hour - 12
        
        # Calculate end time
        if "half" in row_type:
            end_hour = output_hour
            end_minute = minute + 30
            if end_minute >= 60:
                end_hour += 1
                end_minute -= 60
        else:
            end_hour = output_hour + 1
            end_minute = minute
        
        # Determine if end time is morning or afternoon
        end_is_morning = (hour + (1 if "half" not in row_type else 0)) < 12
            
        # Format the time range with appropriate suffix
        return f"{output_hour}:{minute:02d}-{end_hour}:{end_minute:02d}{'a' if is_morning else 'p'}"
    
    except (ValueError, IndexError):
        return "12:00-1:00p"

def get_output_filename(input_path, output_dir):
    """Generate output filename based on input filename"""
    base_name = os.path.basename(input_path)
    student_name = base_name.split('BTWAttendanceHistoryReport')[0].strip()
    date_str = datetime.now().strftime('%Y%m%d')
    return os.path.join(output_dir, f"{student_name}Report_{date_str}.pdf")

def process_single_file(attendance_file, template_path, signature_path, output_dir):
    """Modified to accept paths as parameters"""
    try:
        output_file = os.path.join(
            output_dir, 
            f"{os.path.basename(attendance_file).split('BTWAttendanceHistoryReport')[0].strip()}Report_{datetime.now().strftime('%Y%m%d')}.pdf"
        )
        
        raw_text = extract_pdf_text(attendance_file)
        student_name = parse_student_name(raw_text)
        student_license = parse_license_number(raw_text)
        appointments = parse_appointments(raw_text)
        parsed_rows = parse_detailed_appointments(appointments)
        
        summary = analyze_parsed_rows(parsed_rows, student_name)
        print(summary)
        
        row_dictionary = create_row_dictionary(parsed_rows, signature_path)
        
        doc = fitz.open(template_path)
        doc = place_rows_on_pdf(doc, row_dictionary, signature_path, student_name, student_license)
        doc.save(output_file)
        doc.close()
        
        print(f"✓ Processed: {os.path.basename(attendance_file)}")
        return True, summary
    except Exception as e:
        print(f"✗ Error processing {os.path.basename(attendance_file)}: {str(e)}")
        return False, None

def is_night_time(time_str):
    """
    Determine if a time is during night (6 PM or later)
    Returns True if night time, False if daytime, None if can't determine
    """
    if not time_str:
        return None
        
    try:
        # Handle various time formats
        if "PM" in time_str.upper() or "AM" in time_str.upper():
            time_parts = time_str.replace("PM", "").replace("AM", "").strip().split(":")
            hour = int(time_parts[0])
            if "PM" in time_str.upper() and hour != 12:
                hour += 12
        else:
            hour = int(time_str.split(":")[0])
        
        return hour >= 18  # 6 PM or later
    except (ValueError, IndexError):
        return None

def is_half_hour_time(time_str):
    """Check if time is on the half hour"""
    if not time_str:
        return None
        
    try:
        # Extract minutes
        if ":" in time_str:
            minutes = int(time_str.split(":")[1].split()[0])
            return minutes == 30
    except (ValueError, IndexError):
        return None
    
    return False

def generate_random_time(is_night=False, is_half=False):
    """Generate random time based on requirements"""
    if is_night:
        hour = random.randint(18, 20)  # 6 PM to 8 PM
    else:
        hour = random.randint(12, 17)  # 12 PM to 5 PM
        
    minute = 30 if is_half else 0
    return f"{hour}:{minute:02d}"

def create_row_dictionary(parsed_rows, signature_path):
    row_dict = {}
    assigned_rows = set()
    used_dates = set()  # Track used dates to prevent duplicates
    
    # First handle expressway lessons (rows 28-30)
    completed_rows = [row for row in parsed_rows if row.status == "Complete"]
    
    # Get the first three unique dates
    expressway_lessons = []
    lesson_index = 7  # Start at index 7 (8th lesson)
    while len(expressway_lessons) < 3 and lesson_index < len(completed_rows):
        current_lesson = completed_rows[lesson_index]
        if current_lesson.date not in used_dates:
            expressway_lessons.append(current_lesson)
            used_dates.add(current_lesson.date)
        lesson_index += 1
    
    # If we still need more lessons after going through 8-10, continue through the rest
    if len(expressway_lessons) < 3:
        for lesson in completed_rows[lesson_index:]:
            if lesson.date not in used_dates:
                expressway_lessons.append(lesson)
                used_dates.add(lesson.date)
                if len(expressway_lessons) == 3:
                    break
    
    # Assign the expressway lessons to rows 28-30
    for idx, lesson in enumerate(expressway_lessons[:3]):
        target_row = 28 + idx
        row_type = ROW_LOOKUP[target_row - 1]["type"]
        formatted_time = format_time_range(lesson.time, row_type)
        
        row_dict[len(row_dict)] = {
            "date": lesson.date,
            "time": formatted_time,
            "signature": signature_path,
            "DLnum": "13831911",
            "coordinates": ROW_LOOKUP[target_row - 1]["coordinates"]
        }
        assigned_rows.add(target_row)
    
    # Process remaining lessons
    remaining_lessons = [row for row in completed_rows[10:] 
                        if row.date not in used_dates]  # Filter out used dates
    
    # Separate available rows by type
    available_rows = {
        "dayfull": [],
        "nightfull": [],
        "dayhalf": [],
        "nighthalf": [],
        "daynightfull": []
    }
    
    for row in ROW_LOOKUP:
        if row["row"] not in assigned_rows and row["row"] not in [28, 29, 30]:
            available_rows[row["type"]].append(row)
            
    def get_half_hour_pair(is_night=False):
        """
        Get a pair of half-hour rows (either both day or both night)
        Args:
            is_night: If True, get rows 7 & 16, if False get rows 6 & 15
        Returns:
            Tuple of (first_half_row, second_half_row) or (None, None)
        """
        if is_night:
            # Look for rows 7 and 16 (nighthalf)
            first_half = next((row for row in available_rows["nighthalf"] if row["row"] == 7), None)
            second_half = next((row for row in available_rows["nighthalf"] if row["row"] == 16), None)
        else:
            # Look for rows 6 and 15 (dayhalf)
            first_half = next((row for row in available_rows["dayhalf"] if row["row"] == 6), None)
            second_half = next((row for row in available_rows["dayhalf"] if row["row"] == 15), None)
            
        if first_half and second_half:
            # Remove both rows from available rows
            available_rows["nighthalf" if is_night else "dayhalf"] = [
                row for row in available_rows["nighthalf" if is_night else "dayhalf"]
                if row["row"] not in (first_half["row"], second_half["row"])
            ]
            return first_half, second_half
        return None, None
    
    # Process lessons with times first
    timed_lessons = [l for l in remaining_lessons if l.time is not None]
    
    # Try to pair up lessons for half-hour slots first
    used_in_pairs = set()
    for lesson in timed_lessons:
        if lesson.date in used_dates or lesson.date in used_in_pairs:
            continue
            
        is_night = is_night_time(lesson.time)
        # Get appropriate half-hour row pair if available
        first_half, second_half = get_half_hour_pair(is_night)
        
        if first_half and second_half:
            # Create two time slots for this date
            base_time = lesson.time if lesson.time else "12:00"
            
            # First half-hour slot
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": format_time_range(base_time, first_half["type"]),
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": first_half["coordinates"]
            }
            assigned_rows.add(first_half["row"])
            
            # Second half-hour slot (30 minutes later)
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": format_time_range(base_time, second_half["type"], second_half=True),
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": second_half["coordinates"]
            }
            assigned_rows.add(second_half["row"])
            
            used_in_pairs.add(lesson.date)
            used_dates.add(lesson.date)
            print(f"Paired {lesson.date} in half-hour slots {first_half['row']} and {second_half['row']}")
            continue
    
    # Process lessons with times first
    timed_lessons = [l for l in remaining_lessons if l.time is not None]
    
    # Try to pair up lessons for half-hour slots first
    used_in_pairs = set()
    for i, lesson in enumerate(timed_lessons):
        if lesson.date in used_dates or lesson.date in used_in_pairs:
            continue
            
        # Get half-hour row pair if available
        day_half, night_half = get_half_hour_pair()
        if day_half and night_half:
            # Create two time slots for this date
            base_time = lesson.time if lesson.time else "12:00"
            
            # First half-hour slot
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": format_time_range(base_time, "dayhalf"),
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": day_half["coordinates"]
            }
            assigned_rows.add(day_half["row"])
            
            # Second half-hour slot (30 minutes later)
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": format_time_range(base_time, "nighthalf", second_half=True),
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": night_half["coordinates"]
            }
            assigned_rows.add(night_half["row"])
            
            used_in_pairs.add(lesson.date)
            used_dates.add(lesson.date)
            continue
            
        # If no half-hour pairs available, process normally
        available_types = []
        for row_type, rows in available_rows.items():
            if rows:
                available_types.append(row_type)
        
        if available_types:
            row_type = random.choice(available_types)
            selected = available_rows[row_type].pop(0)
            formatted_time = format_time_range(lesson.time, row_type)
            
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": formatted_time,
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": selected["coordinates"]
            }
            used_dates.add(lesson.date)
    
    # Handle remaining lessons
    untimed_lessons = [l for l in remaining_lessons if l.time is None 
                      and l.date not in used_dates 
                      and l.date not in used_in_pairs]
                      
    available_rows_list = []
    for rows in available_rows.values():
        available_rows_list.extend(rows)
        
    for lesson in untimed_lessons:
        if available_rows_list:
            selected = available_rows_list.pop(0)
            random_time = generate_random_time("night" in selected["type"], 
                                            "half" in selected["type"])
            formatted_time = format_time_range(random_time, selected["type"])
            
            row_dict[len(row_dict)] = {
                "date": lesson.date,
                "time": formatted_time,
                "signature": signature_path,
                "DLnum": "13831911",
                "coordinates": selected["coordinates"]
            }
            
    return row_dict


def place_rows_on_pdf(doc, row_dictionary, image_path, student_name, student_license):
    print("\nDEBUG: Starting PDF placement")
    print(f"Total rows to place: {len(row_dictionary)}")
    
    page_index = 0
    page = doc[page_index]

    # Insert student info
    print(f"\nPlacing student info: {student_name}, {student_license}")
    page.insert_text((120, 99), student_name, fontsize=12, color=(0, 0, 0), rotate=0)
    page.insert_text((420, 99), student_license, fontsize=12, color=(0, 0, 0), rotate=0)

    # Place each row
    for idx, row_data in row_dictionary.items():
        coords = row_data["coordinates"]
        
        # Add random integer variance ONLY to image coordinates (±3 points)
        x_variance = random.randint(-5, 5)
        y_variance = random.randint(-3, 0)
        
        # Calculate varied image coordinates
        img_x = coords["img_x"]
        img_y = coords["img_y"]
        
        print(f"\nPlacing row {idx}:")
        print(f"Date: {row_data['date']} at ({coords['date_x']}, {coords['date_y']})")
        print(f"Time: {row_data['time']} at ({coords['time_x']}, {coords['time_y']})")
        print(f"Image at ({img_x}, {img_y}) [varied from ({coords['img_x']}, {coords['img_y']})]")
        
        # Create image rectangle with varied coordinates
        image_rect = fitz.Rect(
            img_x - 9, 
            img_y - 5, 
            img_x + 90,
            img_y + 11
        )
        page.insert_image(image_rect, filename=image_path)
        
        # Place text elements (fixed positions)
        page.insert_text(
            (coords["date_x"], coords["date_y"]+1), 
            row_data["date"], 
            fontsize=10, 
            color=(0, 0, 0), 
            rotate=0
        )
        page.insert_text(
            (coords["time_x"], coords["time_y"]), 
            row_data["time"], 
            fontsize=8.5, 
            color=(0, 0, 0), 
            rotate=0
        )
        
        # DL number with fixed x position at 515
        page.insert_text(
            (515, coords["text_y"]), 
            row_data["DLnum"], 
            fontsize=10, 
            color=(0, 0, 0), 
            rotate=0
        )

    return doc


class Ui_B(object):
    def __init__(self):
        # Initialize the attributes for PDF and image items
        self.template_pdf = None
        # self.signature_file_path = None
        self.output_pdf_path = None
        self.signature_file_path = None
        self.count = 0
        self.files = []

    def setupUi(self, B):
        if not B.objectName():
            B.setObjectName(u"B")
        B.resize(1103, 703)
        B.setMinimumSize(QSize(1103, 703))
        B.setMaximumSize(QSize(1103, 703))
        B.setAutoFillBackground(False)
        B.setStyleSheet(u"")
        self.centralwidget = QWidget(B)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.MainFrame = QFrame(self.centralwidget)
        self.MainFrame.setObjectName(u"MainFrame")
        self.MainFrame.setMinimumSize(QSize(1081, 681))
        self.MainFrame.setStyleSheet(u"QFrame {\n"
"    background-color: qlineargradient(\n"
"        spread: pad, \n"
"        x1: 0, y1: 0, \n"
"        x2: 1, y2: 0, /* Horizontal gradient for a smooth transition */\n"
"        stop: 0 #3A6073, /* Start color: Steel Blue */\n"
"        stop: 1 #16222A  /* End color: Dark Blue-Gray */\n"
"    );\n"
"    border-radius: 0px; /* Modern rounded corners */\n"
"    padding: 5px; /* Inner spacing for better balance */\n"
"}\n"
"")
        self.MainFrame.setFrameShape(QFrame.StyledPanel)
        self.MainFrame.setFrameShadow(QFrame.Raised)
        self.label = QLabel(self.MainFrame)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 40, 55, 61))
        self.label.setMaximumSize(QSize(55, 61))
        self.label.setStyleSheet(u"QLabel {\n"
"    background-color: transparent; /* White background */\n"
"    border-radius: 20px; /* Makes the QLabel circular */\n"
"    color: black; /* Text color */\n"
"    font-size: 14px; /* Adjust font size */\n"
"    text-align: center; /* Center text */\n"
"}\n"
"")
        self.label.setPixmap(QPixmap(u"check.png"))
        self.label.setScaledContents(True)
        self.label_2 = QLabel(self.MainFrame)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(10, 300, 55, 61))
        self.label_2.setMaximumSize(QSize(55, 61))
        self.label_2.setStyleSheet(u"QLabel {\n"
"    background-color: transparent; /* White background */\n"
"    border-radius: 20px; /* Makes the QLabel circular */\n"
"    color: black; /* Text color */\n"
"    font-size: 14px; /* Adjust font size */\n"
"    text-align: center; /* Center text */\n"
"}\n"
"")
        self.label_2.setPixmap(QPixmap(u"check.png"))
        self.label_2.setScaledContents(True)
        self.label_3 = QLabel(self.MainFrame)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(10, 575, 55, 61))
        self.label_3.setMaximumSize(QSize(55, 61))
        self.label_3.setStyleSheet(u"QLabel {\n"
"    background-color: transparent; /* White background */\n"
"    border-radius: 20px; /* Makes the QLabel circular */\n"
"    color: black; /* Text color */\n"
"    font-size: 14px; /* Adjust font size */\n"
"    text-align: center; /* Center text */\n"
"}\n"
"")
        self.label_3.setPixmap(QPixmap(u"check.png"))
        self.label_3.setScaledContents(True)
        self.line = QFrame(self.MainFrame)
        self.line.setObjectName(u"step1line")
        self.line.setGeometry(QRect(35, 90, 5, 218))
        self.line.setMaximumSize(QSize(5, 235))
        self.line.setStyleSheet(u"background-color: white;")
        self.line.setFrameShape(QFrame.Shape.VLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)
        self.line_2 = QFrame(self.MainFrame)
        self.line_2.setObjectName(u"step2line")
        self.line_2.setGeometry(QRect(35, 350, 5, 235))
        self.line_2.setMaximumSize(QSize(5, 235))
        self.line_2.setStyleSheet(u"background-color: white;")
        self.line_2.setFrameShape(QFrame.Shape.VLine)
        self.line_2.setFrameShadow(QFrame.Shadow.Sunken)
        self.widget = QWidget(self.MainFrame)
        self.widget.setObjectName(u"widget")
        self.widget.setGeometry(QRect(70, 40, 921, 601))
        self.verticalLayout_4 = QVBoxLayout(self.widget)
        self.verticalLayout_4.setSpacing(30)
        self.verticalLayout_4.setObjectName(u"verticalLayout_4")
        self.verticalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.AppName = QLabel(self.widget)
        self.AppName.setObjectName(u"AppName")
        self.AppName.setMinimumSize(QSize(721, 81))
        self.AppName.setMaximumSize(QSize(721, 81))
        font = QFont()
        font.setFamilies([u"Calibri"])
        font.setPointSize(26)
        font.setBold(True)
        self.AppName.setFont(font)
        self.AppName.setLayoutDirection(Qt.LeftToRight)
        self.AppName.setAutoFillBackground(False)
        self.AppName.setStyleSheet(u"color:white; background:transparent;")
        self.AppName.setTextFormat(Qt.PlainText)
        self.AppName.setAlignment(Qt.AlignCenter)
        self.AppName.setMargin(5)

        self.verticalLayout_2.addWidget(self.AppName)

        self.dragDropListWidget = QListWidget(self.widget)
        self.dragDropListWidget.setObjectName(u"dragDropListWidget")
        self.dragDropListWidget.setMaximumSize(QSize(721, 431))
        self.dragDropListWidget.setMouseTracking(False)
        self.dragDropListWidget.setAcceptDrops(True)
        self.dragDropListWidget.setDragEnabled(True)
        self.dragDropListWidget.setDragDropMode(QAbstractItemView.InternalMove)

        self.dragDropListWidget.setStyleSheet(""" 
        QListWidget {
            background-color: qlineargradient(
                spread: pad, 
                x1: 0, y1: 0, 
                x2: 1, y2: 0, /* Horizontal gradient for a smooth transition */
                stop: 0 #1f2e35, /* Darker Steel Blue */
                stop: 1 #0d1318  /* Almost Black for Dark Blue-Gray */
            );
            border-radius: 60px; /* Modern rounded corners */
            border: 1px solid #2e2e2e; /* Dark grey border for a subtler effect */
            padding: 15px; /* Inner spacing for better balance */
            background-image: url('C:/Users/tahir/Desktop/MYProject/file.png');
            background-repeat: no-repeat;
            background-position: center;
            font: 14px;
        }
        """)

        self.verticalLayout_2.addWidget(self.dragDropListWidget)

        # Define the dragEnterEvent and dropEvent inside the setupUi method
        def dragEnterEvent(event: QDragEnterEvent):
            """Handle drag event when data enters the widget."""
            if event.mimeData().hasUrls():  # Check if the dragged data has URLs (files)
                event.acceptProposedAction()  # Accept the drag event
            else:
                event.ignore()  # Ignore the event if it's not a file

        def dropEvent(event):
            """Handle the drop event."""
            mime_data = event.mimeData()
            
            if mime_data.hasUrls():  # Check if the dropped data contains URLs (files)
                for url in mime_data.urls():
                    file_path = url.toLocalFile()  # Get the file path
        
                    # Get the file extension
                    file_extension = os.path.splitext(file_path)[1].lower()
        
                    # Check if the file is a PDF
                    if file_extension == ".pdf":
                        # Add the PDF file to the list of files
                        self.files.append(file_path)

                        # Add each PDF file path to the list widget
                        pdf_item = QListWidgetItem(file_path)
                        self.dragDropListWidget.addItem(pdf_item)
        
                        # Update the count for PDFs and update the step line
                        self.count = len(self.files)  # Set the count for number of PDF files
                        self.updateStepLine(self.count)  # Update the step line
        
                    # Ignore other file types
                    else:
                        event.ignore()  # Ignore if it's not a valid PDF file
        
                event.acceptProposedAction()  # Accept the event after processing the files
            else:
                event.ignore()  # Ignore if the dropped data doesn't contain URLs

        # Set these functions to the dragDropListWidget
        self.dragDropListWidget.dragEnterEvent = dragEnterEvent
        self.dragDropListWidget.dropEvent = dropEvent
        self.horizontalLayout.addLayout(self.verticalLayout_2)

        self.verticalLayout_3 = QVBoxLayout()
        self.verticalLayout_3.setSpacing(0)
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.verticalLayout_3.setContentsMargins(-1, 200, -1, -1)
        self.Pdf = QPushButton(self.widget)
        self.Pdf.setObjectName(u"Pdf")
        self.Pdf.setMinimumSize(QSize(181, 60))
        self.Pdf.setMaximumSize(QSize(181, 60))
        font1 = QFont()
        font1.setFamilies([u"Segoe UI"])
        font1.setPointSize(10)
        font1.setBold(True)
        font1.setUnderline(False)
        font1.setStrikeOut(False)
        self.Pdf.setFont(font1)
        self.Pdf.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.Pdf.setStyleSheet(u"QPushButton {\n"
"    color: rgba(255, 255, 255, 0.8); /* Semi-transparent white text */\n"
"    background: blue; /* Original purple background */\n"
"    padding: 5px 5px; /* Adjusted padding to ensure button shape */\n"
"    border-radius: 30px; /* Fully rounded corners */\n"
"    transition: all 0.3s ease-in-out; /* Smooth hover effect */\n"
"    border: none; /* Clean, borderless design */\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    color: rgba(255, 255, 255, 1); /* Full white text on hover */\n"
"    box-shadow: 0 5px 15px rgba(145, 92, 182, 0.4); /* Neon purple glow effect */\n"
"    background:darkblue; /* Slightly darker purple on hover */\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background: rgb(100, 60, 130); /* Even darker purple when pressed */\n"
"    box-shadow: 0 2px 5px rgba(145, 92, 182, 0.4); /* Subtle shadow when pressed */\n"
"}\n"
"")

        self.verticalLayout_3.addWidget(self.Pdf)

        self.signature = QPushButton(self.widget)
        self.signature.setObjectName(u"signature")
        self.signature.setMinimumSize(QSize(181, 60))
        self.signature.setMaximumSize(QSize(181, 60))
        self.signature.setFont(font1)
        self.signature.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.signature.setStyleSheet(u"QPushButton {\n"
"    color: rgba(255, 255, 255, 0.8); /* Semi-transparent white text */\n"
"    background: blue; /* Original purple background */\n"
"    padding: 5px 5px; /* Adjusted padding to ensure button shape */\n"
"    border-radius: 30px; /* Fully rounded corners */\n"
"    transition: all 0.3s ease-in-out; /* Smooth hover effect */\n"
"    border: none; /* Clean, borderless design */\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    color: rgba(255, 255, 255, 1); /* Full white text on hover */\n"
"    box-shadow: 0 5px 15px rgba(145, 92, 182, 0.4); /* Neon purple glow effect */\n"
"    background: darkblue; /* Slightly darker purple on hover */\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background: rgb(100, 60, 130); /* Even darker purple when pressed */\n"
"    box-shadow: 0 2px 5px rgba(145, 92, 182, 0.4); /* Subtle shadow when pressed */\n"
"}\n"
"")

        self.verticalLayout_3.addWidget(self.signature)

        self.output_path = QPushButton(self.widget)
        self.output_path.setObjectName(u"output_path")
        self.output_path.setMinimumSize(QSize(181, 60))
        self.output_path.setMaximumSize(QSize(181, 60))
        self.output_path.setFont(font1)
        self.output_path.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.output_path.setStyleSheet(u"QPushButton {\n"
"    color: rgba(255, 255, 255, 0.8); /* Semi-transparent white text */\n"
"    background: blue; /* Original purple background */\n"
"    padding: 5px 5px; /* Adjusted padding to ensure button shape */\n"
"    border-radius: 30px; /* Fully rounded corners */\n"
"    transition: all 0.3s ease-in-out; /* Smooth hover effect */\n"
"    border: none; /* Clean, borderless design */\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    color: rgba(255, 255, 255, 1); /* Full white text on hover */\n"
"    box-shadow: 0 5px 15px rgba(145, 92, 182, 0.4); /* Neon purple glow effect */\n"
"    background: darkblue; /* Slightly darker purple on hover */\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background: rgb(100, 60, 130); /* Even darker purple when pressed */\n"
"    box-shadow: 0 2px 5px rgba(145, 92, 182, 0.4); /* Subtle shadow when pressed */\n"
"}\n"
"")

        self.verticalLayout_3.addWidget(self.output_path)


        self.horizontalLayout.addLayout(self.verticalLayout_3)


        self.verticalLayout_4.addLayout(self.horizontalLayout)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.progressBar = QProgressBar(self.widget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setMinimumSize(QSize(0, 0))
        self.progressBar.setMaximumSize(QSize(721, 60))
        font2 = QFont()
        self.progressBar.setFont(font2)
        self.progressBar.setStyleSheet(u"QProgressBar {\n"
"    border: 2px solid #3a3a3a; /* Dark gray border for the progress bar frame */\n"
"    border-radius: 10px; /* Rounded corners for the frame */\n"
"    background-color: #1e1e1e; /* Dark background for the progress bar */\n"
"    text-align: center; /* Text centered within the progress bar */\n"
"    color: white; /* White text for visibility */\n"
"    font-size: 14px; /* Adjust font size */\n"
"}\n"
"\n"
"QProgressBar::chunk {\n"
"     background: qlineargradient(\n"
"        x1: 0, y1: 0, x2: 0, y2: 1, /* Vertical gradient direction */\n"
"        stop: 0 #66ff66, /* Light green */\n"
"        stop: 1 #009933  /* Dark green */\n"
"    );\n"
"    border-radius: 8px; /* Rounded edges for the chunk */\n"
"    box-shadow: 0 0 10px rgba(0, 255, 0, 0.6); /* Green glow effect */\n"
"}\n"
"")
        self.progressBar.setValue(0)
        self.progressBar.setAlignment(Qt.AlignCenter)

        self.horizontalLayout_2.addWidget(self.progressBar)

        self.process = QPushButton(self.widget)
        self.process.setObjectName(u"process")
        self.process.setMinimumSize(QSize(181, 60))
        self.process.setMaximumSize(QSize(181, 60))
        self.process.setFont(font1)
        self.process.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.process.setStyleSheet(u"QPushButton {\n"
"    color: rgba(255, 255, 255, 0.8); /* Semi-transparent white text */\n"
"    background: green; /* Original purple background */\n"
"    padding: 5px 5px; /* Adjusted padding to ensure button shape */\n"
"    border-radius: 30px; /* Fully rounded corners */\n"
"    transition: all 0.3s ease-in-out; /* Smooth hover effect */\n"
"    border: none; /* Clean, borderless design */\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    color: rgba(255, 255, 255, 1); /* Full white text on hover */\n"
"    box-shadow: 0 5px 15px rgba(145, 92, 182, 0.4); /* Neon purple glow effect */\n"
"    background: darkgreen; /* Slightly darker purple on hover */\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background: rgb(100, 60, 130); /* Even darker purple when pressed */\n"
"    box-shadow: 0 2px 5px rgba(145, 92, 182, 0.4); /* Subtle shadow when pressed */\n"
"}\n"
"")

        self.horizontalLayout_2.addWidget(self.process)


        self.verticalLayout_4.addLayout(self.horizontalLayout_2)


        self.verticalLayout.addWidget(self.MainFrame)

        B.setCentralWidget(self.centralwidget)

        self.retranslateUi(B)

        QMetaObject.connectSlotsByName(B)
    # setupUi

    def retranslateUi(self, B):
        B.setWindowTitle(QCoreApplication.translate("B", u"MainWindow", None))
        self.label.setText("")
        self.label_2.setText("")
        self.label_3.setText("")
        self.AppName.setText(QCoreApplication.translate("B", u"BTW Report Maker", None))
        self.Pdf.setText(QCoreApplication.translate("B", u"Upload Template Pdf", None))
        self.signature.setText(QCoreApplication.translate("B", u"Upload Signature", None))
        self.output_path.setText(QCoreApplication.translate("B", u"Output Path", None))
        self.progressBar.setFormat(QCoreApplication.translate("B", u"%p%", None))
        self.process.setText(QCoreApplication.translate("B", u"Process", None))
    # retranslateUi

        self.Pdf.clicked.connect(self.select_template)
        self.signature.clicked.connect(self.select_signature)
        self.output_path.clicked.connect(self.select_output_dir)
        self.process.clicked.connect(self.process_files)
        

    def updateStepLine(self, count):
        if count == 2:
                self.line.setStyleSheet(u"background-color: rgb(32, 216, 11);")
        elif count == 3:
                self.line_2.setStyleSheet(u"background-color: rgb(32, 216, 11);")

    def select_template(self):
            template_path, _ = QFileDialog.getOpenFileName(self.dragDropListWidget, "Select Template PDF", "", "PDF Files (*.pdf)")
            if template_path:
                self.template_pdf = template_path
                self.count = 2
                self.updateStepLine(self.count)
                if hasattr(self, 'template_pdf') and self.template_pdf:  
                    self.dragDropListWidget.takeItem(self.dragDropListWidget.row(self.template_pdf))
                self.template_pdf = QListWidgetItem(self.template_pdf)  
                self.dragDropListWidget.addItem(self.template_pdf)
                self.has_pdf = True 
                print(template_path)
                setup_paths(self.template_pdf,self.signature_file_path,self.output_pdf_path)



    def select_signature(self):
            signature_path, _ = QFileDialog.getOpenFileName(self.dragDropListWidget, "Select Signature Image", "", "Image Files (*.png *.jpg *.jpeg)")
            if signature_path:
                self.signature_file_path = signature_path
                self.count = 3
                self.updateStepLine(self.count)
                if hasattr(self, 'signature_file_path') and self.signature_file_path:  
                    self.dragDropListWidget.takeItem(self.dragDropListWidget.row(self.signature_file_path))
                self.signature_file_path = QListWidgetItem(self.signature_file_path)
                self.dragDropListWidget.addItem(self.signature_file_path)
                self.has_image = True  # Mark that an image is in the list
                setup_paths(self.template_pdf,self.signature_file_path,self.output_pdf_path)


    
    def select_output_dir(self):
        # Use self.parent() to get the parent window
        output_dir = QFileDialog.getExistingDirectory(self.dragDropListWidget, "Select Output Directory")
        if output_dir:
            self.output_pdf_path = output_dir 
            self.count += 1
            self.updateStepLine(self.count)
            setup_paths(self.template_pdf,self.signature_file_path,self.output_pdf_path)

            print(self.output_pdf_path)

    
    def process_files(self):
        """Main file processing loop"""
        print("\nChecking for config file...")
        # config = load_config()

        # First time setup if no config exists
        # if not config:
        #     print("No config found! Starting first-time setup...")
        # else:
        #     print("Found existing config with paths:")
        #     print(f"Template: {config['template_path']}")
        #     print(f"Signature: {config['signature_path']}")
        #     print(f"Output: {config['output_dir']}")

        config = setup_paths(self.template_pdf,self.signature_file_path,self.output_pdf_path)

        

        if not self.files:
            print("No files detected. Please provide valid files.")
            return
        
        self.progressBar.setValue(0)
        success = 0
        total = len(self.files)
        summaries = []

        print(f"\nProcessing {total} file(s)...")
        for file in self.files:
            # Clean up the file path
            self.progressBar.setValue(100)
            file = os.path.expanduser(file.strip().strip('"'))
            print(f"\nChecking file: {file}")
            self.progressBar.setValue(50)

            try:
                if not os.path.exists(file):
                    print(f"✗ File not found: {file}")
                    continue

                # Case-insensitive check for the file ending
                if not file.upper().endswith("BTWATTENDANCEHISTORYREPORT.PDF"):
                    print(f"✗ Not a valid attendance report: {file}")
                    continue

                self.progressBar.setValue(100)
                print(f"Processing with:")
                print(f"- Template: {config['template_path']}")
                print(f"- Signature: {config['signature_path']}")
                print(f"- Output dir: {config['output_dir']}")

                process_result, summary = process_single_file(
                    file,
                    config["template_path"],
                    config["signature_path"],
                    config["output_dir"]
                )

                if process_result:
                    success += 1
                    if summary:
                        summaries.append(summary)
            except Exception as e:
                print(f"✗ Error processing {file}: {str(e)}")
                import traceback
                print(traceback.format_exc())

        print(f"\nComplete! Successfully processed {success}/{total} files.")
        self.files.clear()
        self.dragDropListWidget.clear()
        self.progressBar.setValue(0)
        if summaries:
            print("\nProcessing Summaries:")
            for summary in summaries:
                print(summary)
        print(f"\nOutput location: {config['output_dir']}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_B()
        self.ui.setupUi(self)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()