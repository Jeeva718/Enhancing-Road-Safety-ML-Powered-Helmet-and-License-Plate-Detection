import cv2
import numpy as np
from ultralytics import YOLO
import cvzone
from paddleocr import PaddleOCR
import os
from datetime import datetime
import xlwings as xw
import pandas as pd
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# Initialize PaddleOCR
ocr = PaddleOCR()

# Gmail credentials for email notifications
GMAIL_ADDRESS = 'tnvahaninfo@gmail.com'
GMAIL_APP_PASSWORD = 'uclynnkicbpfmozd'  # App Password with spaces removed

# Base directory for project
BASE_DIR = 'C:/Users/Admin/Desktop/Project1/no-helmet-numberplate-main'

# Function to sanitize filenames
def sanitize_filename(filename):
    invalid_chars = r'[<>:"/\\|?*\x00-\x1F]'
    sanitized = re.sub(invalid_chars, '_', filename)
    return sanitized.strip().strip('.')

# Function to validate email
def is_valid_email(email):
    if email == "Not Found" or email == "DB Error" or not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

# Function to send email with image attachment and bike details using Gmail SMTP
def send_email(email, number_plate, date, time, image_path, name, bike_brand, bike_model, registration):
    if not is_valid_email(email):
        print("Email not sent: Invalid or missing email address")
        return "Failed: Invalid email"
    try:
        # Create multipart email
        msg = MIMEMultipart()
        msg['Subject'] = 'No Helmet Detection Alert'
        msg['From'] = GMAIL_ADDRESS
        msg['To'] = email

        # Create HTML email body for professional formatting
        html_body = f"""
        <html>
            <body style="font-family: Arial, sans-serif; color: #333;">
                <h2 style="color: #d32f2f;">Traffic Safety Division</h2>
                <p>Dear {name},</p>
                <p>A violation has been recorded for failing to wear a helmet while riding. Below are the details:</p>
                <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 10px 0;">
                    <tr style="background-color: #f5f5f5;">
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Number Plate</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{number_plate}</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Date</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{date}</td>
                    </tr>
                    <tr style="background-color: #f5f5f5;">
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Time</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{time}</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Bike Brand</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{bike_brand}</td>
                    </tr>
                    <tr style="background-color: #f5f5f5;">
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Bike Model</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{bike_model}</td>
                    </tr>
                    <tr>
                        <td style="padding: 8px; border: 1px solid #ddd;"><strong>Registration</strong></td>
                        <td style="padding: 8px; border: 1px solid #ddd;">{registration}</td>
                    </tr>
                </table>
                <p>Please wear a helmet for your safety and comply with traffic regulations. The license plate image is attached for reference.</p>
                <p style="margin-top: 20px;">Sincerely,<br>Traffic Safety Division<br>Email: tnvahaninfo@gmail.com<br>Phone: +91 6381136346</p>
            </body>
        </html>
        """
        text = MIMEText(html_body, 'html')
        msg.attach(text)

        # Attach image if it exists
        if os.path.exists(image_path):
            with open(image_path, 'rb') as img_file:
                img = MIMEImage(img_file.read(), name=os.path.basename(image_path))
                msg.attach(img)
            print(f"Attached image {image_path} to email")
        else:
            print(f"Image not found for attachment: {image_path}")
            return f"Failed: Image not found at {image_path}"

        # Send email
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            server.send_message(msg)
        print(f"Email sent successfully to {email} with attachment")
        return "Sent"
    except Exception as e:
        print(f"Failed to send email to {email}: {str(e)}")
        if 'authentication' in str(e).lower():
            print("Error: Check GMAIL_ADDRESS and GMAIL_APP_PASSWORD.")
        return f"Failed: {str(e)}"

# Function to perform OCR on an image array
def perform_ocr(image_array):
    if image_array is None or image_array.size == 0:
        raise ValueError("Image is None or empty")
    results = ocr.ocr(image_array, rec=True)
    detected_text = []
    if results[0] is not None:
        for result in results[0]:
            text = result[1][0]
            detected_text.append(text)
    return ''.join(detected_text) or "OCR_Failed"

# Mouse callback function for RGB window
def RGB(event, x, y, flags, param):
    if event == cv2.EVENT_MOUSEMOVE:
        point = [x, y]
        print(point)

# Load CSV database
def load_database(db_path=os.path.join(BASE_DIR, 'license_plate_database.csv')):
    try:
        if not os.path.exists(db_path):
            print(f"Error: Database file not found at {db_path}")
            return None
        df = pd.read_csv(db_path)
        required_columns = ['Number Plate', 'Name', 'Phone Number', 'Email ID', 'Bike Brand', 'Bike Model', 'Registration']
        if not all(col in df.columns for col in required_columns):
            print(f"Error: Database missing required columns. Found: {list(df.columns)}, Expected: {required_columns}")
            return None
        print(f"Database loaded successfully from {db_path}")
        return df
    except Exception as e:
        print(f"Error loading database from {db_path}: {str(e)}")
        return None

# Function to normalize number plate
def normalize_number_plate(number_plate):
    cleaned = re.sub(r'[^A-Za-z0-9.]', '', number_plate)
    return cleaned.strip().upper()

# Function to fetch details from database
def get_details_from_db(number_plate, df):
    if df is None:
        return "DB Error", "DB Error", "DB Error", "DB Error", "DB Error", "DB Error"
    number_plate = normalize_number_plate(number_plate)
    print(f"Normalized Number Plate for lookup: {number_plate}")
    match = df[df['Number Plate'].apply(normalize_number_plate) == number_plate]
    if not match.empty:
        return (
            match.iloc[0]['Name'],
            match.iloc[0]['Phone Number'],
            match.iloc[0]['Email ID'],
            match.iloc[0]['Bike Brand'],
            match.iloc[0]['Bike Model'],
            match.iloc[0]['Registration']
        )
    else:
        return "Not Found", "Not Found", "Not Found", "Not Found", "Not Found", "Not Found"

# Function to get or create Excel workbook
def get_excel_workbook(current_date, base_dir=BASE_DIR):
    try:
        date_dir = os.path.join(base_dir, current_date)
        os.makedirs(date_dir, exist_ok=True)
        excel_file_path = os.path.join(date_dir, f"{current_date}.xlsx")
        if os.path.exists(excel_file_path):
            wb = xw.Book(excel_file_path)
        else:
            wb = xw.Book()
            wb.sheets[0].range("A1").value = [
                "Number Plate", "Name", "Phone Number", "Email ID",
                "Bike Brand", "Bike Model", "Registration",
                "Date", "Time", "Email Status"
            ]
        ws = wb.sheets[0]
        print(f"Excel workbook opened/created: {excel_file_path}")
        return wb, ws, excel_file_path
    except Exception as e:
        print(f"Error creating Excel workbook: {str(e)}")
        raise

# Check write permissions
def check_write_permissions(directory):
    try:
        test_file = os.path.join(directory, 'test_write.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        return True
    except Exception as e:
        print(f"Error: No write permissions for {directory}: {str(e)}")
        return False

cv2.namedWindow('RGB')
cv2.setMouseCallback('RGB', RGB)

# Verify write permissions
if not check_write_permissions(BASE_DIR):
    raise PermissionError(f"Cannot write to {BASE_DIR}. Run script with administrator privileges or change directory.")

# Load YOLOv11 model
model = YOLO("best.pt")
names = model.names

# Define polygon area
area = [(1, 173), (62, 468), (608, 431), (364, 155)]

# Load database
db_df = load_database()

# Initialize variables for Excel handling
current_date = datetime.now().strftime('%Y-%m-%d')
wb, ws, excel_file_path = get_excel_workbook(current_date)

# Track processed track IDs
processed_track_ids = set()

# Open the video file
cap = cv2.VideoCapture('night.mp4')

while True:
    new_date = datetime.now().strftime('%Y-%m-%d')
    if new_date != current_date:
        try:
            wb.save(excel_file_path)
            wb.close()
            print(f"Closed Excel workbook: {excel_file_path}")
        except Exception as e:
            print(f"Error closing Excel workbook: {str(e)}")
        current_date = new_date
        wb, ws, excel_file_path = get_excel_workbook(current_date)
        processed_track_ids.clear()
        print(f"New Excel workbook opened for {current_date}: {excel_file_path}")

    ret, frame = cap.read()
    if not ret:
        break
    
    frame = cv2.resize(frame, (1020, 500))
    
    results = model.track(frame, persist=True)
    
    no_helmet_detected = False
    numberplate_box = None
    numberplate_track_id = None
    
    if results[0].boxes is not None and results[0].boxes.id is not None:
        boxes = results[0].boxes.xyxy.int().cpu().tolist()
        class_ids = results[0].boxes.cls.int().cpu().tolist()
        track_ids = results[0].boxes.id.int().cpu().tolist()
        confidences = results[0].boxes.conf.cpu().tolist()
        
        for box, class_id, track_id, conf in zip(boxes, class_ids, track_ids, confidences):
            c = names[class_id]
            x1, y1, x2, y2 = box
            cx = (x1 + x2) // 2
            cy = (y1 + y2) // 2
            
            result = cv2.pointPolygonTest(np.array(area, np.int32), (cx, cy), False)
            if result >= 0:
                if c == 'no-helmet':
                    no_helmet_detected = True
                elif c == 'numberplate':
                    numberplate_box = box
                    numberplate_track_id = track_id
        
        if no_helmet_detected and numberplate_box is not None and numberplate_track_id not in processed_track_ids:
            x1, y1, x2, y2 = numberplate_box
            if x1 < x2 and y1 < y2 and x1 >= 0 and y1 >= 0 and x2 <= frame.shape[1] and y2 <= frame.shape[0]:
                crop = frame[y1:y2, x1:x2]
                crop = cv2.resize(crop, (120, 85))
                cvzone.putTextRect(frame, f'{track_id}', (x1, y1), 1, 1)
                
                text = perform_ocr(crop)
                print(f"Detected Number Plate: {text}")
                
                name, phone, email, bike_brand, bike_model, registration = get_details_from_db(text, db_df)
                
                time_for_excel = datetime.now().strftime('%H:%M:%S')  # For Excel and email
                time_for_filename = datetime.now().strftime('%H-%M-%S')  # For filename
                safe_text = sanitize_filename(text)
                safe_time = sanitize_filename(time_for_filename)
                crop_image_path = os.path.join(BASE_DIR, current_date, f"{safe_text}_{safe_time}.jpg")
                
                try:
                    if crop.size > 0:
                        success = cv2.imwrite(crop_image_path, crop)
                        if success:
                            print(f"Image saved successfully to {crop_image_path}")
                        else:
                            print(f"Failed to save image to {crop_image_path}: cv2.imwrite returned False")
                    else:
                        print(f"Failed to save image to {crop_image_path}: Crop image is empty")
                except Exception as e:
                    print(f"Failed to save image to {crop_image_path}: {str(e)}")
                
                email_status = send_email(email, text, current_date, time_for_excel, crop_image_path, name, bike_brand, bike_model, registration)
                
                try:
                    last_row = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row
                    ws.range(f"A{last_row+1}").value = [
                        text, name, phone, email,
                        bike_brand, bike_model, registration,
                        current_date, time_for_excel, email_status
                    ]
                    wb.save()  # Save after each write
                    print(f"Data appended to Excel: {excel_file_path}")
                except Exception as e:
                    print(f"Error writing to Excel: {str(e)}")
                
                processed_track_ids.add(numberplate_track_id)
            else:
                print(f"Invalid crop coordinates: x1={x1}, y1={y1}, x2={x2}, y2={y2}")
    
    cv2.polylines(frame, [np.array(area, np.int32)], True, (255, 0, 255), 2)
    
    cv2.imshow("RGB", frame)
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

cap.release()
cv2.destroyAllWindows()

try:
    wb.save(excel_file_path)
    wb.close()
    print(f"Final save and close of Excel workbook: {excel_file_path}")
except Exception as e:
    print(f"Error during final Excel save/close: {str(e)}")
