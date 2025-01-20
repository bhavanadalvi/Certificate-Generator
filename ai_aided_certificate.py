import xlrd
import os
import smtplib
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import csv
from datetime import datetime

# Load the Excel file
file_path = './data_19Jan_2025.xlsx'
data = pd.read_excel(file_path)

out_file_path = './out_data_19Jan_2025.csv'
csvfile= open(out_file_path, 'w', newline='')
csvwriter = csv.writer(csvfile, delimiter='\t',
                            quotechar='|', quoting=csv.QUOTE_MINIMAL)
    

# Define certificate template and font
certificate_template = './certificate_template_5.png'  # Ensure you have this template
font_path = './ChopinScript.otf'  # Modify if needed
font_size = 60
output_folder = './certificates'
# Create certificates
def create_certificate(student_name, course_name, output_path):
    image = Image.open(certificate_template)
    draw = ImageDraw.Draw(image)
    font_name = ImageFont.truetype(font_path, 120)
    font_course = ImageFont.truetype(font_path, 70)
    sign_font = ImageFont.truetype(font_path, 50)

    # Define positions
    name_position = (350, 700)
    course_position = (350, 1000)
    date_position = (300, 1200)
    sign_position = (1100, 1200)

    current_timestamp = datetime.now().timestamp()
    date_time = datetime.fromtimestamp(current_timestamp)
    date_str = date_time.strftime("%d %B, %Y")
    
    draw.text(name_position, student_name, font=font_name, fill='black')
    draw.text(course_position, course_name, font=font_course, fill='black')
    draw.text(date_position, date_str, font=sign_font, fill='black')
    draw.text(sign_position, str(int(current_timestamp)), font=sign_font, fill='black')

    image.save(output_path)
    print(f'Certificate created for {student_name}')

csvwriter.writerow(['Timestamp', 'student_name', 'email', 'course_name', 'screenshot', 'output_path'])

# Generate certificates for each row
for index, row in data.iterrows():
    student_name = row['student_name']
    course_name = row['course_name']
    email = row['email']
    

    output_path = f'{output_folder}/{student_name.replace(" ", "_")}.{course_name.replace(" ", "_")}.certificate.png'
    create_certificate(student_name, course_name, output_path)
    csvwriter.writerow([row['Timestamp'], row['student_name'], row['email'], row['course_name'], row['screenshot'], output_path])

csvfile.close()
print("Certificates generation completed!")
