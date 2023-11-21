from email.mime.text import MIMEText
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
import datetime
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


def create_excel_file(results, new_history, filename):
    # Convert new_history to DataFrame
    new_history_df = pd.DataFrame(new_history, index=[0])
    new_history_df.columns = ["Thời điểm", "Thời gian hoàn thành", "Kết quả", "Tỷ lệ vắng mặt", "Tỷ lệ có mặt"]

    # Create a DataFrame for absent staff
    absent_staff = []
    for department, result in results.items():
        for staff in result['absents'].split(' ,'):
            absent_staff.append({
                'Bộ phận': department,
                'Nhân viên vắng': staff
            })
    absent_staff_df = pd.DataFrame(absent_staff)

    # Get the directory name
    dir_name = os.path.dirname(filename)

    # Create the directory if it does not exist
    os.makedirs(dir_name, exist_ok=True)

    # Write to Excel file
    with pd.ExcelWriter(filename) as writer:
        new_history_df.to_excel(writer, sheet_name='Kết quả diễn tập', index=False)
        absent_staff_df.to_excel(writer, sheet_name='Danh sách vắng', index=False)


def send_email_with_attachment(recipient, filename, new_history):
    # Set up the Google Drive client
    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)

    # Upload the file to Google Drive
    gfile = drive.CreateFile({'title': filename})
    gfile.SetContentFile(filename)
    gfile.Upload()

    # Get the shareable link
    link = gfile['alternateLink']

    # Set up the SMTP server
    s = smtplib.SMTP_SSL(host='mail.hachiba.com.vn', port=465)
    s.login('vuthanh.hoa@hachiba.com.vn', 'Ho*!!811th')

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = 'vuthanh.hoa@hachiba.com.vn'
    # msg['To'] = 'hophat@hachiba.com.vn'
    msg['To'] = 'vuthanh.hoa@hachiba.com.vn'

    # Format the date
    date = datetime.datetime.strptime(new_history['ThoiDiem'], "%d/%m/%Y %H:%M").date()
    msg['Subject'] = Header(f"Báo cáo kết quả diễn tập PCCC ngày {date.strftime('%d/%m/%Y')}", 'utf-8')

    # Add the link to the email body
    msg.attach(MIMEText(f'Truy cập link này để xem báo cáo : {link}'))

    # Send the email
    s.send_message(msg)
    s.quit()