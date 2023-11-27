from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
import datetime
import os

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials


def create_excel_file(results, new_history, filename):
    # Convert new_history to DataFrame
    new_history_df = pd.DataFrame(new_history, index=[0])
    new_history_df.columns = ["Mốc thời gian", "Kết quả", "Lý do", "Thời gian diễn ra thực tế",
                              "Tổng số có mặt trong ngày", "Có mặt tại nơi tập trung", "Vắng mặt"]


    # Create a DataFrame for detailed results
    detailed_results = []
    for department, result in results.items():
        detailed_results.append({
            'Mốc thời gian': result['last_time'],
            'Xí nghiệp/Phòng ban': department,
            'Tổ/Bộ phận': department,
            'Tổng số có mặt trong ngày': result['total'],
            'Có mặt tại nơi tập trung': result['num_done'],
            'Vắng mặt': result['total'] - result['num_done'],
        })
    detailed_results_df = pd.DataFrame(detailed_results)


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
        detailed_results_df.to_excel(writer, sheet_name='Chi tiết', index=False)
        absent_staff_df.to_excel(writer, sheet_name='Danh sách vắng', index=False)
    print(f"File {filename} has been created.")


def upload_file_to_drive(filename, folder_id):
    creds = Credentials.from_authorized_user_file('static/credentials.json.json')
    service = build('drive', 'v3', credentials=creds)

    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.ms-excel')
    file = service.files().create(body=file_metadata, media_body=media, fields='id,webViewLink').execute()

    permission = {
        'type': 'anyone',
        'role': 'reader',
    }
    try:
        service.permissions().create(fileId=file['id'], body=permission).execute()
    except HttpError as error:
        print(f'An error occurred: {error}')

    return file['webViewLink']

def send_email_with_link(recipient, link, new_history):
    # Set up the SMTP server
    s = smtplib.SMTP_SSL(host='mail.hachiba.com.vn', port=465)
    s.login('vuthanh.hoa@hachiba.com.vn', 'Ho*!!811th')

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = 'vuthanh.hoa@hachiba.com.vn'
    msg['To'] = recipient

    # Format the date
    date = datetime.datetime.strptime(new_history['MocThoiGian'], "%d/%m/%Y %H:%M").date()
    msg['Subject'] = Header(f"Báo cáo kết quả diễn tập PCCC ngày {date.strftime('%d/%m/%Y')}", 'utf-8')

    # Add the link to the email body
    msg.attach(MIMEText(f'<a href="{link}">Click here to view the file</a>', 'html'))

    # Send the email
    s.send_message(msg)
    s.quit()

