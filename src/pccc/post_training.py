from email import encoders
from email.mime.base import MIMEBase
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
import datetime
import os


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


def send_email_with_attachment(recipient, filename, new_history):
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

    # Add the attachment
    part = MIMEBase('application', 'octet-stream')
    with open(filename, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
    msg.attach(part)

    # Send the email
    s.send_message(msg)
    s.quit()