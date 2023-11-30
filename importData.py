import pandas as pd
from sqlalchemy import create_engine, Column, String, Integer
from sqlalchemy.orm import sessionmaker, declarative_base
import gspread
from gspread_dataframe import get_as_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# Authorize and connect to Google Sheets
# Use the json key file you got from Google Cloud
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('baocaolaodong-2c9c27aea864.json', scope)

gc = gspread.authorize(credentials)

# Open the Google Spreadsheet using its name (make sure you have shared it with your service account)
spreadsheet = gc.open('BÁO CÁO LAO ĐỘNG HÀNG NGÀY (Từ 1.6.2023)')

# Select the sheet within the Spreadsheet
worksheet = spreadsheet.worksheet('0. LÝ LỊCH LĐ (01-10-2023)')

# Get the data into a pandas DataFrame
values = worksheet.get_all_values()
staff_df = pd.DataFrame(values[1:], columns=values[0])

# Rename the columns
staff_df = staff_df.rename(columns={
    'HỌ & TÊN': 'HoTen',
    'Mã số nhân viên': 'MaNV',
    'Tên bộ phận\n(01/10/2023)': 'BoPhan',
    'Đơn vị\n(20/03/2023)': 'PhongBan',
    'Tình trạng': 'TinhTrang'
})

# Create engine and session
engine = create_engine('sqlite:///C:/Users/LG/PycharmProjects/PCCC/instance/pccc.db')
Base = declarative_base()
Session = sessionmaker(bind=engine)
session = Session()


# Define classes
class NhanVien(Base):
    Id = Column(Integer, primary_key=True, autoincrement=True)
    MaNV = Column(String(10))
    HoTen = Column(String(255))
    BoPhan = Column(String(255))
    PhongBan = Column(String(255))
    __tablename__  = "NhanVien"
    extend_existing = True

class Account(Base):
    Id = Column(Integer, primary_key=True, autoincrement=True)
    MaNV = Column(String(10))
    HoTen = Column(String(255))
    BoPhan = Column(String(255))
    PhongBan = Column(String(255))
    PassWord = Column(String(255))
    Role = Column(String(10))
    __tablename__  = "Account"
    extend_existing = True


# Create tables
Base.metadata.create_all(engine)


# Delete existing data in NhanVien and Account tables
session.query(NhanVien).delete()
session.query(Account).delete()
session.commit()


# Fileter staff data
filtered_staff_df = staff_df.loc[((staff_df['TinhTrang'] == 'Có mặt') | (staff_df['TinhTrang'] == 'Mới')) &
                                 (staff_df['BoPhan'].notna()) & (staff_df['BoPhan'].str.strip() != '') & (staff_df['BoPhan'] != '#N/A')]

# Insert staff data into NhanVien table
for idx in range(len(filtered_staff_df)):
    staff_info = filtered_staff_df.iloc[idx]
    manv = staff_info.MaNV
    hoten = staff_info.HoTen
    bophan = staff_info.BoPhan
    phongban = staff_info.PhongBan
    nhanvien = NhanVien(MaNV=manv, HoTen=hoten, BoPhan=bophan, PhongBan=phongban)
    session.add(nhanvien)
    session.commit()

# Open the Google Spreadsheet containing MaNV of employees to create accounts for
spreadsheet_accounts = gc.open('PHÂN CÔNG - Các cá nhân nhận nhiệm vụ khi có CHÁY  DIỄN TẬP tại công ty (9-2023)')
worksheet_accounts = spreadsheet_accounts.worksheet('DS người phụ trách điểm danh')

# Get MaNV, Role and Password into a list
values = worksheet_accounts.get_all_values()
manv_role_password_list = [(item[3].lower(), item[5], item[6]) for item in values[1:]]  # assuming Password is in the 6th column

# Filter staff_df to only include rows with MaNV in manv_list
manv_list = [item[0] for item in manv_role_password_list]
filtered_staff_df = staff_df[staff_df['MaNV'].str.lower().isin(manv_list)]  # convert MaNV to lowercase before comparing

# Convert filtered_staff_df to a list of dictionaries
accounts = filtered_staff_df.to_dict('records')

# Add Role and Password to each account
for acc in accounts:
    # Find the role and password for this MaNV
    for manv, role, password in manv_role_password_list:
        if acc["MaNV"].lower() == manv:  # convert MaNV to lowercase before comparing
            acc["PassWord"] = password
            acc["Role"] = role
            break

# Insert accounts
for acc in accounts:
    MaNV = acc["MaNV"]
    HoTen = acc["HoTen"]
    BoPhan = acc["BoPhan"]
    PhongBan = acc["PhongBan"]
    PassWord = acc["PassWord"]
    Role = acc["Role"]

    account = Account(MaNV=MaNV, HoTen=HoTen, BoPhan=BoPhan, PhongBan=PhongBan, PassWord=PassWord, Role=Role)
    session.add(account)

session.commit()

# # Insert accounts
# accounts = [
#     {"MaNV": "H3803", "HoTen": "Vũ Thanh Hoà", "BoPhan": "THOP", "PhongBan": "Phòng Tổng hợp", "PassWord": "123", "Role": "admin"},
#     {"MaNV": "H0844", "HoTen": "Nguyễn Thị Diệu Hương", "BoPhan": "XN", "PhongBan": "Xí nghiệp", "PassWord": "123", "Role": "user"},
#     {"MaNV": "L1349", "HoTen": "Nguyễn Thị Na Ly", "BoPhan": "XN1", "PhongBan": "Xí nghiệp", "PassWord": "123", "Role": "user"},
# ]
#
# for acc in accounts:
#     MaNV = acc["MaNV"]
#     HoTen = acc["HoTen"]
#     BoPhan = acc["BoPhan"]
#     PhongBan = acc["PhongBan"]
#     PassWord = acc["PassWord"]
#     Role = acc["Role"]
#
#     account = Account(MaNV=MaNV, HoTen=HoTen, BoPhan=BoPhan, PhongBan=PhongBan, PassWord=PassWord, Role=Role)
#     session.add(account)
#     session.commit()
print("Done!!!")