import pandas as pd
from sqlalchemy import create_engine, Column, String, ForeignKey, Integer
from sqlalchemy.orm import sessionmaker, relationship, declarative_base


# Load data
staff_df = pd.read_excel("C:/Users/LG/PycharmProjects/PCCC/Test-DSCBCNV.xlsx")
staff_df = staff_df.rename(columns={
    'Họ & tên': 'HoTen',
    'MSNV': 'MaNV',
    'Bộ phận': 'BoPhan',
    'Phòng/Ban': 'PhongBan',
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


# Insert staff data into NhanVien table
for idx in range(staff_df.__len__()):
    staff_info = staff_df.iloc[idx]
    manv = staff_info.MaNV
    hoten = staff_info.HoTen
    bophan = staff_info.BoPhan
    phongban = staff_info.PhongBan
    tinhtrang = staff_info.TinhTrang
    if tinhtrang == "Có mặt":
        nhanvien = NhanVien(MaNV=manv, HoTen=hoten, BoPhan=bophan, PhongBan=phongban)
        session.add(nhanvien)
        session.commit()

# Insert accounts
accounts = [
    {"MaNV": "H3803", "HoTen": "Vũ Thanh Hoà", "BoPhan": "THOP", "PhongBan": "Phòng Tổng hợp", "PassWord": "123", "Role": "admin"},
    {"MaNV": "H0844", "HoTen": "Nguyễn Thị Diệu Hương", "BoPhan": "XN", "PhongBan": "Xí nghiệp", "PassWord": "123", "Role": "user"},
    {"MaNV": "L1349", "HoTen": "Nguyễn Thị Na Ly", "BoPhan": "XN1", "PhongBan": "Xí nghiệp", "PassWord": "123", "Role": "user"},
]

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
print("Done!!!")