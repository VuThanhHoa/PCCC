from flask_login import login_required
from flask import render_template, request, redirect, url_for, session, Blueprint

import os
import pickle
import datetime

from src.pccc.dashboard import get_results
from src.pccc.database import session, engine, DienTap, KetQua
from src.pccc.post_training import send_email_with_attachment, create_excel_file

# Init
admin = Blueprint("admin", __name__)
TRAINING_TIME_DIR="./static/last_training.pkl"

@admin.route('/dashboard', methods=["GET", "POST"])
@login_required
def dashboard():
    if request.method == "POST":
        
        if request.form.get("button-start", False) == "start":
            last_training_time = datetime.datetime.now()
            pickle.dump(last_training_time, open(TRAINING_TIME_DIR, "wb"))
            session.query(DienTap).delete()
            session.query(KetQua).delete()
            session.commit()
            
        elif request.form.get("button-info", False) == "info":
            return redirect(url_for("admin.training_details"))
        
        elif request.form.get("button-logout", False) == "logout":
            return redirect(url_for("login"))
        
    if os.path.isfile(TRAINING_TIME_DIR):
        last_training_time = pickle.load(open(TRAINING_TIME_DIR, "rb"))
        info = f"Lần diễn tập gần nhất: {last_training_time}"
        
    else:    
        info = "Chưa có buổi diễn tập nào được diễn ra"
        
    return render_template("adminPCCC.html", info=info)

@admin.route('/dashboard/training-details', methods=["GET", "POST"])
@login_required
def training_details():
    
    results, new_history = get_results(engine=engine, training_time_dir=TRAINING_TIME_DIR)
    all_done = all(result['is_done'] for result in results.values())

    if all_done:
        # Format the date
        date = datetime.datetime.strptime(new_history['ThoiDiem'], "%d/%m/%Y %H:%M").date()

        # Create the Excel file
        filename = f'Báo cáo diễn tập PCCC ngày {date.strftime("%d/%m/%Y")}.xlsx'
        create_excel_file(results, new_history, filename)

        # Send the email
        recipient = 'vuthanh.hoa@hachiba.com.vn'
        send_email_with_attachment(recipient, filename, new_history)

    return render_template("training-details.html", results=results, new_history=new_history)