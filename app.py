# app.py
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import sys
import io
from datetime import date

from scheduler.tier_1 import generate_tier1_schedule_file

app = Flask(__name__)
app.secret_key = 'hieuvd99'
UPLOAD_FOLDER = 'LichTruc_Exports'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    download_link = None
    current_month = date.today().month
    current_year = date.today().year

    if request.method == 'POST':
        try:
            month = int(request.form['month'])
            year = int(request.form['year'])
            holidays_str = request.form.get('holidays', '')
            selected_tier = request.form.get('tier_selection', 'tier1')

            if selected_tier == 'tier1':
                success, msg, filename = generate_tier1_schedule_file(
                    desired_month=month,
                    desired_year=year,
                    holiday_input_str=holidays_str,
                    output_dir=UPLOAD_FOLDER
                )
            elif selected_tier == 'tier2':
                success, msg, filename = False, "Chức năng Tier 2 chưa được triển khai.", None
            else:
                success, msg, filename = False, "Lựa chọn Tier không hợp lệ.", None

            if success:
                flash(msg, 'success')
                download_link = os.path.basename(filename)
            else:
                flash(msg, 'error')
            

        except ValueError:
            flash("Lỗi: Tháng và năm phải là số nguyên hợp lệ.", 'error')
        except Exception as e:
            flash(f"Lỗi không xác định: {e}", 'error')
            print(f"Lỗi trong quá trình xử lý request: {e}", file=sys.stderr)
    
    return render_template('index.html', download_link=download_link,
                           current_month=current_month, current_year=current_year)

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)