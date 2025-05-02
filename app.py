from flask import Flask, render_template, request, redirect, send_file, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
import pandas as pd
import io
import os
import calendar
import openpyxl.utils

app = Flask(__name__)
app.secret_key = 'attendance_system_secret_key'

# Delete the database if it exists and is corrupted
db_path = 'sqlite:///attendance.db'
app.config['SQLALCHEMY_DATABASE_URI'] = db_path
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    roll_number = db.Column(db.String(50), nullable=False, unique=True)
    registration_date = db.Column(db.DateTime, default=datetime.utcnow)
    attendances = db.relationship('Attendance', backref='student', lazy=True)
    
    def __repr__(self):
        return f"{self.id} - {self.name} ({self.roll_number})"

class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey('student.id'), nullable=False)
    date = db.Column(db.Date, nullable=False, default=date.today)
    status = db.Column(db.String(20), nullable=False, default='present')  # present, absent, late
    remarks = db.Column(db.String(200))
    
    __table_args__ = (db.UniqueConstraint('student_id', 'date', name='unique_student_date'),)
    
    def __repr__(self):
        return f"Attendance: {self.student_id} on {self.date} - {self.status}"

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        name = request.form["title"]  # Keep field names for backward compatibility
        roll_number = request.form["desc"]
        student = Student(name=name, roll_number=roll_number)
        db.session.add(student)
        db.session.commit()
        flash(f"Student '{name}' added successfully!", "success")
    all_students = Student.query.all()
    return render_template("index.html", all_todos=all_students)  # Keep variable name for backward compatibility

@app.route("/delete/<int:id>")
def delete(id):
    student_to_delete = Student.query.get(id)
    name = student_to_delete.name
    db.session.delete(student_to_delete)
    db.session.commit()
    flash(f"Student '{name}' deleted successfully!", "success")
    return redirect("/")

@app.route("/update/<int:id>", methods=['GET', 'POST'])
def update(id):
    student_to_update = Student.query.get(id)
    if request.method == "POST":
        student_to_update.name = request.form['title']
        student_to_update.roll_number = request.form['desc']
        db.session.commit()
        flash(f"Student information updated successfully!", "success")
        return redirect("/")
    # Use the old field names in the template to maintain compatibility
    return render_template("update.html", todo=student_to_update)

@app.route("/search")
def search():
    search_query = request.args.get('query', '')
    if search_query:
        # Search in both name and roll number using LIKE
        search_results = Student.query.filter(
            db.or_(
                Student.name.like(f'%{search_query}%'),
                Student.roll_number.like(f'%{search_query}%')
            )
        ).all()
    else:
        search_results = Student.query.all()
    
    return render_template("index.html", all_todos=search_results, search_query=search_query)

@app.route("/mark-attendance/<int:student_id>", methods=['POST'])
def mark_attendance(student_id):
    student = Student.query.get_or_404(student_id)
    status = request.form.get('status', 'present')
    remarks = request.form.get('remarks', '')
    attendance_date = date.today()
    
    # Check if attendance already exists for this student on this date
    existing_attendance = Attendance.query.filter_by(
        student_id=student_id, 
        date=attendance_date
    ).first()
    
    if existing_attendance:
        existing_attendance.status = status
        existing_attendance.remarks = remarks
        message = f"Attendance updated for {student.name}"
    else:
        new_attendance = Attendance(
            student_id=student_id,
            status=status,
            remarks=remarks,
            date=attendance_date
        )
        db.session.add(new_attendance)
        message = f"Attendance marked for {student.name}"
    
    db.session.commit()
    flash(message, "success")
    return redirect(url_for('home'))

@app.route("/change-password", methods=['POST'])
def change_password():
    current_password = request.form.get('currentPassword')
    new_password = request.form.get('newPassword')
    confirm_password = request.form.get('confirmPassword')
    
    # Here you would typically validate the current password against stored password
    # and update it in your database. For this example, we'll just show a success message.
    
    if new_password != confirm_password:
        flash("New password and confirmation do not match!", "danger")
    else:
        # In a real app, you would update the password in the database here
        flash("Password changed successfully!", "success")
    
    # Redirect to the page the user was on
    referrer = request.referrer or url_for('home')
    return redirect(referrer)

@app.route("/attendance", methods=['GET'])
def attendance_view():
    students = Student.query.all()
    attendance_date = request.args.get('date', date.today().strftime('%Y-%m-%d'))
    
    try:
        selected_date = datetime.strptime(attendance_date, '%Y-%m-%d').date()
    except ValueError:
        selected_date = date.today()
    
    # Get attendance records for the selected date
    attendance_data = {}
    for student in students:
        attendance = Attendance.query.filter_by(
            student_id=student.id,
            date=selected_date
        ).first()
        
        attendance_data[student.id] = {
            'student': student,
            'attendance': attendance
        }
    
    return render_template(
        "attendance.html", 
        attendance_data=attendance_data,
        selected_date=selected_date
    )

@app.route("/attendance-report")
def attendance_report():
    month = request.args.get('month', datetime.now().month)
    year = request.args.get('year', datetime.now().year)
    
    try:
        month = int(month)
        year = int(year)
        if month < 1 or month > 12:
            month = datetime.now().month
    except:
        month = datetime.now().month
        year = datetime.now().year
        
    # Get all days in the selected month
    num_days = calendar.monthrange(year, month)[1]
    days = [date(year, month, day) for day in range(1, num_days + 1)]
    
    students = Student.query.all()
    
    # Prepare data for rendering
    report_data = []
    for student in students:
        student_report = {'student': student, 'days': {}}
        
        # Get all attendance records for this student in the given month
        attendance_records = Attendance.query.filter(
            Attendance.student_id == student.id,
            Attendance.date.between(days[0], days[-1])
        ).all()
        
        # Map attendance records to days
        attendance_map = {record.date: record for record in attendance_records}
        
        # Create a record for each day
        for day in days:
            if day in attendance_map:
                student_report['days'][day.day] = attendance_map[day]
            else:
                student_report['days'][day.day] = None
        
        # Calculate statistics
        present_count = sum(1 for record in attendance_records if record.status == 'present')
        absent_count = sum(1 for record in attendance_records if record.status == 'absent')
        late_count = sum(1 for record in attendance_records if record.status == 'late')
        
        student_report['stats'] = {
            'present': present_count,
            'absent': absent_count,
            'late': late_count,
            'attendance_rate': round(present_count / len(days) * 100 if days else 0, 1)
        }
        
        report_data.append(student_report)
    
    return render_template(
        "report.html",
        report_data=report_data,
        days=days,
        month_name=calendar.month_name[month],
        year=year,
        current_month=month,
        current_year=year
    )

@app.route("/download-excel")
def download_excel():
    search_query = request.args.get('query', '')
    
    # Get the same filtered data as in the search route
    if search_query:
        students = Student.query.filter(
            db.or_(
                Student.name.like(f'%{search_query}%'),
                Student.roll_number.like(f'%{search_query}%')
            )
        ).all()
    else:
        students = Student.query.all()
    
    # Convert to a list of dictionaries for pandas
    student_data = []
    for student in students:
        student_data.append({
            'ID': student.id,
            'Name': student.name,
            'Roll Number': student.roll_number,
            'Registration Date': student.registration_date
        })
    
    # Create a pandas DataFrame
    df = pd.DataFrame(student_data)
    
    # Create an in-memory Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Student List')
        
        # Auto-adjust columns' width
        worksheet = writer.sheets['Student List']
        for idx, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = column_width
    
    output.seek(0)
    
    # Generate the filename
    filename = f"student_list{'_search_'+search_query if search_query else ''}.xlsx"
    
    # Return the Excel file as a downloadable attachment
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name=filename,
        as_attachment=True
    )

@app.route("/download-attendance-report")
def download_attendance_report():
    month = request.args.get('month', datetime.now().month)
    year = request.args.get('year', datetime.now().year)
    
    try:
        month = int(month)
        year = int(year)
    except:
        month = datetime.now().month
        year = datetime.now().year
        
    # Get all days in the selected month
    num_days = calendar.monthrange(year, month)[1]
    days = [date(year, month, day) for day in range(1, num_days + 1)]
    
    students = Student.query.all()
    
    # Prepare data for Excel
    report_data = []
    for student in students:
        # Get all attendance records for this student in the given month
        attendance_records = Attendance.query.filter(
            Attendance.student_id == student.id,
            Attendance.date.between(days[0], days[-1])
        ).all()
        
        # Map attendance records to days
        attendance_map = {record.date: record.status for record in attendance_records}
        
        # Create row for this student
        student_row = {
            'ID': student.id,
            'Name': student.name,
            'Roll Number': student.roll_number
        }
        
        # Add attendance for each day - using a safer approach with day numbers as columns
        for day in days:
            student_row[f"Day_{day.day}"] = attendance_map.get(day, 'N/A')
        
        # Calculate statistics
        present_count = sum(1 for record in attendance_records if record.status == 'present')
        absent_count = sum(1 for record in attendance_records if record.status == 'absent')
        late_count = sum(1 for record in attendance_records if record.status == 'late')
        
        student_row['Present'] = present_count
        student_row['Absent'] = absent_count
        student_row['Late'] = late_count
        student_row['Attendance_Rate'] = f"{round(present_count / len(days) * 100 if days else 0, 1)}%"
        
        report_data.append(student_row)
    
    # Create a pandas DataFrame
    df = pd.DataFrame(report_data)
    
    # Create an in-memory Excel file
    output = io.BytesIO()
    
    # Use openpyxl directly to avoid column dimension issues
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f"{calendar.month_name[month]} {year}")
        
        # Get the worksheet
        worksheet = writer.sheets[f"{calendar.month_name[month]} {year}"]
        
        # Manually adjust column widths - safer approach
        # Set a reasonable width for all columns
        for col_idx in range(1, len(df.columns) + 1):
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15
    
    output.seek(0)
    
    # Generate the filename
    filename = f"attendance_report_{calendar.month_name[month]}_{year}.xlsx"
    
    # Return the Excel file as a downloadable attachment
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name=filename,
        as_attachment=True
    )

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)