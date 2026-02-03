from flask import Flask, render_template, request, redirect, url_for, flash
import mysql.connector
import openpyxl


app = Flask(__name__)
app.secret_key = "excel"


db = mysql.connector.connect(
    host="localhost",
    user="root",
    password="root",
    database="college"
)

@app.route('/')
def index():
    return render_template('register.html')

@app.route('/register', methods=['POST'])
def register():
    fname = request.form['first_name']
    lname = request.form['last_name']
    dob = request.form['dob']
    phone = request.form['mobile']
    email = request.form['email']
    gender = request.form['gender']
    department = request.form['department']
    course = request.form['course']

    try:
        cursor = db.cursor()


        cursor.execute("""
            INSERT INTO students (fname, lname,dob,phone,email,gender,department,course)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
        """, (fname,lname,dob,phone,email,gender, department,course))
        db.commit()


        try:
            wb = openpyxl.load_workbook("students.xlsx")
            sheet = wb.active
        except:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["First Name","Last Name","DOB","Phone","Email","Gender","Department","Courses",])

        sheet.append([fname,lname,dob,phone,email,gender,department,course])
        wb.save("students.xlsx")

        flash("Registration Successful!", "success")

    except mysql.connector.IntegrityError:
        flash("Email already registered!", "error")

    except mysql.connector.Error as err:
        flash(f"Database Error: {err}", "error")

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
