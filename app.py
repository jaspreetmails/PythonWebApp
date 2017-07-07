from flask import Flask, render_template, flash, request, jsonify
from wtforms import Form, TextField, TextAreaField, validators, StringField, SubmitField
import sqlite3
import csv
import xlsxwriter
# -*- coding: utf-8 -*-
# App config.
DEBUG = True
app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = '7d441f27d441f27567d441f2b6176a'

# class that inherits from the 'Form' 
# this class holds all the input feilds entered by the user
class VisitorDetailsForm(Form):
    name = TextField('Name:', validators=[validators.required()])
    email = TextField('Email:', validators=[validators.required(), validators.Length(min=6, max=35)])
    phone = TextField('Phone:', validators=[validators.required()])
    person = TextField('Person:', validators=[validators.required()])
    purpose =TextField('Purpose:', validators=[validators.required()])
    pdate = TextField('Purpose:', validators=[validators.required()])
    time = TextField('Time:', validators=[validators.required()])



@app.route("/", methods=['GET', 'POST'])
def index():
    form = VisitorDetailsForm(request.form)

    print form.errors
    if request.method == 'POST':
        name=request.form['name']
        email=request.form['email']
        phone = request.form['phone']
        person = request.form['person']
        purpose = request.form['purpose']
        pdate = request.form['pdate']
        intime =request.form['intime']
        # create DB with the data from the form.
        conn = sqlite3.connect('test.db')
        c = conn.cursor()
        c.execute("create table if not exists records  (name text, email text, phone text, person text, purpose text, pdate text, intime text)")
        rec = [name, email, phone, person, purpose, pdate, intime]
        c.execute("insert into records values (?, ?, ?, ?, ?, ?, ?)",  rec)
        conn.commit()
        flash('Logged the information successfully. ')
        c.close()
    return render_template('index.html', form=form)

# function to show visitor records
# A typical DB call to retrive all the visitors details untill that point
@app.route("/showRecords", methods=['GET', 'POST'])
def showRecords():
    try:
        conn = sqlite3.connect('test.db')
        c = conn.cursor()
        data = c.execute('SELECT * from records LIMIT 11')
        items = data.fetchall()
    except sqlite3.DatabaseError:
        flash('No data table with name records found')
    finally:
        c.close()
    return render_template('records.html', items=items)

# function to export the visitor records to an excel sheet.
# This code will save the 'Records.xlsx' to the default project location
@app.route("/export", methods=['GET', 'POST'])
def exportrecordstocsv():
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('Records.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    worksheet.write('A1', 'Name', bold)
    worksheet.write('B1', 'Email', bold)
    worksheet.write('C1', 'Phone', bold)
    worksheet.write('D1', 'Met With', bold)
    worksheet.write('E1', 'Purpose', bold)
    worksheet.write('F1', 'Date', bold)
    worksheet.write('G1', 'Time', bold)

    # Write data from the db, with row/column notation.
    message = ''
    items = []
    try:
        conn = sqlite3.connect('test.db')
        c = conn.cursor()
        data = c.execute('SELECT * from records')
        items = data.fetchall()
        j = 1
        for record in items:
            for i in range(7):
                worksheet.write(j, i, record[i])
            j = j+1
        workbook.close()
        message = 'Exported the records successfully. Please check Records.xlsx file'
    except sqlite3.DatabaseError:
        flash('No data table with name records found')
    finally:
        c.close()
    return render_template('records.html', items=items, message= message)

if __name__ == "__main__":
    app.run()