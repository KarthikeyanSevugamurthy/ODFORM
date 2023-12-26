from flask import Flask, render_template, request, redirect, url_for,send_file
from docxtpl import DocxTemplate
import datetime as dt
from io import BytesIO
from docx import Document
import os
import pythoncom 



app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static' 
pythoncom.CoInitialize() 
def process_student_input(student_input):
    rows = []
    lines = student_input.strip().split('\n')
    for line in lines:
        cols = line.split()
        if len(cols) >= 4:  # Ensure at least 4 columns are present
            student = {
                "reg": cols[0],
                "name": cols[1],
                "depart": cols[2],
                "year": cols[3]
            }
            rows.append(student)
    return rows

@app.route('/', methods=['GET', 'POST'])
def index():
    download_links = None

    if request.method == 'POST':
        # Get form data
        faculty_name = request.form['faculty_name']
        event = request.form['event']
        event_dates = request.form['event_dates']
        event_place = request.form['event_place']
        designation = request.form['designation']

        # Process the list of students
        student_list = request.form['student_list']
        salesTblRows = process_student_input(student_list)

        # Create context dictionary
        context = {
            "todayDate": dt.datetime.now().strftime("%d-%b-%Y"),
            "Faculty_Name": faculty_name,
            "Event": event,
            "Event_dates": event_dates,
            "Event_place": event_place,
            "designation": designation,
            "salesTblRows": salesTblRows,
        }

        # Render context into the document object
        doc = DocxTemplate("OD1.docx")
        doc.render(context)

        # Generate unique filenames for docx and pdf

        
        docx_filename = f'OD_FORM_{dt.datetime.now().strftime("%Y%m%d%H%M%S")}.docx'

        # Save the document object as a Word file
        #doc.save(docx_filename)
        # Save the document to a BytesIO object
        doc_stream = BytesIO()
        doc.save(doc_stream)

        doc_stream.seek(0)
        return send_file(BytesIO(doc_stream.read()),
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                     as_attachment=True, download_name=docx_filename)

        
        

    return render_template('index.html', download_links=download_links)

if __name__ == '__main__':
    app.run(debug=True)
