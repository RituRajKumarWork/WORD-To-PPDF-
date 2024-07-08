# working with app.py , index.html, covert.html, download.html enough


from flask import Flask, request, render_template, send_file
import os
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        return render_template('convert.html', filename=file.filename)
    return 'File upload failed'


@app.route('/convert/<filename>')
def convert_file(filename):
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_path = os.path.join(UPLOAD_FOLDER, f'{os.path.splitext(filename)[0]}.pdf')

    try:
        doc = Document(input_path)
        pdf = canvas.Canvas(output_path, pagesize=letter)
        pdf.setFont("Times-Roman", 12)

        width, height = letter
        y = height - 40

        for paragraph in doc.paragraphs:
            text = paragraph.text
            pdf.drawString(30, y, text)
            y -= 14
            if y < 40:
                pdf.showPage()
                pdf.setFont("Times-Roman", 12)
                y = height - 40

        pdf.save()
    except Exception as e:
        return f"Conversion failed: {e}", 500

    return render_template('download.html', filename=os.path.basename(output_path))


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
