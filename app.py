from flask import Flask, render_template, request, send_file
from pdf2docx import Converter
import os
import uuid

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file uploaded", 400
        
        file = request.files['file']
        if file.filename == '':
            return "No file selected", 400
        
        if file and file.filename.lower().endswith('.pdf'):
            # Temporary file handling
            pdf_path = f"temp_{uuid.uuid4()}.pdf"
            docx_path = f"temp_{uuid.uuid4()}.docx"
            file.save(pdf_path)
            
            # Convert PDF to DOCX
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # Send file and cleanup
            response = send_file(
                docx_path,
                as_attachment=True,
                download_name="converted.docx"
            )
            
            # Cleanup after response
            @response.call_on_close
            def cleanup():
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(docx_path):
                    os.remove(docx_path)
            
            return response
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
