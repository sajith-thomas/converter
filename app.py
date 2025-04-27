from flask import Flask, render_template, request, send_file, redirect, url_for
from pdf2docx import Converter
import os
import uuid
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def safe_remove(path, max_retries=3, delay=0.1):
    """Safely remove a file with retries"""
    for _ in range(max_retries):
        try:
            if os.path.exists(path):
                os.remove(path)
                return True
        except PermissionError:
            time.sleep(delay)
    return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            return redirect(request.url)
        
        if file and file.filename.lower().endswith('.pdf'):
            # Generate unique filenames
            unique_id = uuid.uuid4()
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}.pdf")
            docx_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}.docx")
            docx_filename = f"{os.path.splitext(file.filename)[0]}.docx"
            
            try:
                # Save PDF
                file.save(pdf_path)
                
                # Convert to DOCX
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()
                
                # Send the file
                response = send_file(
                    docx_path,
                    as_attachment=True,
                    download_name=docx_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
                
                # Clean up after response is sent
                @response.call_on_close
                def cleanup():
                    safe_remove(pdf_path)
                    safe_remove(docx_path)
                
                return response
                
            except Exception as e:
                safe_remove(pdf_path)
                safe_remove(docx_path)
                return f"Conversion failed: {str(e)}"
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)