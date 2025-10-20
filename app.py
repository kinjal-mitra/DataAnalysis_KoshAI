from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session, jsonify
from werkzeug.utils import secure_filename
import os
from utils.analyzer import process_excel_for_web, get_available_stations, get_available_pcodes
from config import Config

app = Flask(__name__)
app.config.from_object(Config)

# Store uploaded files temporarily in session
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    """Step 1: Upload file and select station"""
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    
    file = request.files['file']
    station_id = request.form.get('station_id', '').strip()
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if not station_id:
        flash('Please select a Station ID')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        try:
            # Validate Station ID
            if station_id not in ['TUS', 'CT']:
                flash('Invalid Station ID. Please select either TUS or CT')
                return redirect(url_for('index'))
            
            # Save file temporarily
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            
            # Get available PCodes
            with open(filepath, 'rb') as f:
                pcodes = get_available_pcodes(f, station_id)
            
            if not pcodes:
                os.remove(filepath)
                flash('No PCodes found for selected station')
                return redirect(url_for('index'))
            
            # Store info in session
            session['filename'] = filename
            session['station_id'] = station_id
            
            return render_template('select_pcodes.html', 
                                 pcodes=pcodes, 
                                 station_id=station_id,
                                 filename=filename)
            
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
    else:
        flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)')
        return redirect(url_for('index'))

@app.route('/analyze', methods=['POST'])
def analyze():
    """Step 2: Generate file with selected PCodes"""
    filename = session.get('filename')
    station_id = session.get('station_id')
    
    if not filename or not station_id:
        flash('Session expired. Please upload file again.')
        return redirect(url_for('index'))
    
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('File not found. Please upload again.')
        return redirect(url_for('index'))
    
    try:
        pcode1 = request.form.get('pcode1', '').strip()
        pcode2 = request.form.get('pcode2', '').strip()
        
        if not pcode1 or not pcode2:
            flash('Please select both PCodes')
            return redirect(url_for('index'))
        
        # Process the file with selected PCodes
        with open(filepath, 'rb') as f:
            result_file = process_excel_for_web(f, station_id, pcode1, pcode2)
        
        # Clean up temporary file
        os.remove(filepath)
        
        # Clear session
        session.pop('filename', None)
        session.pop('station_id', None)
        
        # Ensure we're at the start of the BytesIO object
        result_file.seek(0)
        
        # Send the analyzed file
        download_filename = f'{station_id}_analysis'
        if not download_filename.endswith('.xlsx'):
            base_name = download_filename.rsplit('.', 1)[0]
            download_filename = f'{base_name}.xlsx'
        
        return send_file(
            result_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=download_filename
        )
        
    except Exception as e:
        # Clean up on error
        if os.path.exists(filepath):
            os.remove(filepath)
        flash(f'Error processing file: {str(e)}')
        return redirect(url_for('index'))

@app.route('/cancel')
def cancel():
    """Cancel and clean up"""
    filename = session.get('filename')
    if filename:
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
    
    session.pop('filename', None)
    session.pop('station_id', None)
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)