from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
from utils.analyzer import analyze_excel_file, process_excel_for_web, get_available_stations
from config import Config


app = Flask(__name__)
app.config.from_object(Config)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    
    file = request.files['file']
    station_id = request.form.get('station_id', '').strip()
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if not station_id:
        flash('Please provide a Station ID')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        try:
            # Validate Station ID (must be 'TUS' or 'CT')
            if station_id not in ['TUS', 'CT']:
                flash('Invalid Station ID. Please select either TUS or CT')
                return redirect(url_for('index'))
            
            # Save original filename
            original_filename = secure_filename(file.filename)
            
            # Process the file
            result_file = process_excel_for_web(file, station_id)
            
            # Send the analyzed file
            return send_file(
                result_file,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f'analyzed_station_{station_id}_{original_filename}'
            )
        except ValueError as ve:
            # Handle validation errors (missing columns, invalid station ID, etc.)
            flash(str(ve))
            return redirect(url_for('index'))
        except Exception as e:
            # Handle other errors
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
    else:
        flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)')
        return redirect(url_for('index'))

@app.route('/get_stations', methods=['POST'])
def get_stations():
    """
    Optional endpoint to get available stations from uploaded file.
    Can be used to populate a dropdown dynamically.
    """
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    
    file = request.files['file']
    
    if file and allowed_file(file.filename):
        try:
            stations = get_available_stations(file)
            return {'stations': stations}
        except Exception as e:
            return {'error': str(e)}, 400
    else:
        return {'error': 'Invalid file type'}, 400

if __name__ == '__main__':
    app.run(debug=True)