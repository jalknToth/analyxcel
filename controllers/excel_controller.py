import os
from flask import Blueprint, request, render_template
from werkzeug.utils import secure_filename
from models.excel_model import ExcelModel
from config import Config

excel_bp = Blueprint('excel', __name__)

class ExcelController:
    @staticmethod
    @excel_bp.route('/analyze', methods=['GET', 'POST'])
    def analyze_excel():

        if request.method == 'POST':
            # Check if file is present
            if 'file' not in request.files:
                return render_template('analysis.html', error='No file uploaded')
            
            file = request.files['file']
            
            # If no file is selected
            if file.filename == '':
                return render_template('analysis.html', error='No selected file')
            
            # Check file extension
            if not ExcelController._allowed_file(file.filename):
                return render_template('analysis.html', error='Invalid file type')
            
            try:
                # Secure the filename
                filename = secure_filename(file.filename)
                filepath = os.path.join(Config.UPLOAD_FOLDER, filename)
                
                # Ensure upload directory exists
                os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
                
                # Save the file
                file.save(filepath)
                
                # Analyze the file
                stats = ExcelModel.analyze_excel(filepath)
                
                return render_template('analysis.html', 
                                       stats=stats, 
                                       columns=stats['columns'], 
                                       data=stats['data'])
            
            except ValueError as e:
                return render_template('analysis.html', error=str(e))
            except Exception as e:
                return render_template('analysis.html', error='Unexpected error occurred')
        
        # GET request
        return render_template('analysis.html')

    @staticmethod
    def _allowed_file(filename):
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS
