"""
Flask Web Application for PPT to PDF Conversion
Simple web interface for uploading and converting PowerPoint files to PDF.

@author: Modified for web application
"""

import os
import threading
import time
from flask import Flask, request, render_template, send_file, jsonify, redirect, url_for, flash
from werkzeug.utils import secure_filename
from simple_converter import SimplePPTConverter

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'  # Change this in production
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Initialize converter
converter = SimplePPTConverter()

# Store conversion status for progress tracking
conversion_status = {}

@app.route('/')
def index():
    """Main page with upload form"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and start conversion"""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file selected')
            return redirect(url_for('index'))
        
        file = request.files['file']
        
        # Check if file was actually selected
        if file.filename == '':
            flash('No file selected')
            return redirect(url_for('index'))
        
        # Save uploaded file
        success, file_path, error_msg = converter.save_uploaded_file(file)
        
        if not success:
            flash(error_msg)
            return redirect(url_for('index'))
        
        # Generate conversion ID for tracking
        conversion_id = str(int(time.time() * 1000))  # Use timestamp as ID
        
        # Initialize conversion status
        conversion_status[conversion_id] = {
            'status': 'starting',
            'progress': 0,
            'message': 'Preparing conversion...',
            'file_path': file_path,
            'original_filename': file.filename
        }
        
        # Start conversion in background thread
        thread = threading.Thread(
            target=convert_file_background, 
            args=(conversion_id, file_path, file.filename)
        )
        thread.daemon = True
        thread.start()
        
        # Redirect to progress page
        return redirect(url_for('progress', conversion_id=conversion_id))
        
    except Exception as e:
        flash(f'Error processing upload: {str(e)}')
        return redirect(url_for('index'))

def convert_file_background(conversion_id, file_path, original_filename):
    """Background function to handle file conversion"""
    try:
        print(f"Starting background conversion for {conversion_id}: {file_path}")

        # Update status to validating
        conversion_status[conversion_id].update({
            'status': 'validating',
            'progress': 10,
            'message': 'Validating file and checking PowerPoint...'
        })

        # Add a small delay to ensure status is updated
        import time
        time.sleep(0.5)

        # Check PowerPoint availability first
        available, error_msg = converter.check_powerpoint_availability()
        if not available:
            conversion_status[conversion_id].update({
                'status': 'error',
                'progress': 0,
                'message': f'PowerPoint not available: {error_msg}'
            })
            return

        # Update status to converting
        conversion_status[conversion_id].update({
            'status': 'converting',
            'progress': 30,
            'message': 'Converting PowerPoint to PDF...'
        })

        time.sleep(0.5)

        # Perform direct conversion
        print(f"Starting conversion: {file_path} -> PDF")
        success, pdf_path, error_msg = converter.convert_ppt_to_pdf(
            file_path,
            original_filename
        )

        if success:
            print(f"Conversion successful: {pdf_path}")
            conversion_status[conversion_id].update({
                'status': 'completed',
                'progress': 100,
                'message': 'Conversion completed successfully!',
                'pdf_path': pdf_path,
                'pdf_filename': os.path.basename(pdf_path)
            })
        else:
            print(f"Conversion failed: {error_msg}")
            conversion_status[conversion_id].update({
                'status': 'error',
                'progress': 0,
                'message': error_msg or 'Conversion failed'
            })

        # Clean up uploaded file
        print(f"Cleaning up uploaded file: {file_path}")
        converter.cleanup_file(file_path)

    except Exception as e:
        error_message = f'Conversion error: {str(e)}'
        print(f"Background conversion error: {error_message}")
        conversion_status[conversion_id].update({
            'status': 'error',
            'progress': 0,
            'message': error_message
        })

        # Clean up uploaded file
        try:
            converter.cleanup_file(file_path)
        except Exception as cleanup_error:
            print(f"Error during cleanup: {str(cleanup_error)}")
            pass

@app.route('/progress/<conversion_id>')
def progress(conversion_id):
    """Show conversion progress page"""
    if conversion_id not in conversion_status:
        flash('Invalid conversion ID')
        return redirect(url_for('index'))
    
    return render_template('progress.html', conversion_id=conversion_id)

@app.route('/status/<conversion_id>')
def get_status(conversion_id):
    """API endpoint to get conversion status"""
    if conversion_id not in conversion_status:
        return jsonify({'error': 'Invalid conversion ID'}), 404
    
    status = conversion_status[conversion_id]
    return jsonify(status)

@app.route('/download/<conversion_id>')
def download_file(conversion_id):
    """Download converted PDF file"""
    if conversion_id not in conversion_status:
        flash('Invalid conversion ID')
        return redirect(url_for('index'))
    
    status = conversion_status[conversion_id]
    
    if status['status'] != 'completed' or 'pdf_path' not in status:
        flash('File not ready for download')
        return redirect(url_for('progress', conversion_id=conversion_id))
    
    try:
        pdf_path = status['pdf_path']
        if os.path.exists(pdf_path):
            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=status['pdf_filename'],
                mimetype='application/pdf'
            )
        else:
            flash('File not found')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'Error downloading file: {str(e)}')
        return redirect(url_for('index'))

@app.route('/cleanup/<conversion_id>')
def cleanup_conversion(conversion_id):
    """Clean up conversion files and status"""
    if conversion_id in conversion_status:
        status = conversion_status[conversion_id]
        
        # Clean up PDF file
        if 'pdf_path' in status:
            converter.cleanup_file(status['pdf_path'])
        
        # Remove from status tracking
        del conversion_status[conversion_id]
    
    return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    flash('File is too large. Maximum size is 50MB.')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_error(e):
    """Handle internal server errors"""
    flash('An internal error occurred. Please try again.')
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("Starting PPT2PDF Web Application...")
    print("=" * 50)

    # Create necessary directories
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('downloads', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    os.makedirs('static', exist_ok=True)
    print("✓ Directories created/verified")

    # Check PowerPoint availability at startup
    print("Checking PowerPoint availability...")
    available, error_msg = converter.check_powerpoint_availability()
    if available:
        print("✓ PowerPoint is available and ready")
    else:
        print(f"⚠ WARNING: PowerPoint check failed: {error_msg}")
        print("The application will start, but conversions may fail.")
        print("Please ensure Microsoft PowerPoint is properly installed.")

    print("=" * 50)
    print("Application starting...")
    print("Access the application at: http://localhost:5000")
    print("Press Ctrl+C to stop the server")
    print("=" * 50)

    # Run the Flask application
    app.run(debug=True, host='0.0.0.0', port=5000)
