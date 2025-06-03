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
    """Handle single or multiple file upload and start conversion"""
    try:
        # Check if files were uploaded
        if 'files' not in request.files:
            flash('No files selected')
            return redirect(url_for('index'))

        files = request.files.getlist('files')

        # Check if files were actually selected
        if not files or all(file.filename == '' for file in files):
            flash('No files selected')
            return redirect(url_for('index'))

        # Filter out empty files
        valid_files = [file for file in files if file.filename != '']

        if not valid_files:
            flash('No valid files selected')
            return redirect(url_for('index'))

        # Generate batch conversion ID for tracking
        batch_id = str(int(time.time() * 1000))  # Use timestamp as batch ID

        # Save all uploaded files and prepare for conversion
        uploaded_files = []
        failed_uploads = []

        for file in valid_files:
            success, file_path, error_msg = converter.save_uploaded_file(file)
            if success:
                uploaded_files.append({
                    'file_path': file_path,
                    'original_filename': file.filename
                })
            else:
                failed_uploads.append(f"{file.filename}: {error_msg}")

        if not uploaded_files:
            flash(f'All file uploads failed: {"; ".join(failed_uploads)}')
            return redirect(url_for('index'))

        if failed_uploads:
            flash(f'Some files failed to upload: {"; ".join(failed_uploads)}')

        # Initialize batch conversion status
        conversion_status[batch_id] = {
            'status': 'starting',
            'progress': 0,
            'message': f'Preparing to convert {len(uploaded_files)} files...',
            'batch_mode': True,
            'total_files': len(uploaded_files),
            'completed_files': 0,
            'failed_files': 0,
            'files': uploaded_files,
            'results': [],
            'failed_uploads': failed_uploads
        }

        # Start batch conversion in background thread
        thread = threading.Thread(
            target=convert_batch_background,
            args=(batch_id, uploaded_files)
        )
        thread.daemon = True
        thread.start()

        # Redirect to progress page
        return redirect(url_for('progress', conversion_id=batch_id))

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

def convert_batch_background(batch_id, uploaded_files):
    """Background function to handle batch file conversion"""
    try:
        print(f"Starting batch conversion for {batch_id}: {len(uploaded_files)} files")

        # Update status to validating
        conversion_status[batch_id].update({
            'status': 'validating',
            'progress': 5,
            'message': 'Validating files and checking PowerPoint...'
        })

        # Check PowerPoint availability first
        available, error_msg = converter.check_powerpoint_availability()
        if not available:
            conversion_status[batch_id].update({
                'status': 'error',
                'progress': 0,
                'message': f'PowerPoint not available: {error_msg}'
            })
            return

        # Update status to converting
        conversion_status[batch_id].update({
            'status': 'converting',
            'progress': 10,
            'message': f'Converting {len(uploaded_files)} files...'
        })

        time.sleep(0.5)

        # Process each file
        completed_files = 0
        failed_files = 0
        results = []

        for i, file_info in enumerate(uploaded_files):
            file_path = file_info['file_path']
            original_filename = file_info['original_filename']

            # Update progress for current file
            current_progress = 10 + (i * 80 // len(uploaded_files))
            conversion_status[batch_id].update({
                'progress': current_progress,
                'message': f'Converting {original_filename} ({i+1}/{len(uploaded_files)})...'
            })

            print(f"Converting file {i+1}/{len(uploaded_files)}: {original_filename}")

            # Perform conversion
            success, pdf_path, error_msg = converter.convert_ppt_to_pdf(
                file_path,
                original_filename
            )

            if success:
                completed_files += 1
                results.append({
                    'original_filename': original_filename,
                    'pdf_path': pdf_path,
                    'pdf_filename': os.path.basename(pdf_path),
                    'status': 'success'
                })
                print(f"✓ Conversion successful: {original_filename} -> {pdf_path}")
            else:
                failed_files += 1
                results.append({
                    'original_filename': original_filename,
                    'error_message': error_msg,
                    'status': 'failed'
                })
                print(f"✗ Conversion failed: {original_filename} - {error_msg}")

            # Clean up uploaded file
            print(f"Cleaning up uploaded file: {file_path}")
            converter.cleanup_file(file_path)

            # Update batch status
            conversion_status[batch_id].update({
                'completed_files': completed_files,
                'failed_files': failed_files,
                'results': results
            })

        # Final status update
        if failed_files == 0:
            # All files converted successfully
            conversion_status[batch_id].update({
                'status': 'completed',
                'progress': 100,
                'message': f'All {completed_files} files converted successfully!'
            })
        elif completed_files == 0:
            # All files failed
            conversion_status[batch_id].update({
                'status': 'error',
                'progress': 0,
                'message': f'All {failed_files} files failed to convert'
            })
        else:
            # Mixed results
            conversion_status[batch_id].update({
                'status': 'completed_with_errors',
                'progress': 100,
                'message': f'Batch completed: {completed_files} successful, {failed_files} failed'
            })

        print(f"Batch conversion completed: {completed_files} successful, {failed_files} failed")

    except Exception as e:
        error_message = f'Batch conversion error: {str(e)}'
        print(f"Batch conversion error: {error_message}")
        conversion_status[batch_id].update({
            'status': 'error',
            'progress': 0,
            'message': error_message
        })

        # Clean up uploaded files
        try:
            for file_info in uploaded_files:
                converter.cleanup_file(file_info['file_path'])
        except Exception as cleanup_error:
            print(f"Error during batch cleanup: {str(cleanup_error)}")
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
    """Download converted PDF file or batch of files"""
    if conversion_id not in conversion_status:
        flash('Invalid conversion ID')
        return redirect(url_for('index'))

    status = conversion_status[conversion_id]

    # Handle batch downloads (ZIP file)
    if status.get('batch_mode', False):
        return download_batch(conversion_id, status)

    # Handle single file download
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

@app.route('/download/<conversion_id>/<int:file_index>')
def download_individual_file(conversion_id, file_index):
    """Download individual PDF file from batch conversion"""
    if conversion_id not in conversion_status:
        flash('Invalid conversion ID')
        return redirect(url_for('index'))

    status = conversion_status[conversion_id]

    # Check if it's a batch conversion
    if not status.get('batch_mode', False):
        flash('Invalid download request')
        return redirect(url_for('progress', conversion_id=conversion_id))

    # Check if batch is completed
    if status['status'] not in ['completed', 'completed_with_errors']:
        flash('Batch conversion not ready for download')
        return redirect(url_for('progress', conversion_id=conversion_id))

    # Get the specific file result
    results = status.get('results', [])
    if file_index < 0 or file_index >= len(results):
        flash('Invalid file index')
        return redirect(url_for('progress', conversion_id=conversion_id))

    result = results[file_index]

    # Check if this specific file was successful
    if result['status'] != 'success':
        flash(f'File "{result["original_filename"]}" conversion failed: {result.get("error_message", "Unknown error")}')
        return redirect(url_for('progress', conversion_id=conversion_id))

    try:
        pdf_path = result['pdf_path']
        if os.path.exists(pdf_path):
            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=result['pdf_filename'],
                mimetype='application/pdf'
            )
        else:
            flash(f'PDF file not found: {result["pdf_filename"]}')
            return redirect(url_for('progress', conversion_id=conversion_id))

    except Exception as e:
        flash(f'Error downloading file: {str(e)}')
        return redirect(url_for('progress', conversion_id=conversion_id))

def download_batch(conversion_id, status):
    """Handle batch download - create ZIP file with all PDFs"""
    import zipfile
    import tempfile

    try:
        if status['status'] not in ['completed', 'completed_with_errors']:
            flash('Batch conversion not ready for download')
            return redirect(url_for('progress', conversion_id=conversion_id))

        # Get successful conversions
        successful_results = [r for r in status.get('results', []) if r['status'] == 'success']

        if not successful_results:
            flash('No files were successfully converted')
            return redirect(url_for('progress', conversion_id=conversion_id))

        # Create temporary ZIP file
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        temp_zip.close()

        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for result in successful_results:
                pdf_path = result['pdf_path']
                if os.path.exists(pdf_path):
                    # Add file to ZIP with original name
                    zipf.write(pdf_path, result['pdf_filename'])

        # Send ZIP file
        return send_file(
            temp_zip.name,
            as_attachment=True,
            download_name=f'converted_pdfs_{conversion_id}.zip',
            mimetype='application/zip'
        )

    except Exception as e:
        flash(f'Error creating batch download: {str(e)}')
        return redirect(url_for('progress', conversion_id=conversion_id))

@app.route('/cleanup/<conversion_id>')
def cleanup_conversion(conversion_id):
    """Clean up conversion files and status"""
    if conversion_id in conversion_status:
        status = conversion_status[conversion_id]

        # Handle batch cleanup
        if status.get('batch_mode', False):
            # Clean up all PDF files from batch
            for result in status.get('results', []):
                if result['status'] == 'success' and 'pdf_path' in result:
                    converter.cleanup_file(result['pdf_path'])
        else:
            # Clean up single PDF file
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
