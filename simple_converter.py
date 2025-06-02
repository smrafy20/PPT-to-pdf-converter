"""
Simple PPT to PDF Converter
Direct conversion from PowerPoint to PDF without intermediate steps.
"""

import os
import time
import pythoncom
import win32com.client
from werkzeug.utils import secure_filename

class SimplePPTConverter:
    """Simple class to convert PPT directly to PDF"""

    def __init__(self, upload_folder='uploads', download_folder='downloads'):
        self.upload_folder = upload_folder
        self.download_folder = download_folder
        self.allowed_extensions = {'ppt', 'pptx'}

        # Create directories if they don't exist
        os.makedirs(upload_folder, exist_ok=True)
        os.makedirs(download_folder, exist_ok=True)

    def check_powerpoint_availability(self):
        """
        Check if PowerPoint is available and accessible

        Returns:
            tuple: (available: bool, error_message: str)
        """
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to create PowerPoint application
            ppt = win32com.client.Dispatch("PowerPoint.Application")

            # Test basic functionality
            version = ppt.Version
            print(f"PowerPoint version detected: {version}")

            # Clean up
            ppt.Quit()
            pythoncom.CoUninitialize()

            return True, None

        except Exception as e:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False, f"PowerPoint not available: {str(e)}"

    def validate_file(self, file_path):
        """
        Validate the PowerPoint file before conversion

        Args:
            file_path: Path to the file to validate

        Returns:
            tuple: (valid: bool, error_message: str)
        """
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                return False, "File does not exist"

            # Check file size (not empty, not too large)
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                return False, "File is empty"

            if file_size > 100 * 1024 * 1024:  # 100MB limit
                return False, "File is too large (over 100MB)"

            # Check file extension
            _, ext = os.path.splitext(file_path.lower())
            if ext not in ['.ppt', '.pptx']:
                return False, "Invalid file extension. Must be .ppt or .pptx"

            # Check if file is readable
            try:
                with open(file_path, 'rb') as f:
                    # Read first few bytes to check if file is accessible
                    header = f.read(8)
                    if len(header) < 8:
                        return False, "File appears to be corrupted or incomplete"
            except Exception as e:
                return False, f"Cannot read file: {str(e)}"

            return True, None

        except Exception as e:
            return False, f"File validation error: {str(e)}"

    def allowed_file(self, filename):
        """Check if the uploaded file has an allowed extension"""
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in self.allowed_extensions
    
    def convert_ppt_to_pdf(self, ppt_file_path, output_filename=None):
        """
        Convert PPT directly to PDF with improved error handling

        Args:
            ppt_file_path: Path to the PPT/PPTX file
            output_filename: Optional custom output filename

        Returns:
            tuple: (success: bool, pdf_path: str, error_message: str)
        """
        ppt = None
        presentation = None

        try:
            print(f"Starting conversion of: {ppt_file_path}")

            # Step 1: Validate file
            valid, error_msg = self.validate_file(ppt_file_path)
            if not valid:
                return False, None, f"File validation failed: {error_msg}"

            # Step 2: Check PowerPoint availability
            available, error_msg = self.check_powerpoint_availability()
            if not available:
                return False, None, error_msg

            # Step 3: Prepare paths
            # Convert to absolute path to avoid path issues
            ppt_file_path = os.path.abspath(ppt_file_path)
            print(f"Absolute path: {ppt_file_path}")

            # Get output filename
            if output_filename:
                base_name = os.path.splitext(output_filename)[0]
            else:
                base_name = os.path.splitext(os.path.basename(ppt_file_path))[0]

            # Ensure safe filename
            safe_base_name = "".join(c for c in base_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            if not safe_base_name:
                safe_base_name = "converted_presentation"

            pdf_path = os.path.abspath(os.path.join(self.download_folder, safe_base_name + '.pdf'))
            print(f"Output PDF path: {pdf_path}")

            # Step 4: Initialize COM for this thread
            pythoncom.CoInitialize()

            # Step 5: Start PowerPoint with error handling
            print("Starting PowerPoint application...")
            try:
                ppt = win32com.client.Dispatch("PowerPoint.Application")
                # Note: Don't set Visible = False as it may cause issues in some PowerPoint versions
                print(f"PowerPoint started successfully. Version: {ppt.Version}")
            except Exception as e:
                raise Exception(f"Failed to start PowerPoint: {str(e)}")

            # Step 6: Open presentation with multiple attempts
            print("Opening presentation...")
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    if attempt == 0:
                        # First attempt: Open with minimal parameters
                        presentation = ppt.Presentations.Open(ppt_file_path, ReadOnly=True, Untitled=False, WithWindow=False)
                    elif attempt == 1:
                        # Second attempt: Open with basic parameters
                        presentation = ppt.Presentations.Open(ppt_file_path, ReadOnly=True)
                    else:
                        # Third attempt: Open with minimal parameters
                        presentation = ppt.Presentations.Open(ppt_file_path)

                    print(f"Presentation opened successfully on attempt {attempt + 1}")
                    break

                except Exception as e:
                    print(f"Attempt {attempt + 1} failed: {str(e)}")
                    if attempt == max_attempts - 1:
                        raise Exception(f"Could not open PowerPoint file after {max_attempts} attempts. Last error: {str(e)}")
                    time.sleep(1)  # Wait before retry

            # Step 7: Export to PDF
            print("Exporting to PDF...")
            try:
                # Use positional arguments for better compatibility
                presentation.ExportAsFixedFormat(
                    pdf_path,           # OutputFileName
                    2,                  # FixedFormatType (ppFixedFormatTypePDF)
                    1,                  # Intent (ppFixedFormatIntentPrint)
                    False,              # FrameSlides
                    1,                  # HandoutOrder (ppPrintHandoutHorizontalFirst)
                    1,                  # OutputType (ppPrintOutputSlides)
                    False,              # PrintHiddenSlides
                    None,               # PrintRange
                    1,                  # RangeType (ppPrintAll)
                    "",                 # SlideShowName
                    True,               # IncludeDocProps
                    True,               # KeepIRMSettings
                    True,               # DocStructureTags
                    True,               # BitmapMissingFonts
                    False               # UseDocumentICCProfile
                )
            except Exception as e:
                raise Exception(f"Failed to export to PDF: {str(e)}")

            # Step 8: Verify PDF was created
            if not os.path.exists(pdf_path):
                raise Exception("PDF file was not created")

            pdf_size = os.path.getsize(pdf_path)
            if pdf_size == 0:
                raise Exception("PDF file is empty")

            print(f"PDF created successfully: {pdf_path} (Size: {pdf_size} bytes)")
            return True, pdf_path, None

        except Exception as e:
            error_msg = f"Conversion failed: {str(e)}"
            print(f"ERROR: {error_msg}")
            return False, None, error_msg

        finally:
            # Clean up resources in reverse order
            print("Cleaning up resources...")
            try:
                if presentation:
                    presentation.Close()
                    print("Presentation closed")
            except Exception as e:
                print(f"Error closing presentation: {str(e)}")

            try:
                if ppt:
                    ppt.Quit()
                    print("PowerPoint application closed")
            except Exception as e:
                print(f"Error closing PowerPoint: {str(e)}")

            try:
                pythoncom.CoUninitialize()
                print("COM uninitialized")
            except Exception as e:
                print(f"Error uninitializing COM: {str(e)}")
    
    def save_uploaded_file(self, file):
        """
        Save uploaded file to upload directory
        
        Args:
            file: Flask file object
        
        Returns:
            tuple: (success: bool, file_path: str, error_message: str)
        """
        try:
            if file and self.allowed_file(file.filename):
                # Generate unique filename to avoid conflicts
                import uuid
                filename = secure_filename(file.filename)
                unique_filename = f"{uuid.uuid4()}_{filename}"
                file_path = os.path.join(self.upload_folder, unique_filename)
                file.save(file_path)
                return True, file_path, None
            else:
                return False, None, "Invalid file type. Please upload a PPT or PPTX file."
                
        except Exception as e:
            return False, None, f"Error saving file: {str(e)}"
    
    def cleanup_file(self, file_path):
        """Remove a file safely"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Cleaned up file: {file_path}")
        except Exception as e:
            print(f"Error cleaning up file {file_path}: {str(e)}")
