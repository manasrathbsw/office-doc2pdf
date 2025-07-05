# FineName: Word_Powerpoint_TO_pdf.py
# Description: A Streamlit app to convert Word and PowerPoint files to PDF format.
# Run in Windows Administrator mode only
# Open command prompt as Administrator and run:
# streamlit run Word_Powerpoint_TO_pdf.py
# Requirements:
# pip install streamlit python-docx pywin32 reportlab pathlib2
import os
from pathlib import Path
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import streamlit as st
from mimetypes import guess_type
import zipfile
import tempfile
import pythoncom
import win32com.client
from io import BytesIO

def convert_doc_to_pdf(doc_path, pdf_path):
    """Convert a DOC/DOCX file to a PDF using Microsoft Word (Windows only)."""
    pythoncom.CoInitialize()
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            doc = word.Documents.Open(str(doc_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 is the PDF format ID in Word
            doc.Close()
        except Exception as e:
            raise ValueError(f"Could not convert {doc_path}: {e}")
        finally:
            word.Quit()
    finally:
        pythoncom.CoUninitialize()


def convert_pptx_to_pdf(pptx_path, pdf_path):
    """Convert a PPT/PPTX file to a PDF using Microsoft PowerPoint (Windows only)."""
    pythoncom.CoInitialize()
    
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True  # PowerPoint sometimes needs to be visible
        
        try:
            presentation = powerpoint.Presentations.Open(str(pptx_path))
            presentation.SaveAs(str(pdf_path), FileFormat=32)  # 32 is the PDF format ID in PowerPoint
            presentation.Close()
        except Exception as e:
            raise ValueError(f"Could not convert {pptx_path}: {e}")
        finally:
            powerpoint.Quit()
    finally:
        pythoncom.CoUninitialize()


def is_valid_office_file(file_path):
    """Check if the file is a valid Office document based on its extension."""
    return file_path.suffix.lower() in ['.doc', '.docx', '.ppt', '.pptx']


def convert_file_to_pdf(file_path, output_dir):
    """Convert a single office file to PDF based on its type."""
    file_path = Path(file_path)
    output_dir = Path(output_dir)
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate PDF path
    pdf_path = output_dir / f"{file_path.stem}.pdf"
    
    try:
        if file_path.suffix.lower() in ['.doc', '.docx']:
            convert_doc_to_pdf(file_path, pdf_path)
        elif file_path.suffix.lower() in ['.ppt', '.pptx']:
            convert_pptx_to_pdf(file_path, pdf_path)
        else:
            raise ValueError(f"Unsupported file type: {file_path.suffix}")
        
        return pdf_path
    except Exception as e:
        raise ValueError(f"Failed to convert {file_path.name}: {e}")


def process_folder_recursive(folder_path, output_base_dir, preserve_structure=True):
    """Recursively process all Office files in a folder and its subfolders."""
    folder_path = Path(folder_path)
    output_base_dir = Path(output_base_dir)
    
    if not folder_path.is_dir():
        raise ValueError("The provided path is not a folder.")

    converted_files = []
    skipped_files = []
    
    # Walk through all files and subdirectories
    for file_path in folder_path.rglob('*'):
        if file_path.is_file() and file_path.name != "uploaded.zip":
            
            if is_valid_office_file(file_path):
                try:
                    # Determine output directory
                    if preserve_structure:
                        # Preserve folder structure
                        relative_path = file_path.parent.relative_to(folder_path)
                        output_dir = output_base_dir / relative_path
                    else:
                        # Put all files in the base output directory
                        output_dir = output_base_dir
                    
                    # Convert file
                    pdf_path = convert_file_to_pdf(file_path, output_dir)
                    converted_files.append({
                        'original': str(file_path.relative_to(folder_path)),
                        'pdf': str(pdf_path.relative_to(output_base_dir)),
                        'type': file_path.suffix.upper()
                    })
                    
                    st.success(f"✅ Converted: {file_path.relative_to(folder_path)} → {pdf_path.name}")
                    
                except Exception as e:
                    error_msg = f"❌ Failed to convert '{file_path.relative_to(folder_path)}': {e}"
                    st.warning(error_msg)
                    skipped_files.append(str(file_path.relative_to(folder_path)))
            
            elif file_path.suffix.lower() == '.pdf':
                # Copy existing PDF files
                try:
                    if preserve_structure:
                        relative_path = file_path.parent.relative_to(folder_path)
                        output_dir = output_base_dir / relative_path
                    else:
                        output_dir = output_base_dir
                    
                    output_dir.mkdir(parents=True, exist_ok=True)
                    pdf_copy_path = output_dir / file_path.name
                    
                    # Copy the PDF file
                    import shutil
                    shutil.copy2(file_path, pdf_copy_path)
                    
                    converted_files.append({
                        'original': str(file_path.relative_to(folder_path)),
                        'pdf': str(pdf_copy_path.relative_to(output_base_dir)),
                        'type': 'PDF (copied)'
                    })
                    
                    st.info(f"📄 Copied PDF: {file_path.relative_to(folder_path)}")
                    
                except Exception as e:
                    error_msg = f"❌ Failed to copy PDF '{file_path.relative_to(folder_path)}': {e}"
                    st.warning(error_msg)
                    skipped_files.append(str(file_path.relative_to(folder_path)))
            
            else:
                # Skip unsupported file types (but don't show warning for common system files)
                if not file_path.name.startswith('.') and file_path.suffix.lower() not in ['.txt', '.md', '.log']:
                    skipped_files.append(str(file_path.relative_to(folder_path)))
    
    return converted_files, skipped_files


def process_uploaded_files_with_structure(uploaded_files, preserve_structure=True):
    """Process uploaded files and attempt to recreate folder structure from file names."""
    with tempfile.TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir()
        output_dir.mkdir()
        
        # Save uploaded files, creating directory structure from file names
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            
            # Create subdirectories if file name contains path separators
            if '\\' in file_name or '/' in file_name:
                # Normalize path separators
                normalized_path = file_name.replace('\\', '/')
                file_path = input_dir / normalized_path
                
                # Create parent directories
                file_path.parent.mkdir(parents=True, exist_ok=True)
            else:
                file_path = input_dir / file_name
            
            # Save file
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getvalue())
        
        # Convert files
        converted_files, skipped_files = process_folder_recursive(
            input_dir, 
            output_dir, 
            preserve_structure=preserve_structure
        )
        
        return converted_files, skipped_files, output_dir


def handle_folder_upload_alternative():
    """Display alternative methods for folder upload."""
    st.markdown("""
    ### 🔄 Alternative Methods for Folder Upload:
    
    **Method 1: Create ZIP file**
    - Right-click your folder → "Send to" → "Compressed folder"
    - Upload the ZIP file using the first option
    
    **Method 2: Select all files**
    - Open your folder in File Explorer
    - Press Ctrl+A to select all files
    - Drag and drop them into the upload area
    
    **Method 3: PowerShell/Command Line**
    If you have many nested folders, create a ZIP using:
    ```
    Compress-Archive -Path "C:\\YourFolder\\*" -DestinationPath "C:\\YourFolder.zip"
    ```
    """)


def create_folder_structure_info(uploaded_files):
    """Analyze uploaded files and show the detected folder structure."""
    if not uploaded_files:
        return
    
    folder_structure = {}
    for file in uploaded_files:
        file_name = file.name
        if '\\' in file_name or '/' in file_name:
            # Normalize path separators
            normalized_path = file_name.replace('\\', '/')
            parts = normalized_path.split('/')
            
            current_level = folder_structure
            for part in parts[:-1]:  # All except the file name
                if part not in current_level:
                    current_level[part] = {}
                current_level = current_level[part]
            
            # Add file to the structure
            if '_files' not in current_level:
                current_level['_files'] = []
            current_level['_files'].append(parts[-1])
        else:
            # File in root
            if '_files' not in folder_structure:
                folder_structure['_files'] = []
            folder_structure['_files'].append(file_name)
    
    if folder_structure:
        st.write("**📁 Detected folder structure:**")
        display_folder_structure(folder_structure, "")


def display_folder_structure(structure, prefix=""):
    """Recursively display folder structure."""
    for key, value in structure.items():
        if key == '_files':
            for file_name in value:
                st.write(f"{prefix}📄 {file_name}")
        else:
            st.write(f"{prefix}📁 {key}/")
            display_folder_structure(value, prefix + "  ")


def create_zip_from_folder(folder_path, zip_name="converted_files.zip"):
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in folder_path.rglob('*'):
            if file_path.is_file():
                # Add file to zip with relative path
                arcname = file_path.relative_to(folder_path)
                zipf.write(file_path, arcname)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def handle_uploaded_zip(uploaded_zip, preserve_structure=True):
    """Extract and process uploaded ZIP file."""
    if not zipfile.is_zipfile(uploaded_zip):
        st.error("The uploaded file is not a valid ZIP file.")
        return [], []

    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract ZIP
        extract_dir = Path(temp_dir) / "extracted"
        extract_dir.mkdir()
        
        zip_path = Path(temp_dir) / "uploaded.zip"
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.getvalue())

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_dir)

        # Process files
        output_dir = Path(temp_dir) / "converted"
        st.info("Files extracted. Converting to PDFs...")
        
        converted_files, skipped_files = process_folder_recursive(
            extract_dir, 
            output_dir, 
            preserve_structure=preserve_structure
        )
        
        # Create a ZIP of results
        if converted_files:
            result_zip_data = create_zip_from_folder(output_dir)
            
            st.download_button(
                "📦 Download Converted PDFs",
                result_zip_data,
                "converted_files.zip",
                "application/zip",
                key="download_zip"
            )
        
        return converted_files, skipped_files


# Streamlit App
st.set_page_config(
    page_title="Office to PDF Converter",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Office Files to PDF Converter")
st.markdown("Convert **DOC, DOCX, PPT, PPTX** files to PDF format")

# Add system requirements info
st.info("⚠️ **System Requirements**: This app requires Microsoft Word and PowerPoint to be installed on Windows. Run in Administrator mode for best results.")

# Create three columns for different input methods
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📁 Upload ZIP File")
    uploaded_zip = st.file_uploader(
        "Upload a ZIP file containing Office files",
        type=["zip"],
        accept_multiple_files=False,
        key="zip_uploader"
    )

with col2:
    st.subheader("🗂️ Upload Individual Files")
    uploaded_files = st.file_uploader(
        "Upload individual Office files",
        type=["doc", "docx", "ppt", "pptx"],
        accept_multiple_files=True,
        key="files_uploader"
    )

with col3:
    st.subheader("📂 Drag & Drop Folder")
    st.markdown("""
    **📌 How to upload a folder:**
    1. Select all files in your folder (Ctrl+A)
    2. Drag them to the upload area below
    3. Or click 'Browse files' and select multiple files
    """)
    
    folder_files = st.file_uploader(
        "Select all files from your folder",
        type=["doc", "docx", "ppt", "pptx", "pdf"],
        accept_multiple_files=True,
        key="folder_uploader",
        help="Select all files from your folder structure. File paths will be preserved based on file names."
    )

# Options
st.subheader("⚙️ Options")
preserve_structure = st.checkbox(
    "Preserve folder structure", 
    value=True, 
    help="Keep the original folder structure in the output ZIP file"
)

# Process uploaded files
if uploaded_zip:
    st.markdown("---")
    st.subheader("🔄 Processing ZIP File")
    
    try:
        converted_files, skipped_files = handle_uploaded_zip(uploaded_zip, preserve_structure)
        
        if converted_files:
            st.success(f"🎉 Conversion complete! {len(converted_files)} files were processed.")
            
            # Show conversion summary
            with st.expander("📋 Conversion Summary", expanded=True):
                st.write("**Successfully converted:**")
                for file_info in converted_files:
                    st.write(f"• {file_info['original']} → {file_info['pdf']} ({file_info['type']})")
                
                if skipped_files:
                    st.write("**Skipped files:**")
                    for skipped in skipped_files:
                        st.write(f"• {skipped}")
        else:
            st.warning("No files were converted. Please check that your ZIP contains valid Office files.")
    
    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
        st.write("**Troubleshooting tips:**")
        st.write("- Ensure Microsoft Word and PowerPoint are installed")
        st.write("- Check that the ZIP file contains valid Office files")
        st.write("- Try running the app as administrator")
        st.write("- Close any open Office applications before running")

# Process folder upload (drag and drop)
if folder_files:
    st.markdown("---")
    st.subheader("🔄 Processing Folder Upload")
    
    # Show detected folder structure
    create_folder_structure_info(folder_files)
    
    try:
        converted_files, skipped_files, output_dir = process_uploaded_files_with_structure(
            folder_files, 
            preserve_structure
        )
        
        if converted_files:
            st.success(f"🎉 Conversion complete! {len(converted_files)} files were processed.")
            
            # Create download
            result_zip_data = create_zip_from_folder(output_dir)
            st.download_button(
                "📦 Download Converted PDFs",
                result_zip_data,
                "converted_files.zip",
                "application/zip",
                key="download_folder"
            )
            
            # Show summary
            with st.expander("📋 Conversion Summary", expanded=True):
                for file_info in converted_files:
                    st.write(f"✅ {file_info['original']} → {file_info['pdf']} ({file_info['type']})")
                
                if skipped_files:
                    st.write("**Skipped files:**")
                    for skipped in skipped_files:
                        st.write(f"• {skipped}")
        else:
            st.warning("No files were converted.")
            handle_folder_upload_alternative()
    
    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
        handle_folder_upload_alternative()

# Process individual files
if uploaded_files:
    st.markdown("---")
    st.subheader("🔄 Processing Individual Files")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir()
        output_dir.mkdir()
        
        # Save uploaded files
        for uploaded_file in uploaded_files:
            file_path = input_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getvalue())
        
        # Convert files
        converted_files, skipped_files = process_folder_recursive(
            input_dir, 
            output_dir, 
            preserve_structure=False
        )
        
        if converted_files:
            st.success(f"🎉 Conversion complete! {len(converted_files)} files were processed.")
            
            # Create download
            result_zip_data = create_zip_from_folder(output_dir)
            st.download_button(
                "📦 Download Converted PDFs",
                result_zip_data,
                "converted_files.zip",
                "application/zip",
                key="download_individual"
            )
            
            # Show summary
            with st.expander("📋 Conversion Summary", expanded=True):
                for file_info in converted_files:
                    st.write(f"✅ {file_info['original']} → {file_info['pdf']} ({file_info['type']})")
                
                if skipped_files:
                    st.write("**Skipped files:**")
                    for skipped in skipped_files:
                        st.write(f"• {skipped}")

# Add footer with usage instructions
st.markdown("---")
st.markdown("""
### 📖 Usage Instructions:
1. **ZIP Upload**: Upload a ZIP file containing your Office files (supports nested folders)
2. **Individual Files**: Upload multiple Office files directly
3. **Folder Upload**: Select all files from your folder structure (preserves paths)
4. **Folder Structure**: Choose whether to preserve the original folder structure
5. **Download**: Get your converted PDFs in a ZIP file

### 🔧 Supported Formats:
- **Word Documents**: .doc, .docx
- **PowerPoint Presentations**: .ppt, .pptx
- **PDF Files**: Existing PDFs will be copied to the output

### 💡 Tips for Folder Upload:
- **Method 1**: Create a ZIP file of your folder and use the ZIP upload option
- **Method 2**: Select all files (Ctrl+A) from your folder and drag them to the upload area
- **Method 3**: Use file browser to select multiple files from nested folders
- **Folder Structure**: File paths in names (e.g., "subfolder/file.docx") will be preserved

### 🔧 Technical Notes:
- Close all Office applications before conversion
- Run in Administrator mode for best results
- Large files may take longer to process
- Nested folder structures are automatically detected and preserved
""")



# import os
# from pathlib import Path
# from docx import Document
# from reportlab.lib.pagesizes import letter
# from reportlab.pdfgen import canvas
# import streamlit as st
# from mimetypes import guess_type
# import zipfile
# import tempfile
# import pythoncom
# import win32com.client
# from io import BytesIO

# def convert_doc_to_pdf(doc_path, pdf_path):
#     """Convert a DOC/DOCX file to a PDF using Microsoft Word (Windows only)."""
#     pythoncom.CoInitialize()
    
#     try:
#         word = win32com.client.Dispatch("Word.Application")
#         word.Visible = False
        
#         try:
#             doc = word.Documents.Open(str(doc_path))
#             doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 is the PDF format ID in Word
#             doc.Close()
#         except Exception as e:
#             raise ValueError(f"Could not convert {doc_path}: {e}")
#         finally:
#             word.Quit()
#     finally:
#         pythoncom.CoUninitialize()


# def convert_pptx_to_pdf(pptx_path, pdf_path):
#     """Convert a PPT/PPTX file to a PDF using Microsoft PowerPoint (Windows only)."""
#     pythoncom.CoInitialize()
    
#     try:
#         powerpoint = win32com.client.Dispatch("PowerPoint.Application")
#         powerpoint.Visible = True  # PowerPoint sometimes needs to be visible
        
#         try:
#             presentation = powerpoint.Presentations.Open(str(pptx_path))
#             presentation.SaveAs(str(pdf_path), FileFormat=32)  # 32 is the PDF format ID in PowerPoint
#             presentation.Close()
#         except Exception as e:
#             raise ValueError(f"Could not convert {pptx_path}: {e}")
#         finally:
#             powerpoint.Quit()
#     finally:
#         pythoncom.CoUninitialize()


# def is_valid_office_file(file_path):
#     """Check if the file is a valid Office document based on its extension."""
#     return file_path.suffix.lower() in ['.doc', '.docx', '.ppt', '.pptx']


# def convert_file_to_pdf(file_path, output_dir):
#     """Convert a single office file to PDF based on its type."""
#     file_path = Path(file_path)
#     output_dir = Path(output_dir)
    
#     # Create output directory if it doesn't exist
#     output_dir.mkdir(parents=True, exist_ok=True)
    
#     # Generate PDF path
#     pdf_path = output_dir / f"{file_path.stem}.pdf"
    
#     try:
#         if file_path.suffix.lower() in ['.doc', '.docx']:
#             convert_doc_to_pdf(file_path, pdf_path)
#         elif file_path.suffix.lower() in ['.ppt', '.pptx']:
#             convert_pptx_to_pdf(file_path, pdf_path)
#         else:
#             raise ValueError(f"Unsupported file type: {file_path.suffix}")
        
#         return pdf_path
#     except Exception as e:
#         raise ValueError(f"Failed to convert {file_path.name}: {e}")


# def process_folder_recursive(folder_path, output_base_dir, preserve_structure=True):
#     """Recursively process all Office files in a folder and its subfolders."""
#     folder_path = Path(folder_path)
#     output_base_dir = Path(output_base_dir)
    
#     if not folder_path.is_dir():
#         raise ValueError("The provided path is not a folder.")

#     converted_files = []
#     skipped_files = []
    
#     # Walk through all files and subdirectories
#     for file_path in folder_path.rglob('*'):
#         if file_path.is_file() and file_path.name != "uploaded.zip":
            
#             if is_valid_office_file(file_path):
#                 try:
#                     # Determine output directory
#                     if preserve_structure:
#                         # Preserve folder structure
#                         relative_path = file_path.parent.relative_to(folder_path)
#                         output_dir = output_base_dir / relative_path
#                     else:
#                         # Put all files in the base output directory
#                         output_dir = output_base_dir
                    
#                     # Convert file
#                     pdf_path = convert_file_to_pdf(file_path, output_dir)
#                     converted_files.append({
#                         'original': str(file_path.relative_to(folder_path)),
#                         'pdf': str(pdf_path.relative_to(output_base_dir)),
#                         'type': file_path.suffix.upper()
#                     })
                    
#                     st.success(f"✅ Converted: {file_path.relative_to(folder_path)} → {pdf_path.name}")
                    
#                 except Exception as e:
#                     error_msg = f"❌ Failed to convert '{file_path.relative_to(folder_path)}': {e}"
#                     st.warning(error_msg)
#                     skipped_files.append(str(file_path.relative_to(folder_path)))
            
#             elif file_path.suffix.lower() == '.pdf':
#                 # Copy existing PDF files
#                 try:
#                     if preserve_structure:
#                         relative_path = file_path.parent.relative_to(folder_path)
#                         output_dir = output_base_dir / relative_path
#                     else:
#                         output_dir = output_base_dir
                    
#                     output_dir.mkdir(parents=True, exist_ok=True)
#                     pdf_copy_path = output_dir / file_path.name
                    
#                     # Copy the PDF file
#                     import shutil
#                     shutil.copy2(file_path, pdf_copy_path)
                    
#                     converted_files.append({
#                         'original': str(file_path.relative_to(folder_path)),
#                         'pdf': str(pdf_copy_path.relative_to(output_base_dir)),
#                         'type': 'PDF (copied)'
#                     })
                    
#                     st.info(f"📄 Copied PDF: {file_path.relative_to(folder_path)}")
                    
#                 except Exception as e:
#                     error_msg = f"❌ Failed to copy PDF '{file_path.relative_to(folder_path)}': {e}"
#                     st.warning(error_msg)
#                     skipped_files.append(str(file_path.relative_to(folder_path)))
            
#             else:
#                 # Skip unsupported file types (but don't show warning for common system files)
#                 if not file_path.name.startswith('.') and file_path.suffix.lower() not in ['.txt', '.md', '.log']:
#                     skipped_files.append(str(file_path.relative_to(folder_path)))
    
#     return converted_files, skipped_files


# def create_zip_from_folder(folder_path, zip_name="converted_files.zip"):
#     """Create a ZIP file from a folder, preserving the folder structure."""
#     folder_path = Path(folder_path)
#     zip_buffer = BytesIO()
    
#     with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
#         for file_path in folder_path.rglob('*'):
#             if file_path.is_file():
#                 # Add file to zip with relative path
#                 arcname = file_path.relative_to(folder_path)
#                 zipf.write(file_path, arcname)
    
#     zip_buffer.seek(0)
#     return zip_buffer.getvalue()


# def handle_uploaded_zip(uploaded_zip, preserve_structure=True):
#     """Extract and process uploaded ZIP file."""
#     if not zipfile.is_zipfile(uploaded_zip):
#         st.error("The uploaded file is not a valid ZIP file.")
#         return [], []

#     with tempfile.TemporaryDirectory() as temp_dir:
#         # Extract ZIP
#         extract_dir = Path(temp_dir) / "extracted"
#         extract_dir.mkdir()
        
#         zip_path = Path(temp_dir) / "uploaded.zip"
#         with open(zip_path, "wb") as f:
#             f.write(uploaded_zip.getvalue())

#         with zipfile.ZipFile(zip_path, "r") as zip_ref:
#             zip_ref.extractall(extract_dir)

#         # Process files
#         output_dir = Path(temp_dir) / "converted"
#         st.info("Files extracted. Converting to PDFs...")
        
#         converted_files, skipped_files = process_folder_recursive(
#             extract_dir, 
#             output_dir, 
#             preserve_structure=preserve_structure
#         )
        
#         # Create a ZIP of results
#         if converted_files:
#             result_zip_data = create_zip_from_folder(output_dir)
            
#             st.download_button(
#                 "📦 Download Converted PDFs",
#                 result_zip_data,
#                 "converted_files.zip",
#                 "application/zip",
#                 key="download_zip"
#             )
        
#         return converted_files, skipped_files


# # Streamlit App
# st.set_page_config(
#     page_title="Office to PDF Converter",
#     page_icon="📄",
#     layout="wide"
# )

# st.title("📄 Office Files to PDF Converter")
# st.markdown("Convert **DOC, DOCX, PPT, PPTX** files to PDF format")

# # Add system requirements info
# st.info("⚠️ **System Requirements**: This app requires Microsoft Word and PowerPoint to be installed on Windows. Run in Administrator mode for best results.")

# # Create two columns for different input methods
# col1, col2 = st.columns(2)

# with col1:
#     st.subheader("📁 Upload ZIP File")
#     uploaded_zip = st.file_uploader(
#         "Upload a ZIP file containing Office files",
#         type=["zip"],
#         accept_multiple_files=False,
#         key="zip_uploader"
#     )

# with col2:
#     st.subheader("🗂️ Upload Individual Files")
#     uploaded_files = st.file_uploader(
#         "Upload individual Office files",
#         type=["doc", "docx", "ppt", "pptx"],
#         accept_multiple_files=True,
#         key="files_uploader"
#     )

# # Options
# st.subheader("⚙️ Options")
# preserve_structure = st.checkbox(
#     "Preserve folder structure", 
#     value=True, 
#     help="Keep the original folder structure in the output ZIP file"
# )

# # Process uploaded files
# if uploaded_zip:
#     st.markdown("---")
#     st.subheader("🔄 Processing ZIP File")
    
#     try:
#         converted_files, skipped_files = handle_uploaded_zip(uploaded_zip, preserve_structure)
        
#         if converted_files:
#             st.success(f"🎉 Conversion complete! {len(converted_files)} files were processed.")
            
#             # Show conversion summary
#             with st.expander("📋 Conversion Summary", expanded=True):
#                 st.write("**Successfully converted:**")
#                 for file_info in converted_files:
#                     st.write(f"• {file_info['original']} → {file_info['pdf']} ({file_info['type']})")
                
#                 if skipped_files:
#                     st.write("**Skipped files:**")
#                     for skipped in skipped_files:
#                         st.write(f"• {skipped}")
#         else:
#             st.warning("No files were converted. Please check that your ZIP contains valid Office files.")
    
#     except Exception as e:
#         st.error(f"An error occurred during processing: {e}")
#         st.write("**Troubleshooting tips:**")
#         st.write("- Ensure Microsoft Word and PowerPoint are installed")
#         st.write("- Check that the ZIP file contains valid Office files")
#         st.write("- Try running the app as administrator")
#         st.write("- Close any open Office applications before running")

# # Process individual files
# if uploaded_files:
#     st.markdown("---")
#     st.subheader("🔄 Processing Individual Files")
    
#     with tempfile.TemporaryDirectory() as temp_dir:
#         input_dir = Path(temp_dir) / "input"
#         output_dir = Path(temp_dir) / "output"
#         input_dir.mkdir()
#         output_dir.mkdir()
        
#         # Save uploaded files
#         for uploaded_file in uploaded_files:
#             file_path = input_dir / uploaded_file.name
#             with open(file_path, "wb") as f:
#                 f.write(uploaded_file.getvalue())
        
#         # Convert files
#         converted_files, skipped_files = process_folder_recursive(
#             input_dir, 
#             output_dir, 
#             preserve_structure=False
#         )
        
#         if converted_files:
#             st.success(f"🎉 Conversion complete! {len(converted_files)} files were processed.")
            
#             # Create download
#             result_zip_data = create_zip_from_folder(output_dir)
#             st.download_button(
#                 "📦 Download Converted PDFs",
#                 result_zip_data,
#                 "converted_files.zip",
#                 "application/zip",
#                 key="download_individual"
#             )
            
#             # Show summary
#             with st.expander("📋 Conversion Summary", expanded=True):
#                 for file_info in converted_files:
#                     st.write(f"✅ {file_info['original']} → {file_info['pdf']} ({file_info['type']})")
                
#                 if skipped_files:
#                     st.write("**Skipped files:**")
#                     for skipped in skipped_files:
#                         st.write(f"• {skipped}")

# # Add footer with usage instructions
# st.markdown("---")
# st.markdown("""
# ### 📖 Usage Instructions:
# 1. **ZIP Upload**: Upload a ZIP file containing your Office files (supports nested folders)
# 2. **Individual Files**: Upload multiple Office files directly
# 3. **Folder Structure**: Choose whether to preserve the original folder structure
# 4. **Download**: Get your converted PDFs in a ZIP file

# ### 🔧 Supported Formats:
# - **Word Documents**: .doc, .docx
# - **PowerPoint Presentations**: .ppt, .pptx
# - **PDF Files**: Existing PDFs will be copied to the output

# ### 💡 Tips:
# - Close all Office applications before conversion
# - Run in Administrator mode for best results
# - Large files may take longer to process
# """)