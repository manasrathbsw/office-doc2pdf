# 📄 Office Files to PDF Converter

A powerful Streamlit web application that converts Microsoft Word and PowerPoint files to PDF format with support for batch processing and folder structures.

## 🚀 Features

- **Multiple Input Methods**: ZIP upload, individual files, or folder drag-and-drop
- **Supported Formats**: 
  - Word Documents (`.doc`, `.docx`)
  - PowerPoint Presentations (`.ppt`, `.pptx`)
  - PDF files (copied to output)
- **Folder Structure**: Preserves original folder hierarchy
- **Batch Processing**: Convert multiple files simultaneously
- **Recursive Processing**: Handles nested folder structures
- **Real-time Feedback**: Progress indicators and conversion status
- **Download Options**: Get converted files in a ZIP archive

## 📋 System Requirements

- **Operating System**: Windows (required for COM interface)
- **Microsoft Office**: Word and PowerPoint must be installed
- **Python**: 3.7 or higher
- **Administrator Rights**: Required for COM automation

## 🔧 Installation

### 1. Clone or Download
```bash
git clone <repository-url>
cd office-to-pdf-converter
```

### 2. Install Dependencies
```bash
pip install streamlit python-docx pywin32 reportlab pathlib2
```

**Or using requirements.txt:**
```bash
pip install -r requirements.txt
```

### 3. Complete pywin32 Setup (if needed)
```bash
python -m pywin32_postinstall -install
```

## 🚀 Usage

### Starting the Application

**⚠️ IMPORTANT: Must run as Administrator**

1. Open Command Prompt as Administrator
2. Navigate to the project directory
3. Run the application:
```bash
streamlit run Word_Powerpoint_TO_pdf.py
```

### Accessing the App
- **Local URL**: http://localhost:8501
- **Network URL**: http://192.168.1.6:8501 (accessible from other devices)

## 📖 How to Use

### Method 1: ZIP File Upload 📁
1. Create a ZIP file containing your Office documents
2. Upload via the "Upload ZIP File" section
3. Choose whether to preserve folder structure
4. Download the converted PDFs

### Method 2: Individual Files 🗂️
1. Select multiple Office files using the file picker
2. Files will be processed and converted
3. Download the resulting PDF collection

### Method 3: Folder Upload 📂
1. **Option A**: Select all files from your folder (Ctrl+A) and drag them
2. **Option B**: Use the file browser to select multiple files
3. **Option C**: Create a ZIP of your folder and use Method 1

## 🔧 Supported File Types

| Format | Extension | Action |
|--------|-----------|---------|
| Word Document | `.doc`, `.docx` | Convert to PDF |
| PowerPoint | `.ppt`, `.pptx` | Convert to PDF |
| PDF | `.pdf` | Copy to output |

## 🎯 Usage Examples

### Converting a Single Folder
```
📁 My Documents/
├── 📄 Report.docx
├── 📄 Presentation.pptx
├── 📁 Subfolder/
│   ├── 📄 Analysis.docx
│   └── 📄 Charts.pptx
└── 📄 Summary.pdf
```

**Result**: All files converted to PDF while maintaining folder structure.

### Batch Processing
- Select multiple files: `Ctrl+A` in your folder
- Drag and drop into the upload area
- All supported files will be converted automatically

## ⚠️ Important Notes

### Prerequisites
- **Windows OS**: Uses Windows COM interface
- **Microsoft Office**: Word and PowerPoint must be installed
- **Administrator Mode**: Required for COM automation
- **Close Office Apps**: Close Word/PowerPoint before running

### Troubleshooting

**Common Issues:**

1. **"CoInitialize has not been called"**
   - Solution: Run as Administrator

2. **"Word.Application not found"**
   - Solution: Install Microsoft Word

3. **Files not converting**
   - Solution: Close all Office applications first

4. **Permission errors**
   - Solution: Run Command Prompt as Administrator

### Performance Tips
- Close unnecessary Office applications
- Process large files in smaller batches
- Use ZIP upload for complex folder structures
- Ensure sufficient disk space for temporary files

## 🔧 Technical Details

### Architecture
- **Frontend**: Streamlit web interface
- **Backend**: Python with COM automation
- **File Processing**: Multi-threaded conversion
- **Temporary Storage**: System temp directory

### Libraries Used
- `streamlit` - Web application framework
- `pywin32` - Windows COM interface
- `python-docx` - Word document handling
- `reportlab` - PDF utilities
- `pathlib2` - Path manipulation

## 📁 File Structure

```
office-to-pdf-converter/
├── Word_Powerpoint_TO_pdf.py    # Main application
├── requirements.txt             # Dependencies
├── README.md                   # This file
└── temp/                       # Temporary files (auto-created)
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly on Windows
5. Submit a pull request

## 📞 Support

**Common Solutions:**
- Ensure Windows Administrator mode
- Verify Microsoft Office installation
- Check Python and pip versions
- Close all Office applications

**System Requirements Check:**
```python
import win32com.client
# Should not raise errors if properly installed
```

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🔄 Version History

- **v1.0.0**: Initial release with Word/PowerPoint conversion
- **v1.1.0**: Added folder structure preservation
- **v1.2.0**: Enhanced drag-and-drop functionality
- **v1.3.0**: Improved error handling and user feedback

---

**⭐ If this helps you, please give it a star!**

**🐛 Found a bug? Please report it in the issues section.**