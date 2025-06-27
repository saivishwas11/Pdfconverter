# File Converter - Streamlit App

A modern web application for converting files between different formats using Streamlit.

## Features

- **File Upload**: Upload multiple files (.docx, .xlsx, .html)
- **Format Conversion**: Convert files to PDF or Excel format
- **Batch Processing**: Convert multiple files at once
- **Download**: Download converted files individually or as a ZIP archive
- **Modern UI**: Clean, responsive interface built with Streamlit

## Supported Formats

### Input Formats
- 📝 Microsoft Word (.docx)
- 📊 Microsoft Excel (.xlsx)
- 🌐 HTML files (.html)

### Output Formats
- 📄 PDF
- 📊 Excel (.xlsx)

## Local Development

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Installation

1. Clone the repository:
```bash
git clone <your-repo-url>
cd Pdfmaker
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the Streamlit app:
```bash
streamlit run streamlit_app.py
```

4. Open your browser and navigate to `http://localhost:8501`

## Deployment to Streamlit Cloud

### Step 1: Prepare Your Repository

1. Make sure your repository is on GitHub, GitLab, or Bitbucket
2. Ensure your main file is named `streamlit_app.py`
3. Verify that `requirements.txt` is in the root directory

### Step 2: Deploy to Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub/GitLab/Bitbucket account
3. Click "New app"
4. Fill in the deployment form:
   - **Repository**: Select your repository
   - **Branch**: Select your main branch (usually `main` or `master`)
   - **Main file path**: Enter `streamlit_app.py`
   - **App URL**: Choose a unique URL for your app
5. Click "Deploy"

### Step 3: Configure Advanced Settings (Optional)

If you need to configure advanced settings, you can add a `.streamlit/config.toml` file to your repository (already included).

## Usage

1. **Upload Files**: Use the sidebar to upload one or more files
2. **Select Output Format**: Choose between PDF or Excel
3. **Convert**: Click the "Convert Files" button
4. **Download**: Download your converted files

## File Size Limits

- Maximum file size: 100 MB per file
- Multiple files can be uploaded simultaneously

## Troubleshooting

### Common Issues

1. **File Upload Errors**: Ensure your files are in the supported formats (.docx, .xlsx, .html)
2. **Conversion Failures**: Check that your files are not corrupted and are valid
3. **Download Issues**: Make sure your browser allows downloads

### Platform-Specific Notes

- **Windows**: The app uses `pywin32` for some conversions
- **Linux/Mac**: Some Windows-specific features may not be available

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

If you encounter any issues or have questions, please open an issue on the GitHub repository. 