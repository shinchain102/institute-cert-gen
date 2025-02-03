# Certificate Generator

A robust certificate generation system designed to process document templates and student data for creating customized PDF certificates.

## Features

- Support for both PowerPoint (.ppt, .pptx) and Word (.docx) templates
- Excel (.xlsx) and CSV data file support
- Variable mapping for dynamic content insertion
- Batch certificate generation
- Preserves template formatting and layout
- Automatic PDF conversion
- Clean file naming with student name and registration ID

## Prerequisites

- Python 3.x
- LibreOffice (for PowerPoint to PDF conversion)
- Required Python packages (automatically installed)

## Quick Start

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Install LibreOffice:
   - Windows: Download and install from [LibreOffice Official Website](https://www.libreoffice.org/)
   - Linux: `sudo apt-get install libreoffice`

4. Run the application:
   ```bash
   python main.py
   ```
   The server will start at `http://localhost:5000`

## Usage

1. Prepare your certificate template (.docx, .ppt, or .pptx)
   - Add variables in the template using `{{variable_name}}`
   - Example: `{{name}}`, `{{reg-id}}`

2. Prepare your data file (.xlsx or .csv)
   - Include columns matching your template variables
   - Required columns: `name`, `reg-id`

3. Upload files and generate certificates:
   - Visit `http://localhost:5000`
   - Upload your template
   - Upload your data file
   - Map the variables
   - Click "Generate Certificates"

## Detailed Documentation

For detailed setup instructions and troubleshooting:
- Windows setup guide: [Open guide/index.html](guide/index.html)
- Linux setup guide: [Open guide/index.html](guide/index.html)

## Support

For issues and questions:
1. Check the troubleshooting guide in the documentation
2. Submit an issue on the project repository
3. Contact support team

## License

MIT License - feel free to use and modify for your needs.
