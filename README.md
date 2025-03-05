# Project Overview

This project is a Flask web application designed to process PDF files containing order data. It extracts relevant information from the uploaded PDFs, aggregates the data, and allows users to download the aggregated report in both CSV and DOCX formats.

## Features

- Upload multiple PDF files containing order information.
- Extract and aggregate order data from the uploaded PDFs.
- Download the aggregated report in CSV format.
- Download the aggregated report in DOCX format.

## File Structure

```
webapp
├── templates
│   └── upload_form.html      # HTML form for file upload and report display
├── web_order_report.py       # Main application file with Flask routes and logic
├── requirements.txt          # List of dependencies for the project
└── README.md                 # Documentation for the project
```

## Setup Instructions

1. **Clone the Repository**
   ```bash
   git clone <repository-url>
   cd webapp
   ```

2. **Install Dependencies**
   It is recommended to use a virtual environment. You can create one using:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```
   Then install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**
   Start the Flask application:
   ```bash
   python web_order_report.py
   ```
   The application will be accessible at `http://127.0.0.1:5000`.

## Usage

- Navigate to the application in your web browser.
- Use the upload form to select and upload PDF files.
- After processing, the aggregated report will be displayed.
- You can download the report in CSV or DOCX format using the provided buttons.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.