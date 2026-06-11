# Excel Questionnaire Consolidator

A Python desktop application that automates the consolidation of completed Excel questionnaire submissions into one structured report.

This project was built as a practical business automation tool. It allows a user to select a folder containing multiple completed `.xlsx` questionnaire files, extract each client’s answers, validate the submitted data, and generate a final Excel workbook with grouped responses, validation results, statistics, and charts.

## Features

* Simple desktop GUI built with Tkinter
* Select a folder containing completed questionnaire submissions
* Automatically reads multiple `.xlsx` files
* Extracts client names and questionnaire answers
* Supports dropdown, number, date, and text-based answers
* Validates missing answers and incorrect data types
* Detects invalid dropdown values
* Generates a consolidated Excel report
* Creates separate sheets for:

  * All Responses
  * Validation
  * Statistics
* Automatically generates charts for dropdown-based questions
* Saves the final report to the user’s Desktop
* Can be packaged as a standalone `.exe` file using PyInstaller

## Project Structure

```text
Questionnaire-Consolidator/
│
├── consolidate_app.py
├── questionnaire_template.xlsx
├── Submissions/
│   ├── client_submission_1.xlsx
│   ├── client_submission_2.xlsx
│   ├── client_submission_3.xlsx
│   ├── client_submission_4.xlsx
│   └── client_submission_5.xlsx
│
└── README.md
```

## How It Works

1. A standardized Excel questionnaire template is created and sent to clients.
2. Clients fill in the questionnaire and return the completed `.xlsx` files.
3. All returned files are placed inside a `Submissions` folder.
4. The Python application is launched.
5. The user selects the `Submissions` folder through the GUI.
6. The application processes every Excel file in the folder.
7. A final `consolidated_results.xlsx` file is generated on the Desktop.

## Output Workbook

The generated Excel report contains three main sheets:

### All Responses

This sheet contains one row per client and one column per questionnaire answer. It gives a clean overview of all submitted data.

### Validation

This sheet reports any issues found in the submissions, such as:

* Missing answers
* Invalid number fields
* Invalid dropdown options

If no issues are found, the sheet confirms that all submissions are valid.

### Statistics

This sheet automatically calculates counts and percentages for dropdown-based questions. It also generates bar charts to make the results easier to understand visually.

## Technologies Used

* Python
* Tkinter
* OpenPyXL
* Threading
* Microsoft Excel

## Requirements

Install the required Python package:

```bash
pip install openpyxl
```

Tkinter is included with most standard Python installations.

## How to Run

Run the application using:

```bash
python consolidate_app.py
```

Then:

1. Click **Browse**
2. Select the folder containing the completed questionnaire submissions
3. Click **Run Consolidation**
4. Open the generated `consolidated_results.xlsx` file from your Desktop

## Building an Executable

To package the application as a standalone Windows executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed consolidate_app.py
```

After building, the executable will be available inside the `dist` folder.

## Example Use Case

This tool can be used by businesses, organizations, or internal departments that collect structured information from multiple clients using Excel forms.

Instead of manually opening each returned file and copying the answers into a master spreadsheet, this application automates the full process and creates a ready-to-use report.

## Future Improvements

Possible future additions include:

* Custom questionnaire builder
* Support for multiple questionnaire templates
* Export to PDF
* More chart types
* Email integration for automatically collecting submissions
* Improved error reporting
* Configurable output location

## Status

Working prototype tested using sample questionnaire submissions.
