VacaFlow - Vacation Analytics Dashboard

📊 Complete Absence Management & Planning System

Track and analyze all 7 absence types from your Amazon logistics calendar and generate powerful, shareable PNG reports for vacation planning.

🚀 Quick Start

Open index.html in your web browser.

Upload your calendar HTML file(s) (exported from Excel).

Use Filters to narrow your data by Department, Shift, or Job Level.

Enter "Plan %" for the weeks shown.

Generate Report to create a detailed weekly vacation quote table.

Export Report (PNG) to download a high-quality, shareable image of your report, complete with a "Plan vs. Actual" summary.

📁 Project Structure

vacaflow_dashboard/
├── index.html       # Main HTML structure
├── style.css        # All styling and layout
├── script.js        # Dashboard logic, charts, and report generation
└── README.md        # This file


🎯 Features

Tracked Day Types

Icon

Type

Color

🏖️

Vacation

Lime

⏰

Flextime Day

Purple

👶

Papamonat

Magenta

📋

Special Leave

Cyan

🏥

Long Sick Leave

Red

🤒

Short Sick Leave

Brown

✈️

Active Travel Time

Light Cyan

Main Dashboard Components

Stats Cards: At-a-glance totals for employees and all 7 absence types.

4 Interactive Charts:

Absence Types Distribution (Doughnut)

Absence Days by Shift (Stacked Bar)

Daily Absence Coverage (Stacked Area)

Weekly Distribution (Stacked Bar)

Employee Table: A sortable and searchable list of all employees with their specific absence counts.

Smart Filters: Filter the entire dashboard by Department, Shift, and Job Level.

Weekly Vacation Quote Report (New!)

Dynamic Plan Inputs: Enter your planned vacation percentage for each week found in the uploaded file.

Detailed Report Table: Generates a table showing Station (Dept + Job Level) vs. Shift for all 14 weeks.

PNG Export: Downloads a high-resolution PNG of the report table.

Plan vs. Actual Summary: The exported PNG includes a professional header with a summary comparing your planned % to the actual calculated % for the filtered group.

📋 How to Use

1. Prepare Your Calendar File

Export your team calendar from Excel as HTML:

File → Save As → Web Page (.htm or .html)

2. Upload the File(s)

Drag & drop one or more HTML files onto the upload area.

The tool automatically aggregates data from all uploaded files.

3. Analyze the Dashboard

The dashboard updates instantly.

Use the "Filter Data" dropdowns to analyze specific groups.

Search the "Employee Vacation Details" table for individuals.

4. Generate & Export the Quote Report

After filtering, scroll to the "📅 Weekly Plan (%)" section (below filters) and enter your target percentages for the available weeks.

Click "Generate Report".

A detailed table will appear, showing the calculated vacation quote (%) for each Station and Shift.

Click "📥 Export Report (PNG)" to download a clean image of the report, perfect for sharing in emails, wikis, or presentations.

🔧 Technical Details

Dependencies (loaded from CDN)

Chart.js (v4.4.0): For all interactive charts.

html2canvas (v1.4.1): For generating high-quality PNG exports of the report.

Browser Compatibility

Chrome/Edge (recommended)

Firefox

Safari

🐛 Troubleshooting

No data showing?

Check that your file is in HTML format (.htm or .html).

Verify your calendar cells are colored using the standard xl68-xl75 classes.

Wrong shift/department detected?

Verify the department column follows the format: DAP8 JL1 Früh

Check that employee names contain shift information (e.g., Frühschicht 1)

PNG Export is blank?

Ensure you have clicked "Generate Report" first.

If it fails, a simple page refresh and re-upload usually solves it.

👨‍💻 Development & Credits

This tool was built by:

@tbadakar (Head Developer)

@widowitz (Support)

@ozeitsch (Support)

Version: 3.0 ("VacaFlow")
Last Updated: November 2025