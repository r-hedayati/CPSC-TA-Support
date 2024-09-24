# Late Submission Calculator

This script processes student assignment submissions, calculates late durations, and generates reports in either CSV or Excel format. It extracts submission data from a ZIP file, parses folder names to determine submission times, and compares them against a specified deadline.

# Prerequisites

- Python 3.10 and above
- Required Python packages: `zipfile`, `os`, `pandas`, `datetime`, `yaml`, `sys`

# Installation

1. Clone the repository or download the script files.
2. Install the required Python packages using pip/Conda:

# Configuration

Create a `config.yml` file in the same directory as the script with the following structure:
    
```yaml

deadline: "2019-09-27 22:00" #YYYY-MM-DD HH:MM in 24 hour format
zip_file_name: "AS1.zip" #make sure you have access to the directory of the zip file
csv_input_file_name: "grade_book_CourseName_AssignmentName.csv" #grade book file name
grade_book_analysis: False #Default is False, set to True if you want to analyze the grade book
late_window: 15  #late duration window in minutes for the deadline
course_name: "CourseName"
assignment_name: "AssignmentName"
filter_label: "Late" #choose beetwen "LATE", "EARLY", and "LATE (within offset)"
output_format: "excel" #choose beetwen "excel" and "csv"

```

# Usage

1. Place the ZIP file containing the assignment submissions in the same directory as the script.
2. Run the script with the configuration file as an argument:

```bash
python late_submission_calculator.py config.yml
```

If no argument is provided, the script will look for a `config.yml` file in the same directory.

# Output

The script generates a report file in either CSV or Excel format based on the configuration settings. The report includes the following columns:

- Student Name
- Submission Time
- Late Duration
- Submission Status (Early, Late, or Late within offset)
- Late Days (For tracking how many days late the submission is)

Also, the script generates a summary of the late submissions.

If a grade book is provided and the `grade_book_analysis` flag in input is `True`, the script will generate a new grade book with "Personal Days Used" column appended, which will show the total number of days used by each student for late submissions.

# Notes

- The script assumes that the folder names in the ZIP file are in the format "StudentName_StudentID" and that the submission time is encoded in the folder name.
- The timestamp format in the folder names should be: `%b %d, %Y %I:%M %p` (e.g., "Sep 11, 2024 11:59 PM").
