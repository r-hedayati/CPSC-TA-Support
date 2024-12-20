import zipfile
import os
import pandas as pd
from datetime import datetime, timedelta
import yaml
import sys
import math

def get_user_inputs(yaml_file):
    if not os.path.exists(yaml_file):
        raise FileNotFoundError(f"Configuration file not found: {yaml_file}")
    try:
        with open(yaml_file, 'r') as file:
            config = yaml.safe_load(file)
    except yaml.YAMLError as e:
        raise ValueError(f"Invalid YAML file: {e}")
    
    required_inputs = ['deadline', 'zip_file_name', 'grade_book_csv_input_file_name', 'grade_book_analysis', 'course_name', 'assignment_name', 'personal_days_column_id', 'late_window', 'filter_label', 'early_offset_on', 'output_format']
    for input in required_inputs:
        if input not in config:
            raise KeyError(f"Missing required configuration input: {input}")
    
    deadline_str = config['deadline']
    zip_file_name = config['zip_file_name']
    csv_file_name = config['grade_book_csv_input_file_name']
    grade_book_flag = config['grade_book_analysis']
    personal_days_column = config['personal_days_column_id']
    late_window = config['late_window']
    course_name = config['course_name']
    assignment_name = config['assignment_name']
    filter_label = config['filter_label']
    early_offset_on = config['early_offset_on']
    output_format = config.get('output_format', 'excel') # Default to excel
    
    deadline = datetime.strptime(deadline_str, "%Y-%m-%d %H:%M")
    
    return deadline, zip_file_name, csv_file_name, grade_book_flag, course_name, assignment_name, personal_days_column, late_window, filter_label, early_offset_on, output_format

def extract_zip(zip_file_name):
    if not os.path.exists(zip_file_name):
        raise FileNotFoundError(f"The zip file {zip_file_name} does not exist.")
    
    try:
        with zipfile.ZipFile(zip_file_name, 'r') as zip_ref:
            zip_ref.extractall("extracted")
    except zipfile.BadZipFile as e:
        raise ValueError(f"Error extracting zip file: {e}")

def parse_folder_names():
    folders = os.listdir("extracted")
    submissions = {}
    for folder in folders:
        parts = folder.split(" - ")
        if len(parts) == 3:
            student_id, student_name, timestamp_str = parts
            timestamp = datetime.strptime(timestamp_str, "%b %d, %Y %I%M %p")
            # Check if the student already has a submission
            if student_name in submissions:
                # Keep the latest submission
                if timestamp > submissions[student_name][1]:
                    submissions[student_name] = (student_name, timestamp)
            else:
                submissions[student_name] = (student_name, timestamp)
    # Convert the dictionary to a list of tuples
    return list(submissions.values())

def calculate_late_submissions(submissions, deadline, late_window, early_offset_on):
    late_submissions = []
    late_submissions_dict = {}
    if early_offset_on == False:
        print(f"Early submissions will be considered as Early. The offset will not be included in the late duration.")
    else:
        print(f"Early submissions will be considered as early. The offset will be included in the late duration.")
    for student_name, timestamp in submissions:
        late_duration = timestamp - deadline
        if late_duration < timedelta(0) and early_offset_on == False:
            late_duration = timedelta(0)
        total_seconds = late_duration.total_seconds()
        total_minutes = total_seconds / 60
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        late_duration_str = f"{hours}h {minutes}m"
        timestamp_str = timestamp.strftime("%Y-%m-%d %H:%M:%S")
        if total_minutes > late_window:
            late_flag = "Over-Full" #LATE
        elif total_minutes <= 0 :
            late_flag = "Available" #EARLY
        elif 0 < total_minutes < late_window:
            late_flag = "Available"
        elif total_minutes == late_window:
            late_flag = "Full"
        # elif 0 < total_minutes <= late_window:
        #     late_flag = "LATE (within offset)" #LATE (within offset)

        late_days = (late_duration - late_window * timedelta(minutes=1)).days + 1 # Add 1 to include the day of the deadline
        late_submissions.append((student_name, timestamp_str, late_duration_str, late_flag, late_days))
        late_submissions_dict[student_name] = {"Student Name": student_name, "Submission Time": timestamp_str,
                                               "Late Duration": late_duration_str, "Late Flag": late_flag,
                                               "Late Days": late_days}
    return late_submissions_dict

def generate_output(late_submissions, course_name, assignment_name, output_format, filter_label="Over-Full"):
    if not late_submissions:
        raise ValueError("No late submissions.")
        
    if output_format not in ['csv', 'excel']:
        raise ValueError(f"Unsupported output format: {output_format}. Supported formats are 'csv' and 'excel'.")
    
    # df = pd.DataFrame(late_submissions, columns=["Student Name", "Submission Time", "Late Duration", "Late Flag", "Late Days"])
    df = pd.DataFrame.from_dict(late_submissions, orient='index')
    try: 
        if output_format == 'csv':
            output_file = f"{course_name}_{assignment_name}_submissions.csv"
            df.to_csv(output_file, index=False)
            print(f"CSV sheet generated: {output_file}")
            
            # Filter late submissions
            late_df = df[df["Late Flag"] == filter_label]
            late_output_file = f"{course_name}_{assignment_name}_late_submissions.csv"
            late_df.to_csv(late_output_file, index=False)
            print(f"Late submissions CSV sheet generated: {late_output_file}")
        elif output_format == 'excel':
            output_file = f"{course_name}_{assignment_name}_submissions.xlsx"
            df.to_excel(output_file, index=False)
            print(f"Excel sheet generated: {output_file}")
            
            # Filter late submissions
            late_df = df[df["Late Flag"] == filter_label]
            late_output_file = f"{course_name}_{assignment_name}_late_submissions.xlsx"
            late_df.to_excel(late_output_file, index=False)
            print(f"Late submissions Excel sheet generated: {late_output_file}")
    except IOError as e:
        print(f"An error occurred while writing the file: {e}")

def grade_book_report(grade_book_csv, grade_book_flag, personal_days_column_id, late_submissions, course_name, assignment_name):
    if grade_book_flag == True:
        try:
            df = pd.read_csv(grade_book_csv)
        except pd.errors.ParserError as e:
            raise ValueError(f"Error reading CSV file: {e}")
        
        personal_days_column_name = "Personal Days Used"
        
        if personal_days_column_name in df.columns:
            personal_days_column_id = df.columns.get_loc(personal_days_column_name)
        else:
            if personal_days_column_id >= len(df.columns):
                raise ValueError(f"Column index {personal_days_column_id} is out of range for the CSV file.")
            df.insert(personal_days_column_id, personal_days_column_name, 0)

        # Add "Student Name" column
        df["Student Name"] = df["First Name"] + " " + df["Last Name"]
        
        for student_name, late_info in late_submissions.items():
            if student_name in df['Student Name'].values:
                current_days_used = df.loc[df['Student Name'] == student_name, df.columns[personal_days_column_id]].values[0]

                if current_days_used == "" or math.isnan(current_days_used):
                    current_days_used = 0
                
                additional_days_used = max(late_info['Late Days'], 0)
                df.loc[df['Student Name'] == student_name, df.columns[personal_days_column_id]] = current_days_used + additional_days_used
            else:
                print(f"Student {student_name} not found in grade book.")
        
        df.drop(columns=["Student Name"], inplace=True)
        output_file = f"grade_book_{course_name}_{assignment_name}.csv"
        df.to_csv(output_file, index=False)
        print(f"Updated grade book CSV file generated: {output_file}")
    else: 
        print("Grade book analysis not enabled.")
    



def main():
    default_yaml_file = "config.yml"
    
    if len(sys.argv) == 2:
        yaml_file = sys.argv[1]
    else:
        print(f"No configuration file provided. Using default: {default_yaml_file}")
        yaml_file = default_yaml_file
    
    try:
        deadline, zip_file_name, csv_file_name, grade_book_flag, course_name, assignment_name, personal_days_column, late_window, filter_label, early_offset_on, output_format = get_user_inputs(yaml_file)
        # extract_zip(zip_file_name)
        submissions = parse_folder_names()
        late_submissions = calculate_late_submissions(submissions, deadline, late_window, early_offset_on)
        generate_output(late_submissions, course_name, assignment_name, output_format)

        # Update the grade book
        grade_book_report(csv_file_name, grade_book_flag, personal_days_column, late_submissions, course_name, assignment_name)
    
    except (FileNotFoundError, ValueError, KeyError, OSError) as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()