import zipfile
import os
import pandas as pd
from datetime import datetime, timedelta
import yaml
import sys

def get_user_inputs(yaml_file):
    if not os.path.exists(yaml_file):
        raise FileNotFoundError(f"Configuration file not found: {yaml_file}")
    try:
        with open(yaml_file, 'r') as file:
            config = yaml.safe_load(file)
    except yaml.YAMLError as e:
        raise ValueError(f"Invalid YAML file: {e}")
    
    required_keys = ['deadline', 'zip_file_name', 'course_name', 'assignment_name', 'late_window', 'filter_label', 'output_format']
    for key in required_keys:
        if key not in config:
            raise KeyError(f"Missing required configuration key: {key}")
    
    deadline_str = config['deadline']
    zip_file_name = config['zip_file_name']
    late_window = config['late_window']
    course_name = config['course_name']
    assignment_name = config['assignment_name']
    filter_label = config['filter_label']
    output_format = config.get('output_format', 'excel') # Default to excel
    
    deadline = datetime.strptime(deadline_str, "%Y-%m-%d %H:%M")
    
    return deadline, zip_file_name, course_name, assignment_name, late_window, filter_label, output_format

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

def calculate_late_submissions(submissions, deadline, late_window):
    late_submissions = []
    for student_name, timestamp in submissions:
        late_duration = timestamp - deadline
        total_seconds = late_duration.total_seconds()
        total_minutes = total_seconds / 60
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        late_duration_str = f"{hours}h {minutes}m"
        timestamp_str = timestamp.strftime("%Y-%m-%d %H:%M:%S")
        if total_minutes > late_window:
            late_flag = "LATE"
        elif total_minutes <= 0:
            late_flag = "EARLY"
        elif 0 < total_minutes <= late_window:
            late_flag = "LATE (within offset)"

        late_days = (late_duration - late_window * timedelta(minutes=1)).days + 1 # Add 1 to include the day of the deadline
        late_submissions.append((student_name, timestamp_str, late_duration_str, late_flag, late_days))
    return late_submissions

def generate_output(late_submissions, course_name, assignment_name, output_format, filter_label="LATE"):
    if not late_submissions:
        raise ValueError("No late submissions.")
        
    if output_format not in ['csv', 'excel']:
        raise ValueError(f"Unsupported output format: {output_format}. Supported formats are 'csv' and 'excel'.")
    
    df = pd.DataFrame(late_submissions, columns=["Student Name", "Submission Time", "Late Duration", "Late Flag", "Late Days"])
    
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
def main():
    default_yaml_file = "config.yml"
    
    if len(sys.argv) == 2:
        yaml_file = sys.argv[1]
    else:
        print(f"No configuration file provided. Using default: {default_yaml_file}")
        yaml_file = default_yaml_file
    
    try:
        deadline, zip_file_name, course_name, assignment_name, late_window, filter_label, output_format = get_user_inputs(yaml_file)
        extract_zip(zip_file_name)
        submissions = parse_folder_names()
        late_submissions = calculate_late_submissions(submissions, deadline, late_window)
        generate_output(late_submissions, course_name, assignment_name, output_format)
    
    except (FileNotFoundError, ValueError, KeyError, OSError) as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()