from flask import Flask, request, render_template
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# File name for the Excel file
excel_file = 'attendance.xlsx'

@app.route('/')
def index():
    students = ["Abbas sayyed", "Hasan maklai", "Zulfiqar mirza", "Mehdi", "Mohammed hussain", "Jawad", "Zaid mirza", "Kaif mirza"]
    return render_template('index.html', students=students)

@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    # Create a dictionary to store attendance data
    attendance = {}
    students = ["Abbas sayyed", "Hasan maklai", "Zulfiqar mirza", "Mehdi", "Mohammed hussain", "Jawad", "Zaid mirza", "Kaif mirza"]
    
    # Populate the attendance dictionary with form data
    for student in students:
        attendance[student] = request.form.get(student)


    now = datetime.now()
    column_name = now.strftime('%Y-%m-%d %H:%M:%S')
    # Check if the Excel file exists
    if os.path.exists(excel_file):
        # Load existing data
        existing_df = pd.read_excel(excel_file, index_col='Name')
        
        # Get the current date and time
        
        # Add a new column with the current date and time as the column name
        existing_df[column_name] = None
        
        # Iterate through new attendance data
        for name, status in attendance.items():
            # If the name already exists in the DataFrame, update the attendance status
            if name in existing_df.index:
                existing_df.loc[name, column_name] = status
            # If the name does not exist, append a new row
            else:
                existing_df = existing_df.append(pd.DataFrame({column_name: [status]}, index=[name]))
        
        # Write the updated DataFrame to the Excel file
        existing_df.to_excel(excel_file)
    else:
        # If the Excel file does not exist, create it with new attendance data
        df = pd.DataFrame(attendance.items(), columns=['Name', column_name])
        df.set_index('Name', inplace=True)
        df.to_excel(excel_file)
    
    return render_template('download.html')

if __name__ == '__main__':
    app.run(debug=True)
