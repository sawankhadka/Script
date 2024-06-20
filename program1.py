import pandas as pd
import tkinter as tk
from tkinter import ttk
from pandastable import Table, TableModel

# Read the Excel file into a DataFrame
df = pd.read_excel("C:\\Users\\sawan\\Documents\\TEST DATA CONFIDENTIAL.xlsx")
df = df.sort_values(by='Roll No')
df.loc[(df['Exam Type'].isin(['Aggregate', 'SA'])) & (df['Obt Marks'] == 0), 'Obt Marks'] = ''

# Function to process and display subject data
def display_subject_data():
    selected_subject = subject_var.get()
    if selected_subject:
        # Create a new DataFrame for the selected subject
        subject_df = pd.DataFrame(columns=['Batch Name', 'Roll No', 'Student Name', 'FA', 'SA', 'Aggregate', 'Grade', 'Pass(Y)OrFail(N)'])
        
        # Filter the original DataFrame for the selected subject
        filtered_df = df[df['Subject Name'] == selected_subject]
        
        # Extract unique students' information
        students_info = filtered_df[['Batch Name', 'Roll No', 'Student Name']].drop_duplicates()
        
        # Initialize empty lists for FA, SA, Aggregate, Grade, Pass(Y)OrFail(N)
        fa_marks = []
        sa_marks = []
        aggregate_marks = []
        grades = []
        pass_fail = []
        
        # Iterate over each unique student to aggregate marks and grades
        for index, student in students_info.iterrows():
            student_marks = filtered_df[filtered_df['Roll No'] == student['Roll No']]
            
            # Initialize variables for each student
            student_fa = None
            student_sa = None
            student_aggregate = None
            student_grade = None
            student_pass_fail = None
            
            # Iterate over each row for the student to find the latest FA, SA, Aggregate marks and grade
            for _, row in student_marks.iterrows():
                if row['Exam Type'] == 'FA':
                    student_fa = row['Obt Marks']
                elif row['Exam Type'] == 'SA':
                    student_sa = row['Obt Marks']
                elif row['Exam Type'] == 'Aggregate':
                    student_aggregate = row['Obt Marks']
                    student_grade = row['Obt Grade']  # Extract grade only for Aggregate rows
                
                student_pass_fail = 'Y' if row['Is Pass'] == 'Y' else 'N'
            
            # Append values to lists
            fa_marks.append(student_fa)
            sa_marks.append(student_sa)
            aggregate_marks.append(student_aggregate)
            grades.append(student_grade)
            pass_fail.append(student_pass_fail)
        
        # Assign values to columns in the subject DataFrame
        subject_df['Batch Name'] = students_info['Batch Name'].tolist()
        subject_df['Roll No'] = students_info['Roll No'].tolist()
        subject_df['Student Name'] = students_info['Student Name'].tolist()
        subject_df['FA'] = fa_marks
        subject_df['SA'] = sa_marks
        subject_df['Aggregate'] = aggregate_marks
        subject_df['Grade'] = grades  # Assign extracted grades
        subject_df['Pass(Y)OrFail(N)'] = pass_fail
        
        # Display the data in a table
        display_table(subject_df)
    else:
        # Clear the table if no subject is selected
        display_table(pd.DataFrame())

# Function to display a table in a new window
def display_table(dataframe):
    # Create a new tkinter window
    table_window = tk.Toplevel(root)
    table_window.title("Subject Data Table")
    
    # Create a PandasTable Frame
    frame = ttk.Frame(table_window)
    frame.pack(fill='both', expand=True)
    
    # Create a table object
    pt = Table(frame, dataframe=dataframe, showtoolbar=True, showstatusbar=True)
    pt.show()

# Create the main tkinter window
root = tk.Tk()
root.title("Subject Data Viewer")

# Label for dropdown
subject_label = ttk.Label(root, text="Select Subject:")
subject_label.pack(pady=10)

# Dropdown for subjects
subject_var = tk.StringVar()
subject_dropdown = ttk.Combobox(root, textvariable=subject_var, width=30, state='readonly')
subject_dropdown.pack()

# Button to display subject data
show_data_button = ttk.Button(root, text="Show Subject Data", command=display_subject_data)
show_data_button.pack(pady=10)

# Function to populate the subject dropdown initially
def populate_subject_dropdown():
    unique_subjects = df['Subject Name'].dropna().unique().tolist()
    subject_dropdown['values'] = unique_subjects

# Populate dropdown initially
populate_subject_dropdown()

# Start the tkinter main loop
root.mainloop()
