import streamlit as st
import pandas as pd
import os
from datetime import datetime

# File to store attendance
ATTENDANCE_FILE = 'attendance.xlsx'

# Initialize Excel file or ensure Attendance Count column exists
def initialize_excel():
    try:
        if not os.path.exists(ATTENDANCE_FILE):
            df = pd.DataFrame(columns=['Roll Number', 'Name', 'Attendance Count'])
            df.to_excel(ATTENDANCE_FILE, index=False)
        else:
            # Check if Attendance Count column exists in existing file
            df = pd.read_excel(ATTENDANCE_FILE)
            if 'Attendance Count' not in df.columns:
                df['Attendance Count'] = 0  # Add column with default value 0
                df.to_excel(ATTENDANCE_FILE, index=False)
    except PermissionError:
        st.error(f"Permission denied when accessing {ATTENDANCE_FILE}. Ensure the file is not open and you have write permissions.")
        return False
    return True


# Mark attendance for given roll numbers (last three digits) with cumulative count in date column
def mark_attendance(three_digit_rolls, date_str):
    try:
        df = pd.read_excel(ATTENDANCE_FILE)
        
        # Ensure Attendance Count column exists
        if 'Attendance Count' not in df.columns:
            df['Attendance Count'] = 0
            df.to_excel(ATTENDANCE_FILE, index=False)
        
        date_col = date_str
        
        # Initialize date column if it doesn't exist
        if date_col not in df.columns:
            df[date_col] = 0  # Default to 0 for absent students
        
        # Match three-digit roll numbers to full roll numbers
        matched_rolls = []
        for three_digit in three_digit_rolls:
            # Extract last three digits of full roll numbers
            matches = df[df['Roll Number'].astype(str).str[-3:] == str(three_digit).zfill(3)]['Roll Number'].tolist()
            if len(matches) == 0:
                st.warning(f"No roll number found ending with {three_digit}. Please add the student first.")
            elif len(matches) > 1:
                st.warning(f"Multiple roll numbers found ending with {three_digit}: {matches}. Please use unique last three digits.")
            else:
                matched_rolls.append(matches[0])
        
        if not matched_rolls:
            st.error("No valid roll numbers matched. Attendance not marked.")
            return None
        
        # Mark attendance with the current Attendance Count for matched students
        for roll in matched_rolls:
            # Increment Attendance Count
            df.loc[df['Roll Number'] == roll, 'Attendance Count'] += 1
            # Set date column to the updated Attendance Count
            df.loc[df['Roll Number'] == roll, date_col] = df.loc[df['Roll Number'] == roll, 'Attendance Count']
        
        # Save updated Excel
        df.to_excel(ATTENDANCE_FILE, index=False)
        
        # Return present students for display
        return df[df[date_col] > 0][['Roll Number', 'Name', date_col, 'Attendance Count']]
    except PermissionError:
        st.error(f"Permission denied when writing to {ATTENDANCE_FILE}. Ensure the file is not open and you have write permissions.")
        return None
    except Exception as e:
        st.error(f"An error occurred while marking attendance: {str(e)}")
        return None

# Streamlit interface
def main():
    st.title("Student Attendance System")
    if not initialize_excel():
        return

    # Get current date for column name and display
    current_date = datetime.now().strftime('%Y-%m-%d')
    st.write(f"Marking attendance for {current_date}")

    # Input form
    with st.form("attendance_form"):
        roll_numbers_input = st.text_area("Enter Roll Numbers (comma-separated)", placeholder="e.g., 001,002")
        submit_button = st.form_submit_button("Mark Attendance")

    if submit_button:
        if roll_numbers_input:
            # Process three-digit roll numbers
            three_digit_rolls = [roll.strip() for roll in roll_numbers_input.split(',') if roll.strip()]
            
            # Mark attendance
            present_students = mark_attendance(three_digit_rolls, current_date)
            if present_students is not None:
                st.success(f"Attendance marked for {current_date}")
                
                # Display present students
                if not present_students.empty:
                    st.subheader(f"Students Present for {current_date}")
                    st.dataframe(present_students, use_container_width=True)
                else:
                    st.warning("No students marked as present. Check if roll numbers exist in the database.")
        else:
            pass

    if os.path.exists(ATTENDANCE_FILE):
        with open(ATTENDANCE_FILE, "rb") as file:
            st.download_button(
                label="Download attendance.xlsx",
                data=file,
                file_name="attendance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
