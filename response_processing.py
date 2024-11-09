import os
import pandas as pd
import re

def get_cell_value(df, row, col):
    """
    Safely retrieves a cell value from a DataFrame, handling out-of-bounds and empty cells.
    """
    if row < len(df) and col < len(df.columns):
        value = df.iat[row, col]
        return value if pd.notna(value) else None
    return None

def get_combined_cell_values(df, start_row, col):
    """
    Combines text from a starting cell downwards in the specified column until an empty cell is encountered.
    Useful for extracting multi-line answers.
    """
    combined_text = []
    row = start_row
    
    while row < len(df):
        value = get_cell_value(df, row, col)
        if value is None:  # Stop if an empty cell is encountered
            break
        combined_text.append(str(value))
        row += 1
    
    return " ".join(combined_text) if combined_text else None

def extract_student_answers(file_path):
    """
    Extracts answers from the specified Excel sheets and cells for a single student.
    """
    answers = []
    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            # Extract answers from the "Stock" sheet
            stock_sheet = xls.parse("Stock")
            answers.extend([
                get_cell_value(stock_sheet, 7, 1),
                get_cell_value(stock_sheet, 9, 1),
                get_cell_value(stock_sheet, 11, 1),
                get_cell_value(stock_sheet, 13, 1),
                get_cell_value(stock_sheet, 15, 1)
            ])
            
            # Extract answers from the "Metro1" sheet
            metro1_sheet = xls.parse("Metro1")
            answers.extend([
                get_cell_value(metro1_sheet, 1, 0),
                get_cell_value(metro1_sheet, 3, 0),
                get_cell_value(metro1_sheet, 5, 0)
            ])
            
            # Extract combined answer from "Metro2" sheet for question 9
            metro2_sheet = xls.parse("Metro2")
            answers.append(get_combined_cell_values(metro2_sheet, 1, 0))  # Start from cell A2
        
    except (ValueError, FileNotFoundError, pd.errors.EmptyDataError, OSError, zipfile.BadZipFile) as e:
        print(f"Error reading file {file_path}: {e}")
        return [None] * 9  # Return a list of `None` for each answer if there's an error

    return answers

def read_student_answers(folder_path):
    """
    Reads all student files in the folder and compiles their answers into a single DataFrame.
    """
    all_answers = []
    
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            student_id = re.search(r"Exam Deliverables_(e\d+)_attempt", filename).group(1)
            file_path = os.path.join(folder_path, filename)

            # Extract answers for this student
            student_answers = extract_student_answers(file_path)
            
            # Append student ID and answers to the list
            all_answers.append([student_id] + student_answers)
    
    # Create DataFrame with the required format
    columns = ["Student Id"] + list(range(1, 10))
    all_answers_df = pd.DataFrame(all_answers, columns=columns)
    return all_answers_df#.set_index("Student Id")


folder_path = "/Users/rrishabh/Documents/Thesis related docs/Thesis Data/Students"
compiled_df = read_student_answers(folder_path)

# Save the processed files to CSV/xlsx
compiled_df.to_excel("/Users/rrishabh/Documents/Thesis related docs/Thesis Data/compiled_student_answers.xlsx", index=False)