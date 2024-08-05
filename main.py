import os
import pandas as pd
from Modules import *

###### Change here ######
reviewer = "Ali Emre Kaya"
question = "Local and Global Inversions"
file_name = "example.docx"
excel_file = "exampe.xlsx"
#########################

feedback_df = pd.read_excel(excel_file, header=None, names=["Name", "Score", "Feedback"])
students = {
    "student1": "-",
    "student2": "-",
    "student3": "-",
    "student4": "-"
}


if __name__ == "__main__":
    for index, row in feedback_df.iterrows():
        student_name = row["Name"]
        header = f"{question} ({reviewer}) ({row['Score']})"
        feedback = f"Feedback: {row['Feedback']} \n\n\n"
        
        if student_name in students:
            folder_id = students[student_name]

            existing_file = file_exists(service, folder_id, file_name)
            if existing_file:
                doc_id = existing_file['id']
                add_text_to_google_doc(docs_service, doc_id, feedback, header)
                print(f"{student_name}: feedback added to the existing file '{file_name}' in folder ID: {folder_id}.")
            else:
                doc_id = create_google_doc(service, folder_id, file_name)
                add_text_to_google_doc(docs_service, doc_id, feedback, header)
                print(f"{student_name}: new file '{file_name}' created and feedback added in folder ID: {folder_id}.")