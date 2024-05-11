import tkinter
import pyperclip
import tkinter as tk
from customtkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter.font import Font
from tkinter import font
import pandas as pd
import os

# from docx import Document
# from docx.shared import Inches

finalMessage = "Peter"
def process_excel_file(file_path):
    try:
        # Read The File:
        df = pd.read_excel(file_path)

        # Change it to CSV file "For Easier processing"
        csv_path = os.path.splitext(file_path)[0] + '.csv'
        df.to_csv(csv_path, index=False)

        # Remove the unrelated columns:
        df.drop(columns=['Timestamp', 'Email address', 'اسم الطالب (اختياري):', 'رقم الطالب (اختياري):',
                         'يمكنك إضافة بعض الملاحظات لو رغبت في ذلك'],
                inplace=True)

        # Rename the Columns:
        new_column_names = [str(i) for i in range(1, len(df.columns) + 1)]  # 1 -> 50
        df.columns = new_column_names

        # Replace The Choices with their scores:
        for col in df.columns:
            df[str(col)].replace(to_replace={'غير موافق تماما': '1', 'غير موافق': '2', 'موافق الى حد ما': '3',
                                             'موافق بشدة': '5', 'موافق': '4'}, inplace=True)

        # Dividing the Columns
        course = df.iloc[:, 1:26]  # index 1 -> 25 / col '2' -> '26'        p2
        prof = df.iloc[:, 26:40]  # index 26 -> 38 / col '27' -> '39'        p3
        supp = df.iloc[:, 39:52]  # index 39 -> 51 / col  '40' -> '51'       p4

        # Evaluation of the university's capabilities  p1
        st_disagree_p1 = 0  # 1
        disagree_p1 = 0  # 2
        f_agree_p1 = 0  # 3
        agree_p1 = 0  # 4
        st_agree_p1 = 0  # 5

        # Count the Values
        for value in df['1']:
            if value == "1":
                st_disagree_p1 += 1
            elif value == "2":
                disagree_p1 += 1
            elif value == "3":
                f_agree_p1 += 1
            elif value == "4":
                agree_p1 += 1
            elif value == "5":
                st_agree_p1 += 1

        # Getting p1
        t_value_p1 = st_disagree_p1 * 1 + disagree_p1 * 2 + f_agree_p1 * 3 + agree_p1 * 4 + st_agree_p1 * 5
        voters = len(df)
        p1 = t_value_p1 / voters
        n_questions = 1
        final_p1 = ((p1 / n_questions) / 5) * 100

        # Course evaluation  p2
        st_disagree_p2 = 0  # 1
        disagree_p2 = 0  # 2
        f_agree_p2 = 0  # 3
        agree_p2 = 0  # 4
        st_agree_p2 = 0  # 5

        # Count the Values
        for column in course:
            for value in df[column]:
                if value == "1":
                    st_disagree_p2 += 1
                elif value == "2":
                    disagree_p2 += 1
                elif value == "3":
                    f_agree_p2 += 1
                elif value == "4":
                    agree_p2 += 1
                elif value == "5":
                    st_agree_p2 += 1

        # Getting p2
        t_value_p2 = st_disagree_p2 * 1 + disagree_p2 * 2 + f_agree_p2 * 3 + agree_p2 * 4 + st_agree_p2 * 5
        voters = len(df)
        p2 = t_value_p2 / voters
        n_questions = 25
        final_p2 = ((p2 / n_questions) / 5) * 100

        # Evaluation of the Professor  p3
        st_disagree_p3 = 0  # 1
        disagree_p3 = 0  # 2
        f_agree_p3 = 0  # 3
        agree_p3 = 0  # 4
        st_agree_p3 = 0  # 5

        # Count the Values
        for column in prof:
            for value in df[column]:
                if value == "1":
                    st_disagree_p3 += 1
                elif value == "2":
                    disagree_p3 += 1
                elif value == "3":
                    f_agree_p3 += 1
                elif value == "4":
                    agree_p3 += 1
                elif value == "5":
                    st_agree_p3 += 1

        # Getting p3
        t_value_p3 = st_disagree_p3 * 1 + disagree_p3 * 2 + f_agree_p3 * 3 + agree_p3 * 4 + st_agree_p3 * 5
        voters = len(df)
        p3 = t_value_p3 / voters
        n_questions = 13
        final_p3 = ((p3 / n_questions) / 5) * 100

        # Evaluation of the Supporting body p4
        st_disagree_p4 = 0  # 1
        disagree_p4 = 0  # 2
        f_agree_p4 = 0  # 3
        agree_p4 = 0  # 4
        st_agree_p4 = 0  # 5

        # Count the Values
        for column in supp:
            for value in df[column]:
                if value == "1":
                    st_disagree_p4 += 1
                elif value == "2":
                    disagree_p4 += 1
                elif value == "3":
                    f_agree_p4 += 1
                elif value == "4":
                    agree_p4 += 1
                elif value == "5":
                    st_agree_p4 += 1

        # Getting p4
        t_value_p4 = st_disagree_p4 * 1 + disagree_p4 * 2 + f_agree_p4 * 3 + agree_p4 * 4 + st_agree_p4 * 5
        voters = len(df)
        p4 = t_value_p4 / voters
        n_questions = 12
        final_p4 = ((p4 / n_questions) / 5) * 100

        # Counting The Values:
        # Values Count:
        st_disagree = 0  # Score = 1
        disagree = 0  # Score = 2
        f_agree = 0  # Score = 3
        agree = 0  # Score = 4
        st_agree = 0  # Score = 5

        # Counting:
        for column in df.columns:
            for value in df[column]:
                if value == "1":
                    st_disagree += 1
                elif value == "2":
                    disagree += 1
                elif value == "3":
                    f_agree += 1
                elif value == "4":
                    agree += 1
                elif value == "5":
                    st_agree += 1

        # Getting The Final Percentage:
        t_value = st_disagree * 1 + disagree * 2 + f_agree * 3 + agree * 4 + st_agree * 5
        voters = len(df)
        per = t_value / voters
        n_questions = len(df.columns)
        final_per = ((per / n_questions) / 5) * 100

        # Display the percentages using a message box
        global finalMessage
        finalMessage = (
            f"Overall Percentage: {final_per:.2f}%\n\n"
            f"University Evaluation (p1): {final_p1:.2f}%\n"
            f"Course Evaluation (p2): {final_p2:.2f}%\n"
            f"Professor Evaluation (p3): {final_p3:.2f}%\n"
            f"Supporting Body Evaluation (p4): {final_p4:.2f}%"
        )
        takeP1(final_p1)
        takeP2(final_p2)
        takeP3(final_p3)
        takeP4(final_p4)
        takeFinal(final_per)

    except Exception as e:
        messagebox.showerror("Error", f"Error processing Excel file:{e}")
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx;.xls")])
    if file_path:
        # First, get the base name of the file (e.g., 'report.txt')
        base_name = os.path.basename(file_path)
        # Next, split the base name into the root and the extension (e.g., 'report' and '.txt')
        file_name_without_extension, _ = os.path.splitext(base_name)
        clean_name = file_name_without_extension.replace(" (Responses)", "")

        process_excel_file(file_path)
        reset_copy_buttons()
        excelName(clean_name)


def excelName(hello):
    create_bold_text_widget(responseInfo, 1, 1, hello)
    messageFinale = hello + "\n\n" + finalMessage
    pyperclip.copy(messageFinale)


def create_bold_text_widget(parent, row, column, text):
    text_var = StringVar(value=text)  # Create a StringVar with the text you want to display
    text_box = CTkEntry(parent, height=25, width=200, text_color="white", textvariable=text_var)
    text_box.grid(row=row, column=column, pady=5, padx=10)
    text_box.configure(state=tk.DISABLED)
    return text_box

# Corrected copy to clipboard function
def copy_to_clipboard(text_widget, copybutton_name):
    # Clear the clipboard first
    app.clipboard_clear()
    # Get the text from the CTkEntry widget
    text_to_copy = text_widget.get("1.0", "end-1c")
    # Add text to clipboard
    app.clipboard_append(text_to_copy)
    # Change the button text to give user feedback
    copybutton_name.configure(text='Copied!')

# Reset the text on the copy buttons
def reset_copy_buttons():
    CopyOriginal.configure(text="Copy Comments")
    CopyNew.configure(text="Copy New Comments")



app =CTk()
app.title("Response Extractor")
app.geometry("900x530")
set_appearance_mode("dark")

frame = CTkFrame(app)
frame.pack()

responseInfo = CTkLabel(frame, text="Response Information")
responseInfo.grid(row=0, column=0)

SubjectName = CTkLabel(responseInfo, text="Subject Name")
SubjectName.grid(row=1, column=0)

firstPart = CTkLabel(responseInfo, text="University Evaluation (p1):")
firstPart.grid(row=2, column=0)

secondPart = CTkLabel(responseInfo, text = "Course Evaluation (p2):")
secondPart.grid(row=3, column=0)

thirdPart = CTkLabel(responseInfo, text = "Professor Evaluation (p3):")
thirdPart.grid(row=4, column=0)

fourthPart = CTkLabel(responseInfo, text = "Supporting Body Evaluation (p4):")
fourthPart.grid(row=5, column=0)

fifthPart = CTkLabel(responseInfo, text = "Overall Percentage:")
fifthPart.grid(row=6, column=0)

def takeP1(result):
    create_bold_text_widget(responseInfo, 2,1,result)

def takeP2(result):
    create_bold_text_widget(responseInfo, 3,1,result)

def takeP3(result):
    create_bold_text_widget(responseInfo, 4,1,result)

def takeP4(result):
    create_bold_text_widget(responseInfo, 5,1,result)
def takeFinal(result):
    create_bold_text_widget(responseInfo, 6,1,result)

browse_button = CTkButton(responseInfo, text="Browse", command=browse_file, corner_radius=32)
browse_button.grid(row=7, column=0)

student_comment_section = CTkLabel(frame, text="Student Comments")
student_comment_section.grid(row=1, column=0)

originalComments = CTkTextbox(student_comment_section, scrollbar_button_color="#44DA6B", corner_radius=10,
                              border_color='#44DA6B', border_width=2, width= 350)
originalComments.grid(row= 1, column=0, padx = 50)

CopyOriginal = CTkButton(student_comment_section, text="Copy Comments",
                         corner_radius=32, command=lambda: copy_to_clipboard(originalComments, CopyOriginal))
CopyOriginal.grid(row=3, column=0, pady = 10)

CopyNew = CTkButton(student_comment_section, text="Copy New Comments",
                    corner_radius=32, command=lambda: copy_to_clipboard(newComments, CopyNew))
CopyNew.grid(row=3, column=2, pady = 10)

newComments = CTkTextbox(student_comment_section, scrollbar_button_color="#44DA6B", corner_radius=10,
                              border_color='#44DA6B', border_width=2, width= 350, )
newComments.grid(row= 1, column=2, padx = 50 )

app.mainloop()

