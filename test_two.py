import os
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import time
import getpass
import os
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv
import google.generativeai as genai
from langchain_core.prompts import PromptTemplate
import template_var

load_dotenv()
os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

resume_df = pd.DataFrame(columns=['resume_file_name', 'resume_file_text', 'resume_key_aspect', 'resume_score'])

# Path to the folder containing the PDFs

import tkinter as tk
from tkinter import filedialog

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title="Select Folder")
    return folder_path.replace('/','\\')

# Call the function
folder_path = select_folder()

job_description = input("Please enter JOB description : ")

# print(job_description)

# folder_path = r"C:\Users\Asus\Desktop\Code Files\Resume_score\resumes"

# Iterate through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):  # Check if the file is a PDF
        try:
            file_path = os.path.join(folder_path, filename)
            
            # Open the PDF file
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                
                # Extract text from each page
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text += page.extract_text()
            
            # Print or save the extracted text
            # print(f"Text from {filename}:\n{text}...")
            new_row = pd.DataFrame({
                'resume_file_name': [filename],
                'resume_file_text': [text],
                'resume_key_aspect': [None],
                'resume_score': [None]  # Placeholder for resume score
            })
            resume_df = pd.concat([resume_df, new_row], ignore_index=True)
    
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            

    elif filename.endswith('.docx'):
        try:
            # Open the DOCX file and extract text
            file_path = os.path.join(folder_path, filename)
            doc = docx.Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])

            # Add data to the DataFrame
            new_row = pd.DataFrame({
                'resume_file_name': [filename],
                'resume_file_text': [text],
                'resume_key_aspect': [None],
                'resume_score': [None]
            })
            resume_df = pd.concat([resume_df, new_row], ignore_index=True)
        
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            
        
    elif filename.endswith('.doc'):
        try:
            file_path = os.path.join(folder_path, filename)
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(file_path)
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()

            # Add data to the DataFrame
            new_row = pd.DataFrame({
                'resume_file_name': [filename],
                'resume_file_text': [text],
                'resume_key_aspect': [None],
                'resume_score': [None]
            })
            resume_df = pd.concat([resume_df, new_row], ignore_index=True)

        except Exception as e:
            print(f"Error reading {filename}: {e}")
            


llm = ChatGoogleGenerativeAI(
    model="gemini-1.5-flash",
    temperature=0.1,
    max_tokens=None,
    timeout=None,
    max_retries=3,
    # other params...
)


template = template_var.template_jd


prompt = PromptTemplate.from_template(template)

conversation = prompt | llm

response = conversation.invoke({"job_description_text": job_description})

job_description = response.content

print(job_description)

template = template_var.template_resume

prompt = PromptTemplate.from_template(template)

conversation = prompt | llm

for i in range(len(resume_df)):
    # resume_df["resume_file_text"][i]

    resume_text = f"""{resume_df["resume_file_text"][i]}"""

    response = conversation.invoke({"resume_text": resume_text})

    print(f"""----------------------------------------------------------------------------------{resume_df["resume_file_name"][i]}------------------------------------------------------------------------------------------------------------------""")
    
    # resume_df["resume_key_aspect"][i] = response.content
    resume_df.loc[i, "resume_key_aspect"] = response.content
    print(response.content)
    time.sleep(3)

# print(resume_df.head())

# # Assuming `df` is your DataFrame
# resume_df.to_pickle("resume_df.pkl")

# # Load the DataFrame back
# resume_df = pd.read_pickle("resume_df.pkl")

# print(resume_df.columns)
 

llm = ChatGoogleGenerativeAI(
    model="gemini-1.5-flash",
    temperature=0.1,
    max_tokens=None,
    timeout=None,
    max_retries=3,
    # other params...
)

template = template_var.template_score

prompt = PromptTemplate.from_template(template)

conversation = prompt | llm

for i in range(len(resume_df)):
    response = conversation.invoke({"resume_text": resume_df['resume_file_text'][i], "job_description": job_description})
    
    print(response.content, resume_df['resume_file_name'][i])
    resume_df.loc[i, "resume_score"] = response.content
    time.sleep(3)

resume_df.sort_values(by='resume_score', ascending=False, inplace=True)

print(resume_df.head())

resume_df.to_pickle("resume_df.pkl")

resume_df = resume_df[["resume_file_name", "resume_score"]]

file_path = 'resume_scorecard.xlsx'  # Specify the filename and path
resume_df.to_excel(file_path, index=False)  # Set index=False to avoid writing row indices

print(f"DataFrame saved to {file_path}")
