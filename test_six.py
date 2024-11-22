import os
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import time
import tkinter as tk
from tkinter import filedialog
# from dotenv import load_dotenv
import google.generativeai as genai
from datetime import datetime
from tkinter import simpledialog

def get_job_description():
    """
    Prompts the user to enter a job description using a GUI dialog.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    job_description = simpledialog.askstring("Job Description", "Enter the job description:")
    return job_description


# Load API Key
# load_dotenv()
API_KEY = "your_api_key"
genai.configure(api_key=API_KEY)
# GOOGLE_API_KEY = os.environ['GOOGLE_API_KEY']

# genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')


# Templates
TEMPLATES = {
    "job_description": """The text below is a job description:

            {job_description_text}

            Your task is to summarize the job description by focusing on the following areas:

            ### 1. **Role Requirements**
            - What are the primary responsibilities and duties expected in this role?
            - What specific experiences or industry backgrounds are preferred or required?

            ### 2. **Core Skills and Technical Expertise**
            - What technical and soft skills are required or preferred for this role?
            - Is there any emphasis on proficiency level or depth in these skills?

            ### 3. **Additional Requirements or Preferences**
            - Are there any other relevant requirements, such as language proficiency, travel, or work authorization?
            - Are there specific personality traits, work styles, or soft skills emphasized?

            Provide a concise summary based on these criteria, ensuring that no extraneous information is added. This summary will be used for evaluating candidate resumes. 
            """,
    "resume":""" 
            The text below is a resume:

            {resume_text}

            Your task is to summarize the resume by focusing on the following areas:

            ### 1. **Experience Relevance** 
            - Does the candidate have relevant role-specific and industry experience? 
            - How much experience does the candidate have, and in which domain?

            ### 2. **Skills Alignment**
            - What core technical skills are demonstrated? 
            - How proficient is the candidate in these skills?

            ### 3. **Education & Certifications**
            - Does the candidate meet educational requirements? 
            - Are there any relevant certifications or additional learning?

            Provide a brief and focused summary based on these criteria. Also remember your given summary will be used for 
            evaluating the resume so do not add any extra information.
            """,
    "score": """
        Your task is to evaluate the similarity between the resume and the job description, and provide a score 
        between 0 to 100 based on how well the resume fits the given job description.

        Scoring Guidelines:
        - **90 to 100**: If the resume has More than TWO years experience in the same domain as mentioned in the job description.
        - **80 to 90**: If the resume has ONE OR  TWO years of experience in the same domain as mentioned in the job description.
        - **60 to 70**: If the resume has internship experience in the same domain as mentioned in the job description.
        - **10 to 20**: If the resume has experience in a different domain than mentioned in the job description.

        Resume Text:
        {resume_text}

        Job Description Text:
        {job_description}

        Remember to provide only the score as a single number, with no additional text.
        """,
}

# Initialize DataFrame
resume_df = pd.DataFrame(columns=['resume_file_name', 'resume_file_text', 'resume_key_aspect', 'resume_score'])

def select_folder():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder").replace('/', '\\')

def read_pdf(file_path):
    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
    return text

def read_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return ""

def read_doc(file_path):
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        return text
    except Exception as e:
        print(f"Error reading DOC {file_path}: {e}")
        return ""

def extract_text_from_files(folder_path):
    data = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.pdf'):
            text = read_pdf(file_path)
        elif filename.endswith('.docx'):
            text = read_docx(file_path)
        elif filename.endswith('.doc'):
            text = read_doc(file_path)
        else:
            continue
        data.append({'resume_file_name': filename, 'resume_file_text': text})
    return pd.DataFrame(data)

def call_genai(template, variables):
    """
    Uses the Google Generative AI to generate a response.
    Args:
        template (str): The prompt template.
        variables (dict): Variables to fill the template.
    Returns:
        str: Response text.
    """
    prompt = template.format(**variables)
    try:
        response = model.generate_content(prompt)
        # print(response.text)
        
        # response = genai.generate_message(messages=[{"content": prompt}], model="chat-bison-001")
        # return response.messages[0]["content"]
        return response.text
    except Exception as e:
        print(f"Error in API call: {e}")
        return ""


def process_resumes(resume_df, job_description):
    job_description_summary = call_genai(TEMPLATES["job_description"], {"job_description_text": job_description})

    for i in range(len(resume_df)):
        resume_text = resume_df.loc[i, "resume_file_text"]
        
        # Extract key aspects
        resume_summary = call_genai(TEMPLATES["resume"], {"resume_text": resume_text})
        resume_df.loc[i, "resume_key_aspect"] = resume_summary
        
        time.sleep(3)  # Avoid API rate limits
        
        # Score resume
        score = call_genai(TEMPLATES["score"], {
            "resume_text": resume_text,
            "job_description": job_description_summary
        })
        resume_df.loc[i, "resume_score"] = score.strip()

        print("Processing - ", resume_df.loc[i, "resume_file_name"])

        time.sleep(3)  # Avoid API rate limits
    return resume_df

# def save_results(resume_df):
#     resume_df.sort_values(by='resume_score', ascending=False, inplace=True)
#     resume_df.to_pickle("resume_df.pkl")

#     scorecard = resume_df[["resume_file_name", "resume_score"]]
#     file_path = 'resume_scorecard.xlsx'
#     scorecard.to_excel(file_path, index=False)
#     print(f"Results saved to {file_path}")


def save_results(resume_df):
    """
    Saves the processed results to an Excel file.
    Args:
        resume_df (pd.DataFrame): DataFrame with processed results.
    """
    resume_df.sort_values(by='resume_score', ascending=False, inplace=True)
    # resume_df.to_pickle("resume_df.pkl")

    scorecard = resume_df[["resume_file_name", "resume_score"]]

    # Get the user's Downloads folder
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    # Get the current date and time
    now = datetime.now()

    # Format the time
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # # Define the full file path
    file_path = os.path.join(downloads_folder, f'resume_scorecard{str(current_time)}.xlsx')

    # Save the file (assuming you have a DataFrame named `df`)
    # df.to_excel(file_path, index=False)

    # print(f"File saved to: {file_path}")

    # file_path = 'resume_scorecard.xlsx'
    scorecard.to_excel(file_path, index=False)
    # scorecard.to_excel(f'resume_scorecard{str(current_time)}.xlsx', index=False)
    # print(f"Results saved to {file_path}")


# Main Execution Flow
if __name__ == "__main__":
    folder_path = select_folder()
    job_description = get_job_description()
    resume_df = extract_text_from_files(folder_path)
    resume_df = process_resumes(resume_df, job_description)
    save_results(resume_df)
