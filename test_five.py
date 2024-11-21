import os
import PyPDF2
import pandas as pd
import docx
import win32com.client as win32
import time
import logging
from concurrent.futures import ThreadPoolExecutor
from tkinter import filedialog, Tk
from dotenv import load_dotenv
import google.generativeai as genai
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import PromptTemplate
import template_var

# Initialize Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("resume_processing.log"), logging.StreamHandler()]
)

# Load API Key
load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=API_KEY)

# Constants
TEMPLATES = {
    "job_description": f"{template_var.template_jd}",
    "resume": f"{template_var.template_resume}",
    "score": f"{template_var.template_score}",
}

# Functions
def select_folder():
    """
    Prompts the user to select a folder via a GUI dialog.
    Returns:
        str: Selected folder path.
    """
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Folder")
    return folder_path.replace('/', '\\') if folder_path else None

def read_file(file_path, extension):
    """
    Reads text from a file based on its extension.
    Args:
        file_path (str): Path to the file.
        extension (str): File extension (e.g., .pdf, .docx, .doc).
    Returns:
        str: Extracted text from the file.
    """
    try:
        if extension == ".pdf":
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                return "".join(page.extract_text() for page in reader.pages)
        elif extension == ".docx":
            doc = docx.Document(file_path)
            return "\n".join(paragraph.text for paragraph in doc.paragraphs)
        elif extension == ".doc":
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(file_path)
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            return text
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {e}")
        return ""

def extract_text_from_files(folder_path):
    """
    Extracts text from all supported files in the specified folder.
    Args:
        folder_path (str): Path to the folder containing files.
    Returns:
        pd.DataFrame: DataFrame containing file names and extracted text.
    """
    supported_extensions = {".pdf", ".docx", ".doc"}
    data = []

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        extension = os.path.splitext(filename)[1].lower()

        if extension in supported_extensions:
            logging.info(f"Processing file: {filename}")
            text = read_file(file_path, extension)
            data.append({'resume_file_name': filename, 'resume_file_text': text})
    return pd.DataFrame(data)

def initialize_conversation(template):
    """
    Initializes a conversation using a template and LangChain's LLM integration.
    Args:
        template (str): Template content for the conversation.
    Returns:
        Callable: Configured conversation object.
    """
    llm = ChatGoogleGenerativeAI(
        model="gemini-1.5-flash",
        temperature=0.1,
        max_tokens=None,
        timeout=None,
        max_retries=3
    )
    return PromptTemplate.from_template(template) | llm

def process_resumes(resume_df, job_description):
    """
    Processes resumes by extracting key aspects and scoring them.
    Args:
        resume_df (pd.DataFrame): DataFrame containing resumes.
        job_description (str): Job description text.
    Returns:
        pd.DataFrame: Updated DataFrame with key aspects and scores.
    """
    # Initialize conversations
    conversation_jd = initialize_conversation(TEMPLATES["job_description"])
    conversation_resume = initialize_conversation(TEMPLATES["resume"])
    conversation_score = initialize_conversation(TEMPLATES["score"])

    # Process JD
    logging.info("Processing job description...")
    jd_response = conversation_jd.invoke({"job_description_text": job_description})
    processed_jd = jd_response.content

    def process_single_resume(row):
        try:
            # Extract key aspects
            resume_response = conversation_resume.invoke({"resume_text": row["resume_file_text"]})
            key_aspects = resume_response.content

            # Score resume
            score_response = conversation_score.invoke({
                "resume_text": row["resume_file_text"],
                "job_description": processed_jd
            })
            score = score_response.content

            return key_aspects, score
        except Exception as e:
            logging.error(f"Error processing resume {row['resume_file_name']}: {e}")
            return None, None

    # Use multithreading for faster processing
    with ThreadPoolExecutor() as executor:
        results = executor.map(process_single_resume, resume_df.to_dict(orient="records"))

    # Update DataFrame
    for idx, (key_aspects, score) in enumerate(results):
        resume_df.loc[idx, "resume_key_aspect"] = key_aspects
        resume_df.loc[idx, "resume_score"] = score

    return resume_df

def save_results(resume_df):
    """
    Saves the processed results to an Excel file.
    Args:
        resume_df (pd.DataFrame): DataFrame with processed results.
    """
    resume_df.sort_values(by='resume_score', ascending=False, inplace=True)
    resume_df.to_pickle("resume_df.pkl")

    scorecard = resume_df[["resume_file_name", "resume_score"]]
    file_path = 'resume_scorecard.xlsx'
    scorecard.to_excel(file_path, index=False)
    logging.info(f"Results saved to {file_path}")

# Main Execution Flow
if __name__ == "__main__":
    folder_path = select_folder()
    if not folder_path:
        logging.error("No folder selected. Exiting...")
        exit()

    job_description = input("Please enter JOB description: ").strip()
    if not job_description:
        logging.error("No job description provided. Exiting...")
        exit()

    resume_df = extract_text_from_files(folder_path)
    resume_df = process_resumes(resume_df, job_description)
    save_results(resume_df)
