import pandas as pd
import docx.oxml.ns as ns
from docx import Document
import re
from docx.shared import RGBColor
import docx.oxml as oxml
import openai
from datetime import datetime

API_KEY_FILE = 'apikey.txt'

# 1. read csv file
def read_csv_file(file_path):
    try:
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)
        # Keep only the specified fields (columns)
        fields = ["Title", "Funder", "Deadline", "Amount", "Eligibility", "Abstract", "More Information"]
        filtered_df = df[fields]
        # remove html tag in each cell
        filtered_df = filtered_df.map(remove_html_tag)
        # Return the DataFrame
        return filtered_df
    except FileNotFoundError:
        print("File not found!")
        return None
    except KeyError:
        print("Some specified fields do not exist in the CSV file!")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def remove_html_tag(value):
    if not isinstance(value, str):
        value = str(value)
    value = re.sub(r"<[^>]*>", "", value)
    return value

# 2. convert csv file to word file and format word file
def format_word_file(data_frame, output_file_path):
    get_api_key()

    # Create a new Word document for the formatted content
    formatted_doc = Document()

    # Iterate over each row in the DataFrame, skipping the first row
    for _, row in data_frame.iterrows():
        p = formatted_doc.add_paragraph()
        
        # title
        title_text = f"{row['Funder']} | {row['Title']}"
        if row['More Information']:
            hyperlink_run = add_hyperlink(p, row['More Information'], title_text)
            hyperlink_run.font.color.rgb = RGBColor(0, 0, 255)
            hyperlink_run.bold = True
            p.add_run("\n")
        else:
            title_run = p.add_run(title_text)
            title_run.bold = True
            p.add_run("\n")
        
        # Deadline
        if row['Deadline']:
            deadline_txt = row['Deadline']
            bold_run = p.add_run(f"Due Date: ")
            bold_run.bold = True
            
            # Use current date
            current_date = datetime.now()
            
            # Regular expression to find dates in the format "DD MMM YYYY"
            date_pattern = r"\d{2} \w{3} \d{4}"
            
            # Find all dates
            dates = re.findall(date_pattern, deadline_txt)
            
            # Convert found dates to datetime objects and filter out past dates
            future_dates = [datetime.strptime(date, "%d %b %Y") for date in dates if datetime.strptime(date, "%d %b %Y") > current_date]

            # Find the closest future date
            closest_future_date = min(future_dates, key=lambda x: (x - current_date))

            # Find the line in the text containing the closest future date
            closest_date_line = [line for line in deadline_txt.split('\n') if closest_future_date.strftime("%d %b %Y") in line]
            closest_date = closest_date_line[0] if closest_date_line else "No upcoming date found"
            p.add_run(f"{closest_date}\n")
            
        # Amount
        if row['Amount']:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # Specify the chat model
                messages=[
                    {"role": "system", "content": "You are a helpful assistant who's good at summarization."},
                    {"role": "user", "content": f"Summarize the following text by extracting award amount in"
                                                f"USD. Is amount upper exists, just use that number is enough."
                                                f"Do not include any notes or explanations:\n\n{row['Amount']}"}
                ],
                max_tokens=150,  # Set the maximum length for the summary
                temperature=0.7  # Adjusts randomness in the response. Lower is more deterministic.
            )
            # The response format is different for chat completions
            summary = response['choices'][0]['message']['content'].strip()
            bold_run = p.add_run("Award Amount: ")
            bold_run.bold = True
            p.add_run(f"{summary}\n")

        # Eligibility
        if row['Eligibility']:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # Specify the chat model
                messages=[
                    {"role": "system", "content": "You are a helpful assistant who's good at summarization."},
                    {"role": "user", "content": f"Summarize the following text by extracting which level "
                                                f"of faculty is eligible. If the level is not mentioned, "
                                                f"simply return Any level faculty:\n\n{row['Eligibility']}"
                                                f"also include information if this requires MD or PhD if the information is available"}
                ],
                max_tokens=150,  # Set the maximum length for the summary
                temperature=0.7  # Adjusts randomness in the response. Lower is more deterministic.
            )
            # The response format is different for chat completions
            summary = response['choices'][0]['message']['content'].strip()
            bold_run = p.add_run("Eligibility: ")
            bold_run.bold = True
            p.add_run(f"{summary}\n")
        
        
        # Abstract
        if row['Abstract']:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # Specify the chat model
                messages=[
                    {"role": "system", "content": "You are a helpful assistant who's good at summarization."},
                    {"role": "user", "content": f"Summarize the following text in a concise way:\n\n{row['Abstract']}"}
                ],
                max_tokens=150,  # Set the maximum length for the summary
                temperature=0.7  # Adjusts randomness in the response. Lower is more deterministic.
            )
            # The response format is different for chat completions
            summary = response['choices'][0]['message']['content'].strip()
            bold_run = p.add_run("Program Goal: ")
            bold_run.bold = True
            p.add_run(f"{summary}\n")

    # Save the formatted Word document
    formatted_doc.save(output_file_path)

# Helper functions:
# Add hyperlink
def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)

    hyperlink = oxml.OxmlElement('w:hyperlink')
    hyperlink.set(oxml.ns.qn('r:id'), r_id)

    new_run = oxml.OxmlElement('w:r')
    rPr = oxml.OxmlElement('w:rPr')
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    r = paragraph.add_run()
    r._r.append(hyperlink)

    return r

def get_api_key():
    with open(API_KEY_FILE, 'r') as file:
        openai.api_key = file.read().strip()

def unify_line_endings(file_path):
    with open(file_path, 'r', newline=None) as file:
        content = file.read().replace('\r\n', '\n').replace('\r', '\n')

    with open(file_path, 'w', newline='\n') as file:
        file.write(content)

if __name__ == "__main__":
    file_path = "sample_data/opps_export2.csv"
    formatted_word_file_path = "output_word/formattedOutput.docx"
    unify_line_endings(file_path)
    # 1. read csv file
    data_frame = read_csv_file(file_path)
    if data_frame is not None:
        # 2. convert csv file to word file and format word file
        format_word_file(data_frame, formatted_word_file_path)
    