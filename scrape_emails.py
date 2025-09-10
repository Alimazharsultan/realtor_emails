import openpyxl
import xlrd  # For handling .xls files
import csv
import requests
from bs4 import BeautifulSoup
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter import messagebox
import sys
from datetime import datetime
import os
import threading

# Redirecting print to a tkinter scrolledtext widget
class PrintLogger:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.insert(tk.END, text)
        self.textbox.see(tk.END)  # Auto-scroll to the bottom

    def flush(self):  # This is required for the flush method of file-like objects
        pass

# Function to perform Google search for a realtor's name and phone number
def search_google(query):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

    # Construct Google search URL
    google_search_url = f"https://www.google.com/search?q={query}"

    # Make the request to Google search
    response = requests.get(google_search_url, headers=headers)

    # Check if the request was successful
    if response.status_code == 200:
        return response.text
    else:
        print(f"Failed to retrieve Google search results. Status code: {response.status_code}")
        return None

def extract_emails_and_links(html_content):
    soup = BeautifulSoup(html_content, 'lxml')
    
    div_ids = [f"rso div:nth-of-type({i})" for i in range(1, 11)]
    email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    email = ''
    link = ''
    
    for div_id in div_ids:
        div = soup.select_one(f'#{div_id}')
        if div:
            email_match = re.search(email_regex, div.get_text())
            if email_match:
                email = email_match.group(0)
                
                link_tag = div.find('a', href=True)
                if link_tag:
                    link = link_tag['href']
    
    return email, link

# Function to handle xlsx files
def handle_xlsx(file_path, start_index, end_index):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    data_list = []
    for i, r in enumerate(sheet.iter_rows()):
        if start_index <= i < end_index:
            if r[0].value and len(r[0].value) > 1:
                data_list.append({"name": r[0].value, "number": r[4].value})
        if i >= end_index:
            break
    return data_list

# Function to handle xls files
def handle_xls(file_path, start_index, end_index):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)

    data_list = []
    for i in range(start_index, end_index):
        row = sheet.row(i)
        if len(row) > 0 and row[0].value and len(str(row[0].value)) > 1:
            data_list.append({"name": row[0].value, "number": row[4].value})
    return data_list

# Function to handle csv files
def handle_csv(file_path, start_index, end_index):
    data_list = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)
        for i, row in enumerate(csvreader):
            if start_index <= i < end_index:
                if row[0] and len(row[0]) > 1:
                    data_list.append({"name": row[0], "number": row[4]})
            if i >= end_index:
                break
    return data_list

# Function to run the scraper
def run_scraper(file_path, start_index, end_index, output_widget, run_button):
    try:
        print("Scraper started...")  # Notify that scraper has started

        start_index = int(start_index)
        end_index = int(end_index)

        file_extension = os.path.splitext(file_path)[1].lower()

        # Load data based on file type
        if file_extension == ".xlsx":
            data_list = handle_xlsx(file_path, start_index, end_index)
        elif file_extension == ".xls":
            data_list = handle_xls(file_path, start_index, end_index)
        elif file_extension == ".csv":
            data_list = handle_csv(file_path, start_index, end_index)
        else:
            print(f"Unsupported file type: {file_extension}")
            return

        emails_list = []

        for data in data_list:
            realtor_name = data["name"]
            phone_number = data["number"]
            query = f"Realtor Email for: {realtor_name} {phone_number}"

            print(f"Searching for: {realtor_name} {phone_number}")
            search_results_html = search_google(query)
            if search_results_html:
                email, link = extract_emails_and_links(search_results_html)
                emails_list.append({"name": realtor_name, "number": phone_number, "email": email, "source_link": link})

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder = "realtor_email_output"
        os.makedirs(output_folder, exist_ok=True)
        output_filename = os.path.join(output_folder, f"{timestamp}.csv")

        
        # Save the results to output.csv
        with open(output_filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["name", "number", "email", "page"])
            for data in emails_list:
                writer.writerow([data["name"], data["number"], data["email"], data["source_link"]])

        print("Scraping completed. Results saved in realtor_email_output folder")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Re-enable the run button after completion
        run_button.config(text="Run Scraper", state=tk.NORMAL)

# Function to start scraper in a background thread
def start_scraper_in_thread(file_path, start_index, end_index, output_widget, run_button):
    run_button.config(text="Running...", state=tk.DISABLED)  # Disable the button and change text
    threading.Thread(target=run_scraper, args=(file_path, start_index, end_index, output_widget, run_button), daemon=True).start()

# Create the main window
def create_gui():
    window = tk.Tk()
    window.title("Realtor Email Scraper")

    # Excel file path variable
    file_path = tk.StringVar()

    # Label and file dialog
    label = tk.Label(window, text="Upload Excel/CSV File:")
    label.pack(pady=10)

    # Button to upload file
    def upload_file():
        file = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if file:
            file_path.set(file)
            file_label.config(text=file)

    file_button = tk.Button(window, text="Browse", command=upload_file)
    file_button.pack()

    file_label = tk.Label(window, text="")
    file_label.pack(pady=5)

    # Label and Entry for start index
    start_label = tk.Label(window, text="Start Index (default 3):")
    start_label.pack(pady=5)

    start_entry = tk.Entry(window)
    start_entry.insert(0, "3")  # Default start index is 3 (skip first 2 rows)
    start_entry.pack()

    # Label and Entry for end index
    end_label = tk.Label(window, text="End Index:")
    end_label.pack(pady=5)

    end_entry = tk.Entry(window)
    end_entry.insert(0, "10")  # Default value is 10 for testing
    end_entry.pack()

    # Output text box for print statements
    output_text = scrolledtext.ScrolledText(window, height=15, width=80)
    output_text.pack(pady=10)

    # Redirect print statements to the ScrolledText widget
    sys.stdout = PrintLogger(output_text)

    # Run scraper button
    run_button = tk.Button(window, text="Run Scraper", command=lambda: start_scraper_in_thread(file_path.get(), start_entry.get(), end_entry.get(), output_text, run_button))
    run_button.pack(pady=10)

    # Start the GUI loop
    window.mainloop()

if __name__ == "__main__":
    create_gui()
