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
import time
import random
from urllib.parse import quote_plus

# Redirecting print to a tkinter scrolledtext widget
class PrintLogger:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.insert(tk.END, text)
        self.textbox.see(tk.END)  # Auto-scroll to the bottom

    def flush(self):  # This is required for the flush method of file-like objects
        pass

# Function to perform search with multiple engines
def search_realtor_info(query, search_engine="duckduckgo"):
    """
    Search for realtor information using different search engines
    """
    # Removed delays for faster processing
    
    if search_engine == "duckduckgo":
        return search_duckduckgo(query)
    elif search_engine == "bing":
        return search_bing(query)
    elif search_engine == "google":
        return search_google(query)
    elif search_engine == "yahoo":
        return search_yahoo(query)
    elif search_engine == "ask":
        return search_ask(query)
    elif search_engine == "yandex":
        return search_yandex(query)
    elif search_engine == "ecosia":
        return search_ecosia(query)
    elif search_engine == "startpage":
        return search_startpage(query)
    elif search_engine == "searx":
        return search_searx(query)
    else:
        return None

def search_duckduckgo(query):
    """Search using DuckDuckGo (more bot-friendly)"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        # DuckDuckGo search URL
        search_url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
        
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"DuckDuckGo search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"DuckDuckGo search error: {e}")
        return None

def search_bing(query):
    """Search using Bing"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        # Bing search URL
        search_url = f"https://www.bing.com/search?q={quote_plus(query)}"
        
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Bing search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Bing search error: {e}")
        return None

def search_google(query):
    """Search using Google with enhanced anti-detection"""
    try:
        # Rotate between different user agents
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15"
        ]
        
        headers = {
            "User-Agent": random.choice(user_agents),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Cache-Control": "max-age=0",
        }
        
        # Google search URL with additional parameters
        search_url = f"https://www.google.com/search?q={quote_plus(query)}&num=10"
        
        response = requests.get(search_url, headers=headers, timeout=15)
        
        if response.status_code == 200:
            # Check if we got a challenge page
            if "Please click here if you are not redirected" in response.text:
                print("Google detected bot traffic, trying alternative search engine...")
                return None
            return response.text
        else:
            print(f"Google search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Google search error: {e}")
        return None

def search_yahoo(query):
    """Search using Yahoo"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        search_url = f"https://search.yahoo.com/search?p={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Yahoo search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Yahoo search error: {e}")
        return None

def search_ask(query):
    """Search using Ask.com"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        search_url = f"https://www.ask.com/web?q={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Ask search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Ask search error: {e}")
        return None

def search_yandex(query):
    """Search using Yandex"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        search_url = f"https://yandex.com/search/?text={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Yandex search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Yandex search error: {e}")
        return None

def search_ecosia(query):
    """Search using Ecosia"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        search_url = f"https://www.ecosia.org/search?q={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Ecosia search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Ecosia search error: {e}")
        return None

def search_startpage(query):
    """Search using Startpage (Google proxy)"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        search_url = f"https://www.startpage.com/sp/search?query={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Startpage search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Startpage search error: {e}")
        return None

def search_searx(query):
    """Search using Searx (privacy-focused)"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        
        # Using a public Searx instance
        search_url = f"https://searx.be/search?q={quote_plus(query)}"
        response = requests.get(search_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            return response.text
        else:
            print(f"Searx search failed with status code: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Searx search error: {e}")
        return None

def extract_emails_and_links(html_content, search_engine="duckduckgo"):
    """
    Extract emails and links from search results
    Works with different search engines
    """
    if not html_content:
        return '', ''
        
    soup = BeautifulSoup(html_content, 'html.parser')
    email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    email = ''
    link = ''
    
    # Save HTML for debugging (optional)
    with open("search_result_debug.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    
    if search_engine == "duckduckgo":
        # DuckDuckGo specific selectors
        results = soup.find_all('div', class_='result')
        for result in results[:5]:  # Check first 5 results
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                # Find the main link in this result
                link_elem = result.find('a', class_='result__a')
                if link_elem and link_elem.get('href'):
                    link = link_elem['href']
                break
                
    elif search_engine == "bing":
        # Bing specific selectors
        results = soup.find_all('li', class_='b_algo')
        for result in results[:5]:  # Check first 5 results
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                # Find the main link in this result
                link_elem = result.find('h2').find('a') if result.find('h2') else None
                if link_elem and link_elem.get('href'):
                    link = link_elem['href']
                break
                
    elif search_engine == "google":
        # Google specific selectors (fallback)
        results = soup.find_all('div', class_='g')
        for result in results[:5]:  # Check first 5 results
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                # Find the main link in this result
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "yahoo":
        # Yahoo specific selectors
        results = soup.find_all('div', class_='Sr')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "ask":
        # Ask.com specific selectors
        results = soup.find_all('div', class_='PartialSearchResults-item')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "yandex":
        # Yandex specific selectors
        results = soup.find_all('li', class_='serp-item')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "ecosia":
        # Ecosia specific selectors
        results = soup.find_all('article', class_='result')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "startpage":
        # Startpage specific selectors
        results = soup.find_all('div', class_='w-gl__result')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
                
    elif search_engine == "searx":
        # Searx specific selectors
        results = soup.find_all('div', class_='result')
        for result in results[:5]:
            text = result.get_text()
            email_match = re.search(email_regex, text)
            if email_match:
                email = email_match.group(0)
                link_elem = result.find('a', href=True)
                if link_elem:
                    link = link_elem['href']
                break
    
    # If no email found in structured results, search entire page
    if not email:
        page_text = soup.get_text()
        email_match = re.search(email_regex, page_text)
        if email_match:
            email = email_match.group(0)
    
    return email, link

# Function to handle xlsx files for phone data format (First Name, Last Name, Phone)
def handle_xlsx(file_path, start_index, end_index):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    data_list = []
    for i, r in enumerate(sheet.iter_rows()):
        if start_index <= i < end_index:
            # Check if we have valid data in the first three columns
            if r[0].value and r[1].value and r[2].value:
                first_name = str(r[0].value).strip()
                last_name = str(r[1].value).strip()
                phone = str(r[2].value).strip()
                
                # Skip header row or empty rows
                if first_name.lower() not in ['first name', 'firstname', ''] and len(first_name) > 1:
                    data_list.append({
                        "first_name": first_name,
                        "last_name": last_name,
                        "phone": phone,
                        "full_name": f"{first_name} {last_name}"
                    })
        if i >= end_index:
            break
    return data_list

# Function to handle xls files for phone data format
def handle_xls(file_path, start_index, end_index):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)

    data_list = []
    for i in range(start_index, end_index):
        row = sheet.row(i)
        if len(row) >= 3 and row[0].value and row[1].value and row[2].value:
            first_name = str(row[0].value).strip()
            last_name = str(row[1].value).strip()
            phone = str(row[2].value).strip()
            
            # Skip header row or empty rows
            if first_name.lower() not in ['first name', 'firstname', ''] and len(first_name) > 1:
                data_list.append({
                    "first_name": first_name,
                    "last_name": last_name,
                    "phone": phone,
                    "full_name": f"{first_name} {last_name}"
                })
    return data_list

# Function to handle csv files for phone data format
def handle_csv(file_path, start_index, end_index):
    data_list = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)
        for i, row in enumerate(csvreader):
            if start_index <= i < end_index:
                if len(row) >= 3 and row[0] and row[1] and row[2]:
                    first_name = str(row[0]).strip()
                    last_name = str(row[1]).strip()
                    phone = str(row[2]).strip()
                    
                    # Skip header row or empty rows
                    if first_name.lower() not in ['first name', 'firstname', ''] and len(first_name) > 1:
                        data_list.append({
                            "first_name": first_name,
                            "last_name": last_name,
                            "phone": phone,
                            "full_name": f"{first_name} {last_name}"
                        })
            if i >= end_index:
                break
    return data_list

# Function to run the scraper
def run_scraper(file_path, start_index, end_index, output_widget, run_button, primary_engine="duckduckgo"):
    try:
        print("Phone Email Scraper started...")  # Notify that scraper has started

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

        # Define search engines in order of preference (most reliable first)
        search_engines_order = [
            "yahoo",         # Reliable
            "searx",         # Privacy-focused
            # "duckduckgo",    # Most bot-friendly
            # "bing",          # Good alternative
            # "ecosia",        # Privacy-focused
            "startpage",     # Google proxy
            "ask",           # Alternative
            "yandex",        # Russian search engine
            # "google"         # Last resort (most likely to block)
        ]
        
        # Move primary engine to front if it's in the list
        if primary_engine in search_engines_order:
            search_engines_order.remove(primary_engine)
            search_engines_order.insert(0, primary_engine)
        
        for data in data_list:
            first_name = data["first_name"]
            last_name = data["last_name"]
            phone_number = data["phone"]
            full_name = data["full_name"]
            
            # Use one optimized query per search engine
            query = f"Email for realtor {first_name} {last_name}, {phone_number}"
            
            email = ''
            link = ''
            successful_search = False
            engines_tried = 0
            
            print(f"Searching for: {full_name} - {phone_number}")
            
            # Try search engines in order until we find an email
            for search_engine in search_engines_order:
                if successful_search:
                    break
                    
                engines_tried += 1
                print(f"  Trying search engine {engines_tried}/{len(search_engines_order)}: {search_engine}")
                
                try:
                    search_results_html = search_realtor_info(query, search_engine)
                    
                    if search_results_html:
                        email, link = extract_emails_and_links(search_results_html, search_engine)
                        if email:
                            print(f"    ✓ Found email: {email} (via {search_engine})")
                            successful_search = True
                        else:
                            print(f"    - No email found in {search_engine} results")
                    else:
                        print(f"    - {search_engine} search failed")
                        
                except Exception as e:
                    print(f"    - Error with {search_engine}: {e}")
                    continue
            
            if not successful_search:
                print(f"  ✗ No email found for {full_name} after trying {engines_tried} search engines")
            
            emails_list.append({
                "first_name": first_name,
                "last_name": last_name,
                "phone": phone_number,
                "full_name": full_name,
                "email": email,
                "source_link": link
            })

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder = "phone_email_output"
        os.makedirs(output_folder, exist_ok=True)
        output_filename = os.path.join(output_folder, f"{timestamp}.csv")

        # Save the results to CSV file
        with open(output_filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["First Name", "Last Name", "Phone", "Full Name", "Email", "Source Link"])
            for data in emails_list:
                writer.writerow([
                    data["first_name"], 
                    data["last_name"], 
                    data["phone"], 
                    data["full_name"],
                    data["email"], 
                    data["source_link"]
                ])

        print(f"Scraping completed. Results saved in {output_filename}")
        print(f"Processed {len(emails_list)} records")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Re-enable the run button after completion
        run_button.config(text="Run Scraper", state=tk.NORMAL)

# Function to start scraper in a background thread
def start_scraper_in_thread(file_path, start_index, end_index, output_widget, run_button, primary_engine="duckduckgo"):
    run_button.config(text="Running...", state=tk.DISABLED)  # Disable the button and change text
    threading.Thread(target=run_scraper, args=(file_path, start_index, end_index, output_widget, run_button, primary_engine), daemon=True).start()

# Create the main window
def create_gui():
    window = tk.Tk()
    window.title("Phone Data Email Scraper")

    # Excel file path variable
    file_path = tk.StringVar()

    # Label and file dialog
    label = tk.Label(window, text="Upload Excel/CSV File (First Name, Last Name, Phone format):")
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
    start_label = tk.Label(window, text="Start Index (default 1 to skip header):")
    start_label.pack(pady=5)

    start_entry = tk.Entry(window)
    start_entry.insert(0, "1")  # Default start index is 1 (skip header row)
    start_entry.pack()

    # Label and Entry for end index
    end_label = tk.Label(window, text="End Index:")
    end_label.pack(pady=5)

    end_entry = tk.Entry(window)
    end_entry.insert(0, "10")  # Default value is 10 for testing
    end_entry.pack()

    # Search engine selection
    engine_label = tk.Label(window, text="Primary Search Engine:")
    engine_label.pack(pady=5)

    engine_var = tk.StringVar(value="duckduckgo")
    engine_frame = tk.Frame(window)
    engine_frame.pack()

    engines = [
        ("Yahoo", "yahoo"),
        ("Searx", "searx"),
        ("DuckDuckGo (Recommended)", "duckduckgo"), 
        ("Bing", "bing"), 
        ("Ecosia", "ecosia"),
        ("Startpage", "startpage"),
        ("Ask.com", "ask"),
        ("Yandex", "yandex"),
        ("Google (Last Resort)", "google")
    ]
    
    for text, value in engines:
        rb = tk.Radiobutton(engine_frame, text=text, variable=engine_var, value=value)
        rb.pack(anchor=tk.W)

    # Output text box for print statements
    output_text = scrolledtext.ScrolledText(window, height=15, width=80)
    output_text.pack(pady=10)

    # Redirect print statements to the ScrolledText widget
    sys.stdout = PrintLogger(output_text)

    # Run scraper button
    run_button = tk.Button(window, text="Run Scraper", command=lambda: start_scraper_in_thread(file_path.get(), start_entry.get(), end_entry.get(), output_text, run_button, engine_var.get()))
    run_button.pack(pady=10)

    # Start the GUI loop
    window.mainloop()

if __name__ == "__main__":
    create_gui()
