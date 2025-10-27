import openpyxl
import xlrd  # For handling .xls files
import csv
import requests
from bs4 import BeautifulSoup
import re
import sys
from datetime import datetime
import os
import random
from urllib.parse import quote_plus
import argparse

# Function to perform search with multiple engines
def search_realtor_info(query, search_engine="duckduckgo"):
    """
    Search for realtor information using different search engines
    """
    
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
def handle_xlsx(file_path, start_index, end_index=None):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    data_list = []
    for i, r in enumerate(sheet.iter_rows()):
        if i < start_index:
            continue
        if end_index and i >= end_index:
            break
        
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
    return data_list

# Function to handle xls files for phone data format
def handle_xls(file_path, start_index, end_index=None):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)

    data_list = []
    max_rows = sheet.nrows if not end_index else min(end_index, sheet.nrows)
    
    for i in range(start_index, max_rows):
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
def handle_csv(file_path, start_index, end_index=None):
    data_list = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)
        for i, row in enumerate(csvreader):
            if i < start_index:
                continue
            if end_index and i >= end_index:
                break
                
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
    return data_list

# Function to run the scraper
def run_scraper(file_path, start_index, end_index=None, primary_engine="duckduckgo", output_file="phone_email_output.csv"):
    try:
        print("Phone Email Scraper started...")
        print(f"Input file: {file_path}")
        print(f"Start index: {start_index}")
        print(f"End index: {end_index if end_index else 'End of file'}")
        print(f"Output file: {output_file}")

        start_index = int(start_index)
        if end_index:
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

        print(f"Loaded {len(data_list)} records to process")

        # Define search engines in order of preference (most reliable first)
        search_engines_order = [
            "yahoo",         # Reliable
            "searx",         # Privacy-focused
            "startpage",     # Google proxy
            "ask",           # Alternative
            "yandex",        # Russian search engine
        ]
        
        # Move primary engine to front if it's in the list
        if primary_engine in search_engines_order:
            search_engines_order.remove(primary_engine)
            search_engines_order.insert(0, primary_engine)
        
        # Check if output file exists to determine if we need to write headers
        file_exists = os.path.exists(output_file)
        
        # Open CSV file in append mode
        with open(output_file, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            
            # Write headers only if file doesn't exist
            if not file_exists:
                writer.writerow(["First Name", "Last Name", "Phone", "Full Name", "Email", "Source Link"])
            
            for idx, data in enumerate(data_list, 1):
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
                
                print(f"[{idx}/{len(data_list)}] Searching for: {full_name} - {phone_number}")
                
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
                
                # Write result to CSV immediately
                writer.writerow([
                    data["first_name"], 
                    data["last_name"], 
                    data["phone"], 
                    data["full_name"],
                    email, 
                    link
                ])
                
                # Flush to ensure data is written immediately
                file.flush()

        print(f"\nScraping completed. Results appended to {output_file}")
        print(f"Processed {len(data_list)} records")
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

def main():
    parser = argparse.ArgumentParser(description='Phone Email Scraper - CLI Version')
    parser.add_argument('--input-file', '-i', default='phone_data.xlsx',
                        help='Input file path (default: phone_data.xlsx)')
    parser.add_argument('--start-index', '-s', type=int, default=1,
                        help='Start index (default: 1)')
    parser.add_argument('--end-index', '-e', type=int, default=None,
                        help='End index (default: end of file)')
    parser.add_argument('--output-file', '-o', default='phone_email_output.csv',
                        help='Output CSV file (default: phone_email_output.csv)')
    parser.add_argument('--engine', default='duckduckgo',
                        choices=['duckduckgo', 'bing', 'google', 'yahoo', 'ask', 'yandex', 'ecosia', 'startpage', 'searx'],
                        help='Primary search engine (default: duckduckgo)')
    
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found!")
        sys.exit(1)
    
    # Run the scraper
    run_scraper(
        file_path=args.input_file,
        start_index=args.start_index,
        end_index=args.end_index,
        primary_engine=args.engine,
        output_file=args.output_file
    )

if __name__ == "__main__":
    main()


