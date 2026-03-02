import os
import pandas as pd
import csv
import time
from datetime import datetime  # Import datetime module
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Load environment variables from .env file
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv()

# Get credentials from environment variables
username = "S1SSARAC"  # Forcing the correct username
password = os.getenv("PASSWORD")

# Get the current timestamp for this execution run
execution_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# The script no longer needs to read a CSV file
# # Define the path to the CSV file in the current directory
# file_path = os.path.join(os.path.dirname(__file__), "Download_QEC_Summary.csv")
# print(f"Looking for CSV file at: {file_path}")

# # Read the CSV file and extract unique values from the first column
# try:
#     df = pd.read_csv(file_path, delimiter=';', on_bad_lines='skip', engine='python')
#     first_column_name = df.columns[0]
#     values_list = df[first_column_name].drop_duplicates().tolist()
#     print(f"Successfully read {len(values_list)} unique values from {file_path}")
        
# except (pd.errors.ParserError, FileNotFoundError) as e:
#     print(f"Error reading CSV file: {e}")
#     values_list = []  # Set to empty list if there's an error

from selenium.webdriver.chrome.service import Service

# Set up Chrome options
options = webdriver.ChromeOptions()
options.add_argument('--headless') # Run in headless mode for automation
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36")

# Initialize the WebDriver
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(10)  # Set implicit wait

# Open the login page
driver.get("https://qect.mo360cp.i.mercedes-benz.com/qec-access/supp/pruefberichtVerwalten.qec")

# Login process
try:
    # Check if username and password are provided
    if not username or not password:
        raise ValueError("Username or password not found in .env file")

    print("Attempting to log in...")
    user_id_field = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.NAME, "userid"))
    )
    user_id_field.send_keys(username)
    user_id_field.send_keys(Keys.RETURN)
    
    password_field = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.NAME, "password"))
    )
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)
    print("Login successful!")

    # --- New logic: Scrape values directly from the page ---
    print("Waiting for report numbers to load on the main page...")
    try:
        # Wait up to 15 seconds for at least one link to appear before proceeding
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='/qec-access/supp/pruefberichtBearbeiten.qec?pruefberichtId=']"))
        )
        print("Report numbers loaded. Scraping...")
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        values_list = []
        # Find all 'a' tags whose 'href' contains the specific path for report editing
        for link in soup.find_all('a', href=lambda href: href and '/qec-access/supp/pruefberichtBearbeiten.qec?pruefberichtId=' in href):
            value = link.get_text(strip=True)
            if value.isdigit(): # Ensure we only add numbers
                values_list.append(value)

        if not values_list:
            print("Could not find any report numbers on the page after waiting. Exiting.")
        else:
            print(f"Successfully scraped {len(values_list)} unique values from the page.")

    except Exception as e:
        print(f"An error occurred while waiting for or scraping report numbers: {e}")
        values_list = [] # Ensure list is empty on error

    # Define the headers for the CSV file based on user request
    headers = [
        'Processed_Value', 'Processed', 'Ref. no.', 'Part number', 'Denotation', 
        'Accepted', 'ntf', 'Cust. at fault', 'Consent', 'Delayed', 'Log. delayed',
        'Execution_Timestamp'  # Add new timestamp header
    ]

    # Prepare to write to CSV
    csv_file_path = os.path.join(os.path.dirname(__file__), "table_data.csv")
    print(f"Will write output to: {csv_file_path}")

    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(headers)  # Write the new headers
        
        # Only proceed if values were found
        if values_list:
            print(f"Starting to process {len(values_list)} values...")
            for i, value in enumerate(values_list):
                try:
                    print(f"Processing value {i+1}/{len(values_list)}: {value}")
                    search_field = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.NAME, "pruefberichtsnummerDirect"))
                    )
                    search_field.clear()
                    search_field.send_keys(value)
                    
                    submit_button = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.NAME, "directJumpPruefbericht"))
                    )
                    submit_button.click()
                    
                    # Wait for the page to load by looking for a known element
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "scrollTableLight"))
                    )

                    # Parse the page source with BeautifulSoup
                    soup = BeautifulSoup(driver.page_source, 'html.parser')

                    # Find the header "Processed" and then its parent table
                    processed_header = soup.find('td', class_='scrollTableLight', string='Processed')
                    if processed_header:
                        data_table = processed_header.find_parent('table')
                        if data_table:
                            header_row = processed_header.find_parent('tr')
                            
                            # Get all header texts and their indices
                            header_cells = header_row.find_all('td')
                            header_texts = [cell.get_text(strip=True) for cell in header_cells]
                            
                            # Map desired headers to their index
                            column_indices = {header: i for i, header in enumerate(header_texts) if header in headers}

                            # Find all data rows
                            data_rows = header_row.find_next_siblings('tr')

                            for row in data_rows:
                                cells = row.find_all('td')
                                if len(cells) == len(header_texts):
                                    row_data = {
                                        'Processed_Value': value,
                                        'Execution_Timestamp': execution_timestamp  # Add timestamp to data
                                    }
                                    for header, index in column_indices.items():
                                        cell = cells[index]
                                        if header == 'Processed':
                                            img = cell.find('img')
                                            row_data[header] = os.path.basename(img['src']) if img and img.has_attr('src') else 'N/A'
                                        else:
                                            row_data[header] = cell.get_text(strip=True)
                                    
                                    # Write the extracted data in the correct order
                                    csv_writer.writerow([row_data.get(h, '') for h in headers])
                        else:
                            print(f"  > Could not find parent table for value {value}.")
                    else:
                        print(f"  > Could not find the 'Processed' header for value {value}.")

                    # Add a delay and navigate back to the search page
                    print("  > Pausing for 2 seconds...")
                    time.sleep(2)
                    print("  > Navigating back to the search page.")
                    driver.get("https://qect.mo360cp.i.mercedes-benz.com/qec-access/supp/pruefberichtVerwalten.qec")

                except Exception as e:
                    print(f"  > An error occurred while processing value {value}: {e}")
                    # Save a screenshot for debugging
                    screenshot_path = os.path.join(os.path.dirname(__file__), f"error_screenshot_{value}.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"  > Screenshot saved to {screenshot_path}")
        else:
            print("No values found to process. CSV file with headers has been created.")

finally:
    # The browser will remain open for inspection until the user presses Enter.
    print("Closing the browser.")
    driver.quit()
