import requests
import pandas as pd
import os
import time 
from bs4 import BeautifulSoup

def download_excel_file(url, save_path):
    response = requests.get(url)
    if response.status_code == 200:

        # Parse the HTML content of the page
        soup = BeautifulSoup(response.text, "html.parser")
        # Find the "Current" section in the parsed HTML
        current_section = soup.find("div", string="Current")
        # Find the first href with "Current" class
        link = current_section.find_next("a", href=True)
        if link:
            file_url = link.get("href")
            print(file_url)

        with open(save_path, 'wb') as file:
            file.write(response.content)
        return True
    else:
        print("Failed to download file.")
        time.sleep(2)
        return False

def open_excel_file(file_path):
    time.sleep(2)
    print("Attempting to open file at path:", file_path) 
    try:
        df = pd.read_excel(file_path)
        os.startfile(file_path)
        time.sleep(1)
    except FileNotFoundError:
        print("Error: File not found at", file_path)
        time.sleep(3)
    except Exception as e:
        print("Unknown error, please contact Chris Sexton @ cksexton@hntb.com", e)
        time.sleep(3)

def main():
    # Webpage from the Excel file to download
    url = "https://www.in.gov/indot/doing-business-with-indot/files"
    
    # Path to save the downloaded file
    desktop_path = os.path.join("C:\\", "Users", "cksexton", "Desktop", "print", "Indot Pay items")
    save_path = os.path.join(desktop_path, 'downloaded_file.xlsx')
    
    # Download the file
    if download_excel_file(url, save_path) is True:
        print("File downloaded successfully.")
        
        # Open the downloaded file
        open_excel_file(save_path)
    else:
        print("Failed to download file.")

if __name__ == "__main__":
    main()
