import requests
import pandas as pd
import os
import sys 
import time
from bs4 import BeautifulSoup
import re
# import xlwings as xw

def download_excel_file(url, save_path):

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content of the page
        soup = BeautifulSoup(response.text, "html.parser")
    
        # Find all <a> elements that contain the text "Current English Pay Items List"
        link = soup.find("a", string=re.compile(r'Current English Pay Items List', re.IGNORECASE))
        if link:
            # Extract the URL from the link
            file_url = link.get("href")
            file_url = 'https://www.in.gov' + file_url

            # Download the file
            file_response = requests.get(file_url)
            
            # Check if the file download was successful
            if file_response.status_code == 200:
                # Save the file to local storage
                with open(save_path, "wb") as file:
                    file.write(file_response.content)
                    print("File downloaded successfully.")
                    return True
            else:
                print("Failed to download file:", file_response.status_code)
                time.sleep(2)
                return False
        else:
            print("No link found containing 'Current English Pay Items List'.")
    else:
        print("Failed to fetch page:", response.status_code)

def main():
    # Webpage from the Excel files to download
    url = "https://www.in.gov/indot/doing-business-with-indot/home/contracts/standards/indot-pay-items-listunit-price-summaries/"
    
    # Path to save the downloaded file
    if len(sys.argv) != 2:
        print("Usage: python download_excel.py <directory_path>")
        return
    
    directory_path = sys.argv[1]
    save_path = os.path.join(directory_path, 'downloaded_file.xlsx')
    # Download the file
    if download_excel_file(url, save_path) is True:
        print("File downloaded successfully.")

    else:
        print("Failed to download file.")

if __name__ == "__main__":
    main()
