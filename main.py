import os
import requests
from bs4 import BeautifulSoup
from datetime import date, timedelta
import time

def download_eppo_oil_prices(start_date, end_date, output_dir):
    url = 'https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx'
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    session = requests.Session()
    session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

    # Get initial viewstate
    print("Fetching initial page to get ViewState...")
    try:
        response = session.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch initial page: {e}")
        return

    soup = BeautifulSoup(response.text, 'html.parser')
    viewstate = soup.find(id='__VIEWSTATE')['value']
    viewstategen = soup.find(id='__VIEWSTATEGENERATOR')['value']
    eventval = soup.find(id='__EVENTVALIDATION')['value']

    current_date = start_date
    while current_date <= end_date:
        date_str = current_date.strftime('%d/%m/%Y')
        filename_date_str = current_date.strftime('%Y%m%d')
        filename = f'EPPO_RetailOilPrice_on_{filename_date_str}.xls'
        filepath = os.path.join(output_dir, filename)

        if os.path.exists(filepath):
            print(f"File {filename} already exists. Skipping...")
            current_date += timedelta(days=1)
            continue

        print(f"Downloading data for {date_str}...")
        data = {
            '__VIEWSTATE': viewstate,
            '__VIEWSTATEGENERATOR': viewstategen,
            '__EVENTVALIDATION': eventval,
            'TbxToDate': date_str,
            'BtnGenerate.x': '10',
            'BtnGenerate.y': '10'
        }

        try:
            res = session.post(url, data=data)
            res.raise_for_status()
            
            # Check if we actually got a file download or just the page reloading (e.g. no data)
            if 'application/vnd.ms-excel' in res.headers.get('Content-Type', '') or res.headers.get('Content-Disposition'):
                with open(filepath, 'wb') as f:
                    f.write(res.content)
                print(f"Successfully downloaded {filename}")
            else:
                print(f"Warning: Did not receive an Excel file for {date_str}. (Maybe no data for this date?)")
                # Even if we didn't get an excel file, we might want to update the viewstate for the next request 
                # from this response, but usually for this site, reusing the same viewstate works for multiple date queries.
        
        except requests.exceptions.RequestException as e:
            print(f"Error downloading for {date_str}: {e}")

        # Add a short delay to be polite
        time.sleep(1)
        current_date += timedelta(days=1)

if __name__ == "__main__":
    start_date = date(2015, 1, 1)
    # Get current date or you can set a specific end date
    end_date = date(2023, 12, 31)
    output_directory = r'D:\Project\Thai-Oil-Price\Data'
    
    print(f"Starting download from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    download_eppo_oil_prices(start_date, end_date, output_directory)
    print("All tasks completed.")
