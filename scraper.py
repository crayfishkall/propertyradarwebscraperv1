import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import logging
import re

# Set up logging
logging.basicConfig(level=logging.INFO)

# Prompt user for property URL
property_url = input("Enter the Property URL: ")

# Set up the Chrome WebDriver using WebDriver Manager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the property page
driver.get(property_url)

# Wait for the page to load fully before proceeding
logging.info("Waiting for the page to load...")
WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "x-panel-bodyWrap")))

# Wait for user to log in manually and press Enter to continue
input("Please log in to the website and press Enter here when you're logged in...")

# Once the user presses Enter, the script continues
logging.info("Login successful, continuing script...")

# Function to extract label-value pairs for each tab (Property, Contacts, etc.)
def extract_label_value(soup, label):
    try:
        label_element = soup.find('label', string=re.compile(label, re.IGNORECASE))
        if label_element:
            parent_div = label_element.find_parent('div')
            if parent_div:
                value_element = parent_div.find_all('label')[1]
                value = value_element.text.strip() if value_element else 'Not Found'
                return value
        return 'Not Found'
    except Exception as e:
        logging.error(f"Error extracting {label}: {e}")
        return 'Not Found'

# Function to scrape each tab (Contacts, Property, etc.)
def scrape_tab(tab_name, tab_xpath):
    # Navigate to the tab
    logging.info(f"Navigating to the '{tab_name}' tab...")
    tab = driver.find_element(By.XPATH, tab_xpath)
    tab.click()
    time.sleep(3)  # Allow time for the page to load

    # Get the page source after loading
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Define the data to extract for each tab (this can be customized for each tab)
    if tab_name == "Property":
        data = {'Location': extract_label_value(soup, 'Location'),
            'County': extract_label_value(soup, 'County'),
            'Assessor Parcel Number': extract_label_value(soup, 'Assessor Parcel Number'),
            'Radar ID': extract_label_value(soup, 'Radar ID'),
            'Lat/Lon': extract_label_value(soup, 'Lat/Lon'),
            'Subdivision': extract_label_value(soup, 'Subdivision'),
            'Site Congressional District': extract_label_value(soup, 'Congressional District'),
            'School Tax District': extract_label_value(soup, 'School Tax District'),
            'Census Tract': extract_label_value(soup, 'Census Tract'),
            'Census Block': extract_label_value(soup, 'Census Block'),
            'Carrier Route': extract_label_value(soup, 'Carrier Route'),
            'Tax Rate Area': extract_label_value(soup, 'Tax Rate Area'),
            'Legal Book/Page/Block/Lot': extract_label_value(soup, 'Legal Book/Page/Block/Lot'),
            'Legal Description': extract_label_value(soup, 'Legal Description'),
            'Lot SqFt': extract_label_value(soup, 'Lot SqFt'),
            'Lot Acres': extract_label_value(soup, 'Lot Acres'),
            'Zoning': extract_label_value(soup, 'Zoning'),
            'View Type': extract_label_value(soup, 'View Type'),
            'Flood Zone Code': extract_label_value(soup, 'Flood Zone Code'),
            'Flood Risk': extract_label_value(soup, 'Flood Risk'),
            'FEMA Map Date': extract_label_value(soup, 'FEMA Map Date'),
            'Year Built': extract_label_value(soup, 'Year Built'),
            'Square Feet': extract_label_value(soup, 'Square Feet'),
            'Beds': extract_label_value(soup, 'Beds'),
            'Baths': extract_label_value(soup, 'Baths'),
            'Units': extract_label_value(soup, 'Units'),
            'Stories': extract_label_value(soup, 'Stories'),
            'Rooms': extract_label_value(soup, 'Rooms'),
            'Pool': extract_label_value(soup, 'Pool'),
            'Fireplace': extract_label_value(soup, 'Fireplace'),
            'Air Conditioning': extract_label_value(soup, 'Air Conditioning'),
            'Heating': extract_label_value(soup, 'Heating'),
            'Improvement Condition': extract_label_value(soup, 'Improvement Condition'),
            'Building Quality': extract_label_value(soup, 'Building Quality'),
            'Assessed Value': extract_label_value(soup, 'Assessed Value'),
            'Assessed Land Value': extract_label_value(soup, 'Assessed Land Value'),
            'Assessed Improvements': extract_label_value(soup, 'Assessed Improvements'),
            'Annual Taxes': extract_label_value(soup, 'Annual Taxes'),
            'Tax Payment 1 Amount/Status': extract_label_value(soup, 'Tax Payment 1 Amount/Status'),
            'Tax Payment 2 Amount/Status': extract_label_value(soup, 'Tax Payment 2 Amount/Status'),
            'Taxpayer': extract_label_value(soup, 'Taxpayer'),
            'Mail Vacant': extract_label_value(soup, 'Mail Vacant'),
            'Homeowner Tax Exemption': extract_label_value(soup, 'Homeowner Tax Exemption'),
        
        }

    elif tab_name == "Contacts":
        data = {
            'Contact Name': extract_label_value(soup, 'Contact Name'),
            'Phone Number': extract_label_value(soup, 'Phone Number'),
            'Email': extract_label_value(soup, 'Email'),
            'Mailing Address': extract_label_value(soup, 'Mailing Address'),
            'Primary Contact': extract_label_value(soup, 'Primary Contact'),
            'Ownership Role': extract_label_value(soup, 'Ownership Role'),
            'Gender': extract_label_value(soup, 'Gender'),
            'Age': extract_label_value(soup, 'Age'),
        }

    elif tab_name == "Value & Equity":
        data = {
             'Estimated Value': extract_label_value(soup, 'Estimated Value'),
            'Assessed Value': extract_label_value(soup, 'Assessed Value'),
            'Estimated Open Loans Balance': extract_label_value(soup, 'Estimated Open Loans Balance'),
            'Estimated Equity': extract_label_value(soup, 'Estimated Equity'),
            'Purchase Date': extract_label_value(soup, 'Purchase Date'),
            'Purchase Amount': extract_label_value(soup, 'Purchase Amount'),
            'Market Value': extract_label_value(soup, 'Market Value'),
            'Rent Break Even': extract_label_value(soup, 'Rent Break Even'),
            'Market Rent': extract_label_value(soup, 'Market Rent'),
            'HUD Fair Market Rent': extract_label_value(soup, 'HUD Fair Market Rent'),
        }

    elif tab_name == "Transactions":
        data = {
            'Transaction Type': extract_label_value(soup, 'Transaction Type'),
            'Sale Date': extract_label_value(soup, 'Sale Date'),
            'Sale Price': extract_label_value(soup, 'Sale Price'),
            'Buyer': extract_label_value(soup, 'Buyer'),
            'Seller': extract_label_value(soup, 'Seller'),
        }

    elif tab_name == "Listings":
        data = {
            'Listing Price': extract_label_value(soup, 'Listing Price'),
            'Listed Date': extract_label_value(soup, 'Listed Date'),
            'Price History': extract_label_value(soup, 'Price History'),
            'Agent': extract_label_value(soup, 'Agent'),
        }

    
    # Return the extracted data
    return data
                         
# Function to create the DataFrame for a given tab
def create_tab_dataframe(data):
    # Convert the data into a pandas DataFrame (labels as rows, values in columns)
    return pd.DataFrame(list(data.items()), columns=['Label', 'Value'])

# Extract the data for each tab
tabs = {
    'Property': '//*[text()="Property"]',  # Adjust XPaths as needed for the tabs
    'Contacts': '//*[text()="Contacts"]',
    'Value & Equity': '//*[text()="Value & Equity"]',
    'Transactions': '//*[text()="Transactions"]',
    'Listings': '//*[text()="Listings"]',
}

tab_data = {}
for tab_name, tab_xpath in tabs.items():
    tab_data[tab_name] = scrape_tab(tab_name, tab_xpath)

# Assuming you have the address from your URL and mailing address from your scraping function
address = property_url.split("/")[-1]

# Create Excel file
excel_filename = f"{address}_property_data.xlsx"

# Create Excel file with multiple tabs
with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
    # Write each tab's data to a separate sheet
    for tab_name, data in tab_data.items():
        # Convert to DataFrame and save to respective sheet
        df = create_tab_dataframe(data)
        df.to_excel(writer, sheet_name=tab_name, index=False)

        # Optionally, adjust column widths for better readability
        worksheet = writer.sheets[tab_name]
        worksheet.set_column(0, 0, 30)  # Set column 0 width for labels
        worksheet.set_column(1, 1, 50)  # Set column 1 width for values

        logging.info(f"Data for '{tab_name}' tab saved.")

logging.info(f"Data saved to Excel file: {excel_filename}")

# Close the browser after extracting data
driver.quit()

#pip install selenium webdriver-manager beautifulsoup4 pandas logging
