from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
from bs4 import BeautifulSoup
import pandas as pd

def fetch_and_save_html(url, filename):
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (optional)
    chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration (optional)
    chrome_options.add_argument("--no-sandbox")  # Bypass OS security model (optional)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")

    # Set up the Chrome driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Open the URL in the browser
    driver.get(url)
    
    # Wait for the page to load
    time.sleep(5)  # Adjust the sleep time if necessary
    
    # Get the page source (HTML content)
    html_content = driver.page_source
    
    # Save the HTML content to a local file
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    print(f"HTML content saved to {filename}")
    
    # Close the browser
    driver.quit()

def extract_company_name(soup):
    
    # Find the span element with the data-tracking-name attribute
    company_name_span = soup.find('span', {'data-tracking-name': 'Doing Business As:'})
    
    # Extract and return the text content of the span element
    if company_name_span:
        return company_name_span.text.strip()
    else:
        return None

def extract_company_website(soup):
    
    # Find the span element with the name attribute set to "company_website"
    company_website_span = soup.find('span', {'name': 'company_website'})
    
    # Extract and return the href attribute of the a tag within the span element
    if company_website_span:
        a_tag = company_website_span.find('a')
        if a_tag and 'href' in a_tag.attrs:
            return a_tag['href']
    return None

def extract_key_principal(soup):
    
    # Find the span element with the name attribute set to "key_principal"
    key_principal_span = soup.find('span', {'name': 'key_principal'})
    
    # Extract and return the text content of the first span element within key_principal_span
    if key_principal_span:
        inner_span = key_principal_span.find('span')
        if inner_span:
            return inner_span.contents[0].strip()
    return None

def extract_company_address(soup):
    
    # Find the span element with the name attribute set to "company_address"
    company_address_span = soup.find('span', {'name': 'company_address'})
    
    # Extract and return the text content of the a tag within the span element
    if company_address_span:
        a_tag = company_address_span.find('a')
        if a_tag:
            return a_tag.text.strip()
    return None

def extract_maps_location(soup):
    
    # Find the span element with the name attribute set to "company_address"
    company_address_span = soup.find('span', {'name': 'company_address'})
    
    # Extract and return the href attribute of the a tag within the span element
    if company_address_span:
        a_tag = company_address_span.find('a')
        if a_tag and 'href' in a_tag.attrs:
            return a_tag['href']
    return None

def extract_industry_list(soup):
    
    # Find the span element with the name attribute set to "industry_links"
    industry_links_span = soup.find('span', {'name': 'industry_links'})
    
    # Extract and return the text content of each a tag and inner span within the span element
    industries = []
    if industry_links_span:
        for span in industry_links_span.find_all('span'):
            a_tag = span.find('a')
            if a_tag:
                industries.append(a_tag.text.strip())
            else:
                industries.append(span.text.strip())
    return industries

def extract_other_industries_list(soup):
   
    # Find the span element with the name attribute set to "other_industries_links"
    other_industries_span = soup.find('span', {'name': 'other_industries_links'})
    
    # Extract and return the text content of each a tag within the span element
    industries = []
    if other_industries_span:
        for a_tag in other_industries_span.find_all('a'):
            industries.append(a_tag.text.strip())
    return industries

def extract_data_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extract data using the appropriate functions
    company_name = extract_company_name(soup)
    key_principal = extract_key_principal(soup)
    company_website = extract_company_website(soup)
    company_address = extract_company_address(soup)
    maps_location = extract_maps_location(soup)
    industry_list = extract_industry_list(soup)
    other_industries_list = extract_other_industries_list(soup)
    
    return {
        'Company Name': company_name,
        'Key Principal': key_principal,
        'Company Website': company_website,
        'Company Address': company_address,
        'Maps Location': maps_location,
        'Industry List': industry_list,
        'Other Industries List': other_industries_list
    }

def store_data_in_excel(data_list, filename='company_data.xlsx'):
    # Create a DataFrame
    df = pd.DataFrame(data_list)
    
    # Write the DataFrame to an Excel file
    df.to_excel(filename, index=False)

# Directory containing the HTML files
directory = 'Company'

# List to store the extracted data
data_list = []

# Traverse the directory and parse each HTML file
for filename in os.listdir(directory):
    if filename.endswith('.html'):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r', encoding='utf-8') as file:
            html_content = file.read()
            data = extract_data_from_html(html_content)
            data_list.append(data)

# Store the data in an Excel file
store_data_in_excel(data_list)