import requests
import pandas as pd
import time
from bs4 import BeautifulSoup
from tqdm import tqdm
import math

# URL of the page to parse
main_url = "http://dnb.com/business-directory/industry-analysis.agriculture_forestry_fishing_and_hunting.html"
headers = {"User-Agent": "Mozilla/5.0"}

# Parse the main page to extract all country links and numbers of companies
response = requests.get(main_url, headers=headers)
if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")

    # Extract all matching elements
    country_elements = soup.select("div.col-md-6.col-xs-6.data a")
    if country_elements:
        overall_data = []

        for country_element in country_elements:
            # Extract URL and number of companies
            country_url = "https://www.dnb.com" + country_element["href"]
            number_countries_element = country_element.select_one("span.number-countries")
            number_countries = int(
                number_countries_element.text.strip("()").replace(",", "").replace("\n", "").replace("(", "").replace(")", "")
            )

            # Extract country name
            country_name = country_element.text.strip().split('\n')[0]
            print(country_name)
            # parsed_country_name = ["China", "Brazil", "Poland", "Italy", "United States", "Argentina", "Ukraine", "Kazakhstan", "France", "Germany", "India", "Sweden", "South Africa", "Russian Federation", "Finland", "Spain", "Belgium", "Norway", "Colombia", "United Kingdom", "Hungary", "Romania", "Australia", "Republic Of Korea", "Portugal", "Netherlands", "Chile", "Canada", "Denmark", "Peru", "Czech Republic", "Vietnam", "Japan", "Morocco", "Latvia", "New Zealand", "Malaysia", "Slovakia", "Bulgaria", "Thailand", "Taiwan", "Bosnia-Herzegovina", "Moldova", "Algeria", "Austria", "Switzerland", "Croatia", "Serbia", "Kyrgyzstan", "Estonia", "Israel", "Ireland", "Belarus"]
            if country_name == 'Ecuador':
                break
            print(f"Country: {country_name}, URL: {country_url}, Number of companies: {number_countries}")

            # Append data to the list
            overall_data.append({"country_name": country_name, "url": country_url, "number_countries": number_countries})

        # Save data to an Excel sheet under "Overall Data"
        overall_df = pd.DataFrame(overall_data)
        with pd.ExcelWriter("/kagglebusiness_data.xlsx", engine="openpyxl") as writer:
            overall_df.to_excel(writer, sheet_name="Overall Data", index=False)

        print("Data saved to 'business_data.xlsx'.")
    else:
        print("No matching elements found.")
else:
    print(f"Failed to fetch the main page. Status code: {response.status_code}")

# List to store business details
business_data = []

# Function to extract key person name & website from a company details page
def scrape_company_details(company_url):
    response = requests.get(company_url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")

        # Extract Key Person Name
        key_person_element = soup.select_one(
            "span.company_data_point[name='key_principal'] span"
        )
        key_person = (
            key_person_element.text.strip() if key_person_element else "N/A"
        )
        key_person = key_person.replace("See more contacts", "").strip()

        # Extract Website
        website_element = soup.select_one(
            "span.company_profile_overview_underline_links a.ext-class"
        )
        website = website_element["href"].strip() if website_element else "N/A"

        return key_person, website
    return "N/A", "N/A"

# Function to scrape a single page
def scrape_page(url):
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")

        # Extract company names and details page URLs
        business_elements = soup.select("div.col-md-12.data div.col-md-6 a")

        for business in tqdm(business_elements):
            business_name = business.text.strip()
            business_url = "https://www.dnb.com" + business["href"]

            if business_name and business_url:
                key_person, website = scrape_company_details(business_url)

                business_data.append(
                    {
                        "Business Name": business_name,
                        "Details URL": business_url,
                        "Key Person": key_person,
                        "Website": website,
                    }
                )
                time.sleep(1)  # Delay to prevent getting blocked

# Iterate through all country URLs and scrape data
if overall_data:
    with pd.ExcelWriter("business_data_with_countries.xlsx", engine="openpyxl") as writer:
        for country in overall_data:
            country_name = country["country_name"]
            
            base_url = country["url"]
            total_pages = math.ceil(country["number_countries"] / 50)
            if total_pages > 0:
                total_pages = 10

            print(f"Scraping data for {country_name}...")

            # Clear business_data for each country
            business_data = []

            # Scrape first page
            print(f"Scraping page 1 for {country_name}...")
            scrape_page(base_url)

            # Scrape remaining pages
            for page in range(2, total_pages + 1):
                next_page_url = f"{base_url}?page={page}"
                print(f"Scraping page {page} for {country_name}...")
                scrape_page(next_page_url)
                time.sleep(2)  # Add delay to avoid getting blocked

            # Save data for the current country to a new sheet
            if business_data:
                df = pd.DataFrame(business_data)
                df.to_excel(writer, sheet_name=country_name[:31], index=False)  # Sheet name max length is 31 characters

            print(f"Data for {country_name} saved.")

    print("Scraping completed. Data saved in 'business_data_with_countries.xlsx'.")
else:
    print("No country data found. Skipping scraping.")
