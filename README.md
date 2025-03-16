## Web Scraping and Data Extraction Project
This project is designed to scrape HTML content from web pages, extract specific data points, and store the extracted data in an Excel file. The project uses Python along with several libraries such as Selenium, BeautifulSoup, and Pandas.

#### Project Structure
.
├── Company/                # Directory containing HTML files
├── main.py                 # Main script file
├── requirements.txt        # List of dependencies
└── README.md               # This README file

### Installation
Clone the repository:
```bash
git clone https://github.com/yourusername/your-repo-name.git
cd your-repo-name
```

Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

Install the required dependencies:
```bash
pip install -r requirements.txt
```

### Usage
1. Place your HTML files in the Company directory.
2. Run the main script:
```bash
python main.py
```
3. The extracted data will be stored in an Excel file named company_data.xlsx.

### Functions
fetch_and_save_html
```python
def fetch_and_save_html(url, filename):
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")

    # Set up the Chrome driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Open the URL in the browser
    driver.get(url)
    
    # Wait for the page to load
    time.sleep(5)
    
    # Get the page source (HTML content)
    html_content = driver.page_source
    
    # Save the HTML content to a local file
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    print(f"HTML content saved to {filename}")
    
    # Close the browser
    driver.quit()
```
This function fetches the HTML content of a given URL and saves it to a local file.

#### extract_company_name
```python
def extract_company_name(soup):
    company_name_span = soup.find('span', {'data-tracking-name': 'Doing Business As:'})
    if company_name_span:
        return company_name_span.text.strip()
    else:
        return None
```
This function extracts the company name from the HTML content.

#### extract_company_website
```python
def extract_company_website(soup):
    company_website_span = soup.find('span', {'name': 'company_website'})
    if company_website_span:
        a_tag = company_website_span.find('a')
        if a_tag and 'href' in a_tag.attrs:
            return a_tag['href']
    return None
```
This function extracts the company website URL from the HTML content.

#### extract_key_principal
```python
def extract_key_principal(soup):
    key_principal_span = soup.find('span', {'name': 'key_principal'})
    if key_principal_span:
        inner_span = key_principal_span.find('span')
        if inner_span:
            return inner_span.contents[0].strip()
    return None
```
This function extracts the key principal's name from the HTML content.

#### extract_company_address
```python
def extract_company_address(soup):
    company_address_span = soup.find('span', {'name': 'company_address'})
    if company_address_span:
        a_tag = company_address_span.find('a')
        if a_tag:
            return a_tag.text.strip()
    return None
```
This function extracts the company address from the HTML content.

#### extract_maps_location
```python
def extract_maps_location(soup):
    company_address_span = soup.find('span', {'name': 'company_address'})
    if company_address_span:
        a_tag = company_address_span.find('a')
        if a_tag and 'href' in a_tag.attrs:
            return a_tag['href']
    return None
```
This function extracts the Google Maps location URL from the HTML content.

#### extract_industry_list
```python
def extract_industry_list(soup):
    industry_links_span = soup.find('span', {'name': 'industry_links'})
    industries = []
    if industry_links_span:
        for span in industry_links_span.find_all('span'):
            a_tag = span.find('a')
            if a_tag:
                industries.append(a_tag.text.strip())
            else:
                industries.append(span.text.strip())
    return industries
```
This function extracts the list of industries from the HTML content.

#### extract_other_industries_list
```python
def extract_other_industries_list(soup):
    other_industries_span = soup.find('span', {'name': 'other_industries_links'})
    industries = []
    if other_industries_span:
        for a_tag in other_industries_span.find_all('a'):
            industries.append(a_tag.text.strip())
    return industries
```
This function extracts the list of other industries from the HTML content.

#### extract_data_from_html
```python
def extract_data_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
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
```
This function extracts all the required data points from the HTML content and returns them as a dictionary.

#### store_data_in_excel
```python
def store_data_in_excel(data_list, filename='company_data.xlsx'):
    df = pd.DataFrame(data_list)
    df.to_excel(filename, index=False)
```
This function stores the extracted data in an Excel file.
----------- 

Feel free to contribute to this project by submitting issues or pull requests. For any questions, please contact the project maintainer.