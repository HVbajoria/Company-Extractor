import os
import requests
import pandas as pd
import time
import math
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from zipfile import ZipFile
from tqdm import tqdm

# Headers for HTTP requests
default_headers = {"User-Agent": "Mozilla/5.0"}

# --- Company Details Scraper ---
def scrape_company_details(company_url, headers=default_headers):
    resp = requests.get(company_url, headers=headers)
    if resp.status_code != 200:
        return "N/A", "N/A"
    soup = BeautifulSoup(resp.text, 'html.parser')
    kp_elem = soup.select_one("span.company_data_point[name='key_principal'] span")
    key_person = kp_elem.get_text(strip=True).replace("See more contacts", "").strip() if kp_elem else "N/A"
    site_elem = soup.select_one("span.company_profile_overview_underline_links a.ext-class")
    website = site_elem['href'].strip() if site_elem else "N/A"
    return key_person, website

# --- Page Scraper for business listings ---
def scrape_business_listings(page_url, headers=default_headers):
    """
    Return list of {'Business Name', 'Details URL', 'Key Person', 'Website'} for a given listings page.
    """
    results = []
    resp = requests.get(page_url, headers=headers)
    if resp.status_code != 200:
        return results
    soup = BeautifulSoup(resp.text, 'html.parser')
    elems = soup.select("div.col-md-12.data div.col-md-6 a")
    for a in elems:
        name = a.get_text(strip=True)
        url = urljoin(page_url, a['href'])
        if not name or not url:
            continue
        kp, site = scrape_company_details(url, headers)
        results.append({
            'Business Name': name,
            'Details URL': url,
            'Key Person': kp,
            'Website': site
        })
        time.sleep(1)
    return results

# --- Recursive Crawler for all region combinations ---
def crawl_regions(url, path, leaves, headers=default_headers):
    """
    Recursively traverse all region pages under <div class="locationResults">
    stopping when no sub-regions are found. Accumulates leaf URLs with full paths.
    leaves: list to append {'path': [...], 'url': ...}
    """
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        return
    soup = BeautifulSoup(resp.text, 'html.parser')

    # Locate sub-region container
    loc_container = soup.find('div', class_='locationResults')
    if not loc_container:
        leaves.append({'path': path.copy(), 'url': url})
        return

    # Find all direct data divs under locationResults
    data_divs = loc_container.select('div.col-md-6.col-xs-6.data')
    if not data_divs:
        leaves.append({'path': path.copy(), 'url': url})
        return

    # Traverse each sub-region
    for div in data_divs:
        a = div.find('a', href=True)
        if not a:
            continue
        name = a.get_text(strip=True)
        child_url = urljoin(url, a['href'])
        # New path for recursion
        new_path = path + [name]
        time.sleep(1)
        crawl_regions(child_url, new_path, leaves, headers)

# --- Save list of dicts to Excel ---
def save_to_excel(data, filepath, sheet_name='Data'):
    if not filepath.lower().endswith('.xlsx'):
        filepath += '.xlsx'
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        pd.DataFrame(data).to_excel(writer, sheet_name=sheet_name, index=False)

# --- Main Routine ---
def main(start_url):
    # Prepare
    out_folder = 'output_excels'
    os.makedirs(out_folder, exist_ok=True)
    leaves = []

    # Crawl all region combinations
    crawl_regions(start_url, [], leaves)

    # If no sub-regions at all, treat start as leaf
    if not leaves:
        leaves.append({'path': [], 'url': start_url})
    print(leaves)
    # Log all URL combinations
    combos_file = os.path.join(out_folder, 'url_combinations.xlsx')
    combos_data = [{'Path': ' > '.join(l['path']) or '(root)', 'URL': l['url']} for l in leaves]
    save_to_excel(combos_data, combos_file, sheet_name='Combinations')

    # Scrape each leaf and save details
    excel_files = [combos_file]
    for leaf in leaves:
        path = leaf['path']
        url = leaf['url']
        # Count total listings to paginate
        resp = requests.get(url, headers=default_headers)
        if resp.status_code != 200:
            continue
        soup = BeautifulSoup(resp.text, 'html.parser')
        # Extract count from first data div if exists
        # Extract the count from the div
        results_summary_div = soup.select_one('div.results-summary')
        if results_summary_div:
            # Extract the text and parse the count
            text = results_summary_div.text.strip()
            count = int(text.split('of')[-1].strip("()").replace(",", "").replace("\n", "").replace("(", "").replace(")", ""))  # Get the number after 'of'
            print(count)
        else:
            print("Results summary div not found.")
        # Paginate through listings
        all_entries = []
        pages = max(math.ceil(count / 50), 1)
        if pages > 20:
          pages = 20
        for page in range(1, pages + 1):
            page_url = url + (f"?page={page}" if page > 1 else "")
            entries = scrape_business_listings(page_url)
            all_entries.extend(entries)
            time.sleep(2)
        # Save details
        fname = '_'.join([p.replace(' ', '_') for p in path] or ['root']) + '_details.xlsx'
        out_path = os.path.join(out_folder, fname)
        save_to_excel(all_entries, out_path, sheet_name='Details')
        excel_files.append(out_path)

    # Zip all outputs
    zip_path = 'all_towns_data.zip'
    with ZipFile(zip_path, 'w') as zf:
        for f in excel_files:
            zf.write(f, arcname=os.path.basename(f))

    print(f"Done. Output zip created: {zip_path}")

if __name__ == '__main__':
    start = input("Enter starting URL: ").strip()
    main(start)