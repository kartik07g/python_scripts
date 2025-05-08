import pandas as pd
import requests
import re
from bs4 import BeautifulSoup
import time
import os
from urllib.parse import urlparse

# Headers to simulate a real browser
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/115.0.0.0 Safari/537.36"
    )
}

# Input and output file paths
INPUT_EXCEL = "US Lead List.xlsx"       # should have one column: college names
OUTPUT_EXCEL = "college_info_output.xlsx"

# Read college names from Excel
df = pd.read_excel(INPUT_EXCEL)
college_names = df.iloc[:, 0].dropna().tolist()

def search_college_website(college_name):
    """Search for the college website using Bing (less restricted)."""
    query = f"{college_name} official site"
    url = f"https://www.bing.com/search?q={requests.utils.quote(query)}"
    print(f"🔍 Searching Bing for: {query}")

    allowed_domains = [".edu", ".org", ".ac.in", ".edu.in"]
    blacklist = ["linkedin.com", "facebook.com", "wikipedia.org", "youtube.com"]

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        results = soup.select("li.b_algo h2 a")

        for link in results[:10]:
            href = link.get("href")
            if href:
                domain = urlparse(href).netloc.lower()
                if any(domain.endswith(suffix) for suffix in allowed_domains):
                    if not any(bad in domain for bad in blacklist):
                        print(f"✅ Found: {href}")
                        return href
                    else:
                        print(f"⛔ Skipped (blacklisted): {href}")
                else:
                    print(f"⛔ Skipped (not allowed domain): {href}")

        print("🔍 No matching result found in Bing.")

    except requests.exceptions.RequestException as e:
        print(f"❌ Bing search failed: {e}")

    return ""

def extract_contact_info(website_url):
    """Simplified extraction of address, email, phone, and departments."""
    info = {
        "Address": "Not found",
        "Email": "Not found",
        "Phone": "Not found",
        "Departments": "Not found",
        "Website": website_url,
    }

    def get_visible_text(soup):
        return soup.get_text(" ", strip=True)

    try:
        resp = requests.get(website_url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        text = get_visible_text(soup)

        email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
        phone_match = re.search(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text)
        address_match = re.search(r"\d{2,5}\s[\w\s]+,\s?[A-Z]{2}\s\d{5}", text)

        info["Email"] = email_match.group(0) if email_match else "Not found"
        info["Phone"] = phone_match.group(0) if phone_match else "Not found"
        info["Address"] = address_match.group(0) if address_match else "Not found"

        dept_keywords = [
            "Engineering", "Science", "Math", "Arts", "Business", "Psychology",
            "Education", "Health", "Nursing", "Computer", "Programs", "Departments"
        ]
        found_depts = [word for word in dept_keywords if word.lower() in text.lower()]
        info["Departments"] = ", ".join(set(found_depts)) if found_depts else "Not found"

    except requests.exceptions.RequestException as e:
        print(f"❌ Failed to access {website_url}: {e}")
    except Exception as e:
        print(f"❌ Error extracting info from {website_url}: {e}")

    print("***** Extracted Info:", info)
    return info

# Main loop to process and append one-by-one
cnt = 1
for college in college_names:
    print(f"\nProcessing: {college}, count: {cnt}")
    cnt += 1
    college_info = {"College Name": college}

    website = search_college_website(college)
    if website:
        contact_info = extract_contact_info(website)
        college_info.update(contact_info)
    else:
        print(f"Website not found for {college}")
        college_info.update({
            "Address": "Not found",
            "Email": "Not found",
            "Phone": "Not found",
            "Departments": "Not found",
            "Website": ""
        })

    row_df = pd.DataFrame([college_info])

    try:
        if os.path.exists(OUTPUT_EXCEL):
            # Read existing data
            existing_df = pd.read_excel(OUTPUT_EXCEL)
            updated_df = pd.concat([existing_df, row_df], ignore_index=True)
            # Overwrite file with combined data
            with pd.ExcelWriter(OUTPUT_EXCEL, mode='w', engine='openpyxl') as writer:
                updated_df.to_excel(writer, index=False)
        else:
            # First time creation
            with pd.ExcelWriter(OUTPUT_EXCEL, mode='w', engine='openpyxl') as writer:
                row_df.to_excel(writer, index=False)
    except PermissionError:
        print(f"❌ Permission denied when trying to write to {OUTPUT_EXCEL}. Make sure it's closed and accessible.")
        continue

    time.sleep(2)

print(f"\n✅ All records saved incrementally to {OUTPUT_EXCEL}")
