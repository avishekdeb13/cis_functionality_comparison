from bs4 import BeautifulSoup
import csv, os

HTML_FILE = "dashboard.html"  # or dashboard.html - whichever you saved
CSV_FILE  = "menu_items.csv"

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, "html.parser")
base = "https://drt.etribunals.gov.in"

rows = []

# Find all nav-items in sidebar
sidebar = soup.find("ul", id="accordionSidebar")
if sidebar:
    nav_items = sidebar.find_all("li", class_="nav-item")
    for li in nav_items:
        # Main menu label
        main_link = li.find("a", class_="nav-link")
        if not main_link:
            continue
        main_name = main_link.get_text(strip=True)

        # Sub items inside collapse div
        dropdown = li.find("div", class_="collapse-inner")
        if dropdown:
            sub_links = dropdown.find_all("a", class_="dropdown-item")
            for a in sub_links:
                sub_name = a.get_text(strip=True)
                sub_url  = base + a.get("href", "")
                rows.append([main_name, sub_name, sub_url])
                print(f"  {main_name} > {sub_name}")
        else:
            # Top level direct link
            url = main_link.get("href", "")
            if url and url != "javascript:void(0)":
                full_url = base + url if url.startswith("/") else url
            else:
                full_url = ""
            rows.append([main_name, "", full_url])
            print(f"  {main_name} (direct)")

# Save CSV
with open(CSV_FILE, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["Main Menu", "Sub Menu", "URL"])
    w.writerows(rows)

print("\n[OK] Saved: " + os.path.abspath(CSV_FILE))
print("     Total rows: " + str(len(rows)))