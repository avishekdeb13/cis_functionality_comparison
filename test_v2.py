from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time, os

URL       = "https://drt.etribunals.gov.in/cis2.0/filing/login"
USERNAME  = "filingdrt1"
PASSWORD  = "Admin@123"
OUTPUT    = "dashboard.html"
WAIT_SECS = 120

def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    wait = WebDriverWait(driver, 30)

    try:
        print("[*] Opening login page...")
        driver.get(URL)
        time.sleep(2)

        print("[*] Filling username...")
        username_field = wait.until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        username_field.clear()
        username_field.send_keys(USERNAME)

        print("[*] Filling password...")
        password_field = driver.find_element(By.NAME, "password")
        password_field.clear()
        password_field.send_keys(PASSWORD)

        print("\n" + "="*55)
        print("  Solve the CAPTCHA in the browser, then click LOGIN.")
        print(f"  Waiting up to {WAIT_SECS} seconds...")
        print("="*55 + "\n")

        WebDriverWait(driver, WAIT_SECS).until(
            EC.url_contains("dashboard")
        )
        print("[✓] Dashboard loaded!")
        time.sleep(3)

        html = driver.page_source
        with open(OUTPUT, "w", encoding="utf-8") as f:
            f.write(html)

        print(f"[✓] HTML saved → {os.path.abspath(OUTPUT)}")
        print(f"    Size: {len(html):,} bytes")

    except Exception as e:
        print(f"[✗] Error: {e}")
        try:
            with open("fallback_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("[!] Fallback HTML saved.")
        except:
            pass

    finally:
        input("\nPress ENTER to close browser...")
        driver.quit()

if __name__ == "__main__":
    main()