from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time, os

URL       = "https://drt.etribunals.gov.in/cis2.0/filing/login"
USERNAME  = "filingdrt1"
PASSWORD  = "Admin@123"
OUTPUT    = "dashboard_full.html"
WAIT_SECS = 120

def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    actions = ActionChains(driver)

    try:
        # LOGIN
        print("[*] Opening login page...")
        driver.get(URL)
        time.sleep(3)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "user_name"))
        )
        driver.find_element(By.ID, "user_name").send_keys(USERNAME)
        driver.find_element(By.ID, "user_pass").send_keys(PASSWORD)

        print("\n" + "="*55)
        print("  1. Type CAPTCHA in the browser")
        print("  2. Click LOGIN")
        print("  Waiting up to 120 seconds...")
        print("="*55 + "\n")

        WebDriverWait(driver, WAIT_SECS).until(
            lambda d: "login" not in d.current_url
        )
        print("[OK] Logged in! URL: " + driver.current_url)
        time.sleep(3)

        # EXPAND ALL SIDEBAR MENUS
        print("[*] Expanding all sidebar menus...")

        # Try clicking all sidebar toggle items (accordion/collapse menus)
        for attempt in range(3):
            # Find all clickable sidebar menu items
            menu_selectors = [
                "li.nav-item",
                "a.nav-link",
                ".sidebar-heading",
                "[data-bs-toggle='collapse']",
                "[data-toggle='collapse']",
                ".menu-item",
                "mat-expansion-panel-header",
                "mat-list-item",
                ".mat-expansion-panel-header",
            ]

            clicked = 0
            for selector in menu_selectors:
                try:
                    items = driver.find_elements(By.CSS_SELECTOR, selector)
                    for item in items:
                        try:
                            if item.is_displayed():
                                driver.execute_script("arguments[0].scrollIntoView(true);", item)
                                actions.move_to_element(item).perform()
                                time.sleep(0.2)
                                item.click()
                                time.sleep(0.3)
                                clicked += 1
                        except:
                            pass
                except:
                    pass

            print(f"  Pass {attempt+1}: clicked {clicked} elements")
            time.sleep(1)

        # Also try JavaScript click on all collapsed items
        print("[*] Force-expanding collapsed panels via JS...")
        driver.execute_script("""
            // Click all collapsed Bootstrap items
            document.querySelectorAll('[aria-expanded="false"]').forEach(el => {
                try { el.click(); } catch(e) {}
            });
            // Expand all Angular Material panels
            document.querySelectorAll('mat-expansion-panel').forEach(el => {
                try { el.classList.add('mat-expanded'); } catch(e) {}
            });
            // Show all collapse divs
            document.querySelectorAll('.collapse').forEach(el => {
                el.classList.add('show');
            });
        """)
        time.sleep(2)

        # Scroll through entire page to trigger lazy loading
        print("[*] Scrolling page to load all content...")
        total_height = driver.execute_script("return document.body.scrollHeight")
        for i in range(0, total_height, 300):
            driver.execute_script(f"window.scrollTo(0, {i});")
            time.sleep(0.05)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)

        # SAVE HTML
        html = driver.page_source
        with open(OUTPUT, "w", encoding="utf-8") as f:
            f.write(html)

        size = os.path.getsize(OUTPUT)
        print("[OK] HTML saved: " + os.path.abspath(OUTPUT))
        print("     Size: " + str(size) + " bytes")

    except Exception as e:
        print("[ERROR] " + str(e))
        try:
            with open("fallback_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("[!] Fallback saved.")
        except:
            pass

    finally:
        input("\nPress ENTER to close browser...")
        driver.quit()

if __name__ == "__main__":
    main()