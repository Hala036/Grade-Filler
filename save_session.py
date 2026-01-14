from playwright.sync_api import sync_playwright

PORTAL_LOGIN_URL = "https://portal.alquds.edu/en/login"

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # shows real Chrome window
        context = browser.new_context()
        page = context.new_page()

        page.goto(PORTAL_LOGIN_URL)
        print("Log in manually, then navigate to a page that clearly indicates you're logged in.")
        input("After you're fully logged in, press Enter here...")

        context.storage_state(path="state.json")
        print("Saved session to state.json")
        browser.close()

if __name__ == "__main__":
    main()
