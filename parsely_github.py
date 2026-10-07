from playwright.sync_api import sync_playwright

AUTH_FILE = r"C:\Users\User\parsely_automation\parsely_auth.json"

sites = [
    {
        "name": "Kitchn",
        "url": "https://dash.parsely.com/thekitchn.com/posts/"
    },
    {
        "name": "Apartment Therapy",
        "url": "https://dash.parsely.com/apartmenttherapy.com/posts/"
    }
]

with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)

    context = browser.new_context(
        storage_state=AUTH_FILE,
        accept_downloads=True
    )

    page = context.new_page()

    for site in sites:

        print()
        print("==========================================")
        print("Opening:", site["name"])
        print("==========================================")

        page.goto(
            site["url"],
            wait_until="domcontentloaded",
            timeout=120000
        )

        page.wait_for_timeout(10000)

        print(site["name"], "opened.")

        input(
            f"Check {site['name']} is logged in, "
            "then press ENTER..."
        )

    context.close()
    browser.close()
