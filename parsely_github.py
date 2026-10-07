from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import date, timedelta
from email.message import EmailMessage
import smtplib
import os
import time
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

# Files are stored beside this script.
# GitHub Actions will create parsely_auth.json from a GitHub Secret.
BASE_DIR = Path(__file__).resolve().parent

AUTH_FILE = str(BASE_DIR / "parsely_auth.json")
DOWNLOAD_FOLDER = str(BASE_DIR / "downloads")
LOG_FILE = str(BASE_DIR / "automation_log.txt")

os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

if not os.path.exists(AUTH_FILE):
    raise FileNotFoundError(
        "Parse.ly authentication file was not found. "
        "Create parsely_auth.json locally or configure the GitHub Actions "
        "workflow to create it from the PARSELY_AUTH_B64 secret."
    )


# ============================================================
# LOGGING FUNCTION
# ============================================================

def log(message):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


# ============================================================
# GMAIL CREDENTIALS
# ============================================================

SENDER_EMAIL = os.environ.get(
    "SENDER_EMAIL",
    "vedasri.vangari@databeat.io"
)

APP_PASSWORD = os.environ.get("APP_PASSWORD")

RECIPIENT_EMAIL = os.environ.get(
    "RECIPIENT_EMAIL",
    SENDER_EMAIL
)

if not APP_PASSWORD:
    raise RuntimeError(
        "APP_PASSWORD environment variable is missing. "
        "Set it locally or add it as a GitHub Actions Secret."
    )


# ============================================================
# YESTERDAY
# ============================================================

yesterday = date.today() - timedelta(days=1)

yesterday_str = yesterday.strftime("%Y-%m-%d")


print()
print("============================================================")
print("PARSE.LY TWO-SITE DAILY AUTOMATION")
print("============================================================")
print()

print("Yesterday:", yesterday_str)
print()

log("============================================================")
log("AUTOMATION STARTED")
log(f"REPORT DATE: {yesterday_str}")


# ============================================================
# SITES
# ============================================================

sites = [

    {
        "name": "Kitchn",

        "url": (
            f"https://dash.parsely.com/thekitchn.com/posts/"
            f"?start={yesterday_str}"
            f"&end={yesterday_str}"
            f"&interval=1d"
        ),

        "subject": "Parse.ly Kitchn Daily Report",

        "filename":
            f"parsely_kitchn_{yesterday_str}.xlsx"
    },

    {
        "name": "Apartment Therapy",

        "url": (
            f"https://dash.parsely.com/apartmenttherapy.com/posts/"
            f"?start={yesterday_str}"
            f"&end={yesterday_str}"
            f"&interval=1d"
        ),

        "subject": "Parse.ly Apartment Therapy Daily Report",

        "filename":
            f"parsely_apartment_therapy_{yesterday_str}.xlsx"
    }

]


# ============================================================
# SEND EMAIL FUNCTION
# ============================================================

def send_email(file_path, subject, site_name):

    print()
    print("============================================================")
    print("SENDING EMAIL")
    print("============================================================")

    print("Site:", site_name)
    print("Subject:", subject)
    print("Attachment:", file_path)
    print()

    log(f"EMAIL STARTED: {site_name}")

    if not os.path.exists(file_path):

        print("ERROR: Attachment file not found.")

        log(
            f"EMAIL ERROR: ATTACHMENT FILE NOT FOUND: {site_name}"
        )

        return False

    try:

        msg = EmailMessage()

        msg["Subject"] = subject
        msg["From"] = SENDER_EMAIL
        msg["To"] = RECIPIENT_EMAIL

        msg.set_content(
            f"Hello,\n\n"
            f"Please find the Parse.ly daily report for "
            f"{site_name} attached.\n\n"
            f"Report date: {yesterday_str}\n\n"
            f"Thanks"
        )

        with open(file_path, "rb") as f:
            file_data = f.read()

        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(file_path)
        )

        print("Connecting to Gmail...")

        log(
            f"CONNECTING TO GMAIL: {site_name}"
        )

        server = smtplib.SMTP(
            "smtp.gmail.com",
            587,
            timeout=30
        )

        server.ehlo()

        print("Starting TLS...")

        server.starttls()

        server.ehlo()

        print("TLS connection successful.")

        print("Logging into Gmail...")

        server.login(
            SENDER_EMAIL,
            APP_PASSWORD
        )

        print("Gmail login successful.")

        log(
            f"GMAIL LOGIN SUCCESS: {site_name}"
        )

        print("Sending email...")

        server.send_message(msg)

        print("Email sent successfully.")

        server.quit()

        print()
        print("EMAIL SUCCESS")
        print("Site:", site_name)
        print("Subject:", subject)
        print()

        log(
            f"EMAIL SUCCESS: {site_name}"
        )

        return True

    except Exception as e:

        print()
        print("EMAIL FAILED")
        print("Site:", site_name)

        print(
            "Error type:",
            type(e).__name__
        )

        print(
            "Error:",
            e
        )

        print()

        log(
            f"EMAIL FAILED: {site_name}"
        )

        log(
            f"EMAIL ERROR TYPE: {type(e).__name__}"
        )

        log(
            f"EMAIL ERROR: {e}"
        )

        return False


# ============================================================
# PLAYWRIGHT AUTOMATION
# ============================================================

with sync_playwright() as p:

    print("Starting browser...")

    log("STARTING BROWSER")

    browser = p.chromium.launch(
        headless=True,
        args=[
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-first-run",
            "--no-default-browser-check"
        ]
    )

    context = browser.new_context(
        storage_state=AUTH_FILE,
        accept_downloads=True
    )

    print("Browser started.")
    print()

    log("BROWSER STARTED")


    # ========================================================
    # PROCESS EACH SITE
    # ========================================================

    for site in sites:

        site_name = site["name"]

        url = site["url"]

        file_path = os.path.join(
            DOWNLOAD_FOLDER,
            site["filename"]
        )

        print()
        print()
        print("############################################################")
        print("PROCESSING:", site_name)
        print("############################################################")
        print()

        print("URL:")
        print(url)
        print()

        print("Output file:")
        print(file_path)
        print()

        log(
            f"STARTING SITE: {site_name}"
        )

        log(
            f"URL: {url}"
        )


        # ----------------------------------------------------
        # REMOVE OLD FILE
        # ----------------------------------------------------

        if os.path.exists(file_path):

            print(
                "Removing previous file..."
            )

            log(
                f"REMOVING OLD FILE: {site_name}"
            )

            try:

                os.remove(file_path)

            except Exception as e:

                print(
                    "Could not remove old file:"
                )

                print(e)

                log(
                    f"OLD FILE REMOVE FAILED: {site_name}"
                )

                log(
                    f"ERROR: {e}"
                )

                continue


        # ----------------------------------------------------
        # CREATE PAGE
        # ----------------------------------------------------

        page = (
            context.pages[0]
            if context.pages
            else context.new_page()
        )


        # ----------------------------------------------------
        # OPEN PARSE.LY
        # ----------------------------------------------------

        print(
            "--------------------------------------------"
        )

        print(
            "Opening Parse.ly..."
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"OPENING PARSE.LY: {site_name}"
        )

        try:

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=120000
            )

        except PlaywrightTimeoutError:

            print(
                "Page load timeout. Continuing..."
            )

            log(
                f"PAGE LOAD TIMEOUT: {site_name}"
            )

        page.wait_for_timeout(10000)

        print(
            "Parse.ly opened."
        )

        print()

        log(
            f"PARSE.LY OPENED: {site_name}"
        )


        # ----------------------------------------------------
        # STEP 1: PUBLISH DATE
        # ----------------------------------------------------

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 1: Publish Date"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 1 STARTED: PUBLISH DATE: {site_name}"
        )

        try:

            print(
                "Looking for Publish Date button..."
            )

            # ------------------------------------------------
            # METHOD 1
            # ------------------------------------------------

            publish_date = page.get_by_role(
                "button",
                name="Publish Date"
            )

            print(
                "Publish Date button count:",
                publish_date.count()
            )

            # ------------------------------------------------
            # METHOD 2
            # ------------------------------------------------

            if publish_date.count() == 0:

                print(
                    "Trying text selector..."
                )

                publish_date = page.get_by_text(
                    "Publish Date",
                    exact=True
                )

                print(
                    "Text selector count:",
                    publish_date.count()
                )

            # ------------------------------------------------
            # METHOD 3
            # ------------------------------------------------

            if publish_date.count() == 0:

                print(
                    "Trying partial text selector..."
                )

                publish_date = page.get_by_text(
                    "Publish Date",
                    exact=False
                )

                print(
                    "Partial text selector count:",
                    publish_date.count()
                )

            # ------------------------------------------------
            # METHOD 4
            # ------------------------------------------------

            if publish_date.count() == 0:

                print(
                    "Trying generic button text selector..."
                )

                publish_date = page.locator(
                    "button:has-text('Publish Date')"
                )

                print(
                    "Generic button count:",
                    publish_date.count()
                )

            # ------------------------------------------------
            # WAIT AND CLICK
            # ------------------------------------------------

            if publish_date.count() == 0:

                raise Exception(
                    "Could not find Publish Date using any selector."
                )

            publish_date.first.wait_for(
                state="visible",
                timeout=30000
            )

            print(
                "Publish Date found."
            )

            print(
                "Clicking Publish Date..."
            )

            publish_date.first.click(
                timeout=60000,
                no_wait_after=True
            )

            print(
                "Publish Date opened."
            )

            page.wait_for_timeout(3000)

            log(
                f"PUBLISH DATE OPENED: {site_name}"
            )


            # ------------------------------------------------
            # THIS YEAR
            # ------------------------------------------------

            print(
                "Looking for This Year..."
            )

            this_year = page.get_by_role(
                "radio",
                name="this year"
            )

            print(
                "This Year radio count:",
                this_year.count()
            )

            if this_year.count() == 0:

                this_year = page.get_by_text(
                    "this year",
                    exact=False
                )

                print(
                    "This Year text count:",
                    this_year.count()
                )

            if this_year.count() == 0:

                raise Exception(
                    "Could not find This Year option."
                )

            this_year.first.wait_for(
                state="visible",
                timeout=30000
            )

            this_year.first.click(
                timeout=60000,
                no_wait_after=True
            )

            print(
                "This Year selected."
            )

            page.wait_for_timeout(3000)

            log(
                f"THIS YEAR SELECTED: {site_name}"
            )

        except Exception as e:

            print()
            print(
                "PUBLISH DATE FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error type:",
                type(e).__name__
            )

            print(
                "Error:",
                e
            )

            print()

            log(
                f"PUBLISH DATE FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 2: APPLY
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 2: Apply"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 2 STARTED: APPLY: {site_name}"
        )

        try:

            apply_button = page.get_by_role(
                "button",
                name="Apply"
            )

            apply_button.wait_for(
                state="visible",
                timeout=30000
            )

            apply_button.click(
                timeout=60000,
                no_wait_after=True
            )

            print(
                "Apply clicked."
            )

            page.wait_for_timeout(10000)

            log(
                f"APPLY CLICK COMPLETED - WAITING FOR REPORT: {site_name}"
            )

            log(
                f"APPLY SUCCESS: {site_name}"
            )

        except Exception as e:

            print()
            print(
                "APPLY FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error:",
                e
            )

            log(
                f"APPLY FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 3: EXPORT
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 3: Export"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 3 STARTED: EXPORT: {site_name}"
        )

        try:

            export_button = page.get_by_role(
                "button",
                name="Export"
            )

            export_button.wait_for(
                state="visible",
                timeout=30000
            )

            export_button.click()

            print(
                "Export opened."
            )

            page.wait_for_timeout(1500)

            log(
                f"EXPORT MENU OPENED: {site_name}"
            )

        except Exception as e:

            print()
            print(
                "EXPORT MENU FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error:",
                e
            )

            log(
                f"EXPORT MENU FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 4: TOP 10,000
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 4: TOP 10,000"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 4 STARTED: TOP 10,000: {site_name}"
        )

        try:

            top_10000 = page.get_by_text(
                "10,000",
                exact=False
            )

            top_10000.wait_for(
                state="visible",
                timeout=30000
            )

            top_10000.click()

            print(
                "Top 10,000 selected."
            )

            page.wait_for_timeout(1000)

            log(
                f"TOP 10,000 SELECTED: {site_name}"
            )

        except Exception as e:

            print()
            print(
                "TOP 10,000 FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error:",
                e
            )

            log(
                f"TOP 10,000 FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 5: SPREADSHEET
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 5: SPREADSHEET"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 5 STARTED: SPREADSHEET: {site_name}"
        )

        try:

            spreadsheet_option = page.get_by_text(
                "Spreadsheet",
                exact=True
            )

            spreadsheet_option.wait_for(
                state="visible",
                timeout=30000
            )

            spreadsheet_option.click()

            print(
                "Spreadsheet selected."
            )

            page.wait_for_timeout(1500)

            log(
                f"SPREADSHEET SELECTED: {site_name}"
            )

        except Exception as e:

            print()
            print(
                "SPREADSHEET FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error:",
                e
            )

            log(
                f"SPREADSHEET FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 6: FINAL EXPORT
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 6: EXPORT SPREADSHEET"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 6 STARTED: FINAL EXPORT: {site_name}"
        )

        try:

            export_final = page.get_by_role(
                "button",
                name="Export"
            )

            export_final.wait_for(
                state="visible",
                timeout=30000
            )

            print(
                "Final Export button found."
            )

            print(
                "Clicking Export..."
            )

            print()

            log(
                f"FINAL EXPORT BUTTON FOUND: {site_name}"
            )

            with page.expect_download(
                timeout=120000
            ) as download_info:

                export_final.click(
                    timeout=10000,
                    no_wait_after=True
                )

            download = download_info.value

            print()
            print(
                "DOWNLOAD EVENT RECEIVED"
            )

            print(
                "--------------------------------------------"
            )

            print(
                "Original filename:"
            )

            print(
                download.suggested_filename
            )

            print()

            print(
                "Temporary path:"
            )

            print(
                download.path()
            )

            log(
                f"DOWNLOAD EVENT RECEIVED: {site_name}"
            )

            log(
                f"ORIGINAL FILENAME: {download.suggested_filename}"
            )

            download.save_as(
                file_path
            )

            print()
            print(
                "File saved to:"
            )

            print(
                file_path
            )

            log(
                f"DOWNLOAD SAVED: {site_name}"
            )

            log(
                f"FILE PATH: {file_path}"
            )

        except Exception as e:

            print()
            print(
                "DOWNLOAD FAILED"
            )

            print(
                "Site:",
                site_name
            )

            print(
                "Error:",
                e
            )

            log(
                f"DOWNLOAD FAILED: {site_name}"
            )

            log(
                f"ERROR TYPE: {type(e).__name__}"
            )

            log(
                f"ERROR: {e}"
            )

            continue


        # ----------------------------------------------------
        # STEP 7: VERIFY FILE
        # ----------------------------------------------------

        print()

        print(
            "--------------------------------------------"
        )

        print(
            "STEP 7: VERIFY FILE"
        )

        print(
            "--------------------------------------------"
        )

        log(
            f"STEP 7 STARTED: VERIFY FILE: {site_name}"
        )

        file_found = False

        for i in range(30):

            if os.path.exists(file_path):

                size = os.path.getsize(
                    file_path
                )

                if size > 0:

                    file_found = True

                    print()
                    print(
                        "DOWNLOAD SUCCESSFUL"
                    )

                    print()
                    print(
                        "File:"
                    )

                    print(
                        file_path
                    )

                    print()
                    print(
                        "Size:"
                    )

                    print(
                        f"{size:,} bytes"
                    )

                    print()

                    log(
                        f"DOWNLOAD SUCCESSFUL: {site_name}"
                    )

                    log(
                        f"FILE SIZE: {size:,} bytes"
                    )

                    break

            time.sleep(1)

        if not file_found:

            print()
            print(
                "FILE NOT FOUND"
            )

            print(
                file_path
            )

            print()

            log(
                f"FILE NOT FOUND: {site_name}"
            )

            continue


        # ----------------------------------------------------
        # STEP 8: SEND EMAIL
        # ----------------------------------------------------

        email_success = send_email(
            file_path,
            site["subject"],
            site_name
        )


        # ----------------------------------------------------
        # STEP 9: DELETE LOCAL FILE
        # ----------------------------------------------------

        if email_success:

            print()

            print(
                "--------------------------------------------"
            )

            print(
                "CLEANING UP LOCAL FILE"
            )

            print(
                "--------------------------------------------"
            )

            try:

                os.remove(
                    file_path
                )

                print(
                    "Local file deleted."
                )

                print(
                    "Email contains the attachment."
                )

                log(
                    f"LOCAL FILE DELETED: {site_name}"
                )

            except Exception as e:

                print(
                    "Could not delete local file:"
                )

                print(e)

                log(
                    f"LOCAL FILE DELETE FAILED: {site_name}"
                )

                log(
                    f"ERROR: {e}"
                )

        else:

            print()

            print(
                "Email failed, so the local file was kept."
            )

            print(
                file_path
            )

            log(
                f"EMAIL FAILED - FILE KEPT: {site_name}"
            )

        print()

        print(
            "Finished:",
            site_name
        )

        print()

        log(
            f"FINISHED SITE: {site_name}"
        )


    # ========================================================
    # FINISH
    # ========================================================

    print()

    print(
        "############################################################"
    )

    print(
        "ALL SITES PROCESSED"
    )

    print(
        "############################################################"
    )

    print()

    log(
        "ALL SITES PROCESSED"
    )

    print(
        "Closing browser..."
    )

    context.close()
    browser.close()

    log(
        "BROWSER CLOSED"
    )


print()

print(
    "Automation finished."
)

log(
    "AUTOMATION FINISHED"
)

log(
    "============================================================"
)
