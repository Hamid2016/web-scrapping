

# # working version with enter
import os
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from persiantools.jdatetime import JalaliDate
import logging



#1. make chrome download and logs into excle folder
# Get the directory where the script is running
script_dir = os.path.dirname(os.path.abspath(__file__))

# Move one directory up
parent_dir = os.path.dirname(script_dir)

# 1.1. Define the log directory (e.g., "../excle/one-month")
log_dir = os.path.join(parent_dir, "excle", "one-month", "logs")  # Platform-safe path

# 1.1.1. Create the directory if it doesn't exist
os.makedirs(log_dir, exist_ok=True)

# 1.1.2 Define the log file path
log_file = os.path.join(log_dir, "app.log")
# Configure logging
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    filemode='a',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    encoding='utf-8'  # ← Add this line (Python 3.9+)
)

logger = logging.getLogger(__name__)
logger.info("Logging configured successfully!")

# 1.2. Go into the "Excel" directory inside the parent directory
download_dir = os.path.join(parent_dir, "excle/one-month")

print("Excel Directory:", download_dir)


# Configure Chrome to auto-save PDFs
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": os.path.abspath(download_dir),
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True  # This forces download instead of preview

}

chrome_options.add_experimental_option("prefs", prefs)

# Launch browser
driver = webdriver.Chrome(options=chrome_options)
# driver = webdriver.Chrome()
driver.get("https://codal.ir/ReportList.aspx?search")
time.sleep(2)  # Wait for page to load

# 2. Click the dropdown .select2-chosen" to activate it
dropdown = driver.find_element(By.CSS_SELECTOR, ".select2-chosen")
dropdown.click()

# 3. Find and type in the search field that appears
# Type and press Enter to get first
search_input = driver.find_element(By.CSS_SELECTOR, ".select2-input")
search_input.send_keys("فولاد")
# Verify and log the entered text
entered_text = search_input.get_attribute("value")
logger.info(f"'{entered_text}' is selected")  # Logs: 'فولاد' is selected
time.sleep(2)

# 4. press enter to show results the site is finding for فولاد
search_input.send_keys(Keys.ENTER)  # This selects the first result
time.sleep(1)


# 5.1. Find and type in the search field that appears
dropdown = Select(driver.find_element(By.CSS_SELECTOR, "#reportType"))
dropdown.select_by_visible_text("گزارش عملکرد ماهانه")  # Properly selects the dropdown option
selected_text = dropdown.first_selected_option.text  # Get the selected option's text
logger.info(f"'{selected_text}' is selected")

time.sleep(2)



# 5.2. press search button to see result
submit_button = driver.find_element(by=By.CSS_SELECTOR, value=".btn-block")
submit_button.click()
time.sleep(2)



# 6.1. put today for from date
# 6.1.1. get today's date in Shamsi
today_shamsi = JalaliDate.today()
print("Today's Shamsi Date:", today_shamsi.strftime("%Y/%m/%d"))
logger.info(f"{today_shamsi} to")


# 6.1.2. Calculate 5 years before today
five_years_ago = today_shamsi.replace(year=today_shamsi.year - 5)
print("5 Years Ago (Shamsi):", five_years_ago.strftime("%Y/%m/%d"))
logger.info(f"{five_years_ago} from")


from_date = driver.find_element(by=By.CSS_SELECTOR, value="#txtFromDate .ng-empty")
from_date.clear()  # Clears the field
from_date.send_keys(five_years_ago.strftime("%Y/%m/%d"))  # Sets the date to 5 years before today

# 6.2. put today for to date
from_date = driver.find_element(by=By.CSS_SELECTOR, value="#txtToDate .ng-empty")
from_date.clear()  # Clears the field
from_date.send_keys(today_shamsi.strftime("%Y/%m/%d"))  # Sets the date to today

# 6.3. press search button to show result
submit_button = driver.find_element(by=By.CSS_SELECTOR, value=".btn-block")
submit_button.click()
time.sleep(2)
print(from_date.text)



# 7. download  excel
def dwn_excel(row):
    if row<=2:
        css_selector = f"#divLetterFormList .ng-scope:nth-child({row}) .icon-excel"
    else:
        css_selector = f".ng-scope:nth-child({row}) .icon-excel"

    try:
        excel_button = driver.find_element(by=By.CSS_SELECTOR, value=css_selector)
        #create an ActionChains instance
        action = ActionChains(driver)
        #move to the button, and then click it
        action.move_to_element(excel_button)\
            .click()\
            .perform()
    except Exception as e:  # Catch-all for other unexpected errors
        print(f"⚠️ Unexpected error (n={row}): {e}")
        logger.error(f"⚠️ Unexpected error (n={row}): {e}")

# 8. rename excle
def rename_excle(row):
    if row<=2:
        css_selector = f"#divLetterFormList .ng-scope:nth-child({row}) .letter-title"
    else:
        css_selector = f".ng-scope:nth-child({row}) .letter-title"
    try:
        # 8.1  extract the name to name excle
        element= driver.find_element("css selector", css_selector)
        print(element .text)
        excle_name=element.text
        # print(type(excle_name))
        time.sleep(2)

        # 8.2 renaming
        # 8.2.1 Find the latest file in the directory
        downloaded_file = max([os.path.join(download_dir, f) for f in os.listdir(download_dir)], key=os.path.getctime)


        #sleep for a few seconds so we can see the download
        time.sleep(10)

        # 8.2.2 Rename the file
        nasafe_name = (excle_name
            .replace('/', '-') .strip()  # Remove extra whitespace
        ) + '.xls'
        new_file_path = os.path.join(download_dir, nasafe_name)
    except Exception as e:  # Catch-all for other unexpected errors
        print(f"⚠️ Unexpected error (n={row}): {e}")
        logger.error(f"⚠️ Unexpected error (n={row}): {e}")

    try:
        os.rename(downloaded_file, new_file_path)
        logger.info(f"{new_file_path} is successfully downloaded")
    except Exception as e:  # Catch-all for other unexpected errors
        print(f"⚠️ Unexpected error (n={row}): {e}")
        logger.error(f"⚠️ Unexpected error (n={row}): {e}")

    time.sleep(5)
    print("Search submitted. You should now see results.")

# 9. find number of excle
def find_excle_number():
    # Find all elements that match both selector formats
    elements_div = driver.find_elements(By.CSS_SELECTOR, "#divLetterFormList .ng-scope .icon-excel")

    numbers_excle=len(elements_div)
    print(numbers_excle)
    logger.info(f"number of excle in this page is {numbers_excle}")

    # call functions download and rename
    for i in range(1,numbers_excle+1):
        dwn_excel(i)
        rename_excle(i)

# 10. return true if next page exist
def go_next_page():
    # go to next page
    try:
        # Try to find and click the "next page" button
        driver.find_element(By.CSS_SELECTOR, ".active + .ng-scope .ng-binding").click()
        time.sleep(2)  # Wait for page to load
        return True  # Success!
    except:
        return False  # No more pages

# 11. main loop. dwonload start from here
while True:

    # 1. Process current page
    find_excle_number()  # Fixed typo in function name

    # 2. Try to go to next page
    if not go_next_page():
        break  # Exit when no more pages

logger.info("Finished all pages!")




driver.quit()
