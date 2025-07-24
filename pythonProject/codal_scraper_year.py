import os

from selenium.webdriver.common import keys
from selenium.webdriver.support import expected_conditions as EC, expected_conditions
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from persiantools.jdatetime import JalaliDate
import logging

from selenium.webdriver.support.wait import WebDriverWait


# Get the directory where the script is running
script_dir = os.path.dirname(os.path.abspath(__file__))

# Move one directory up
parent_dir = os.path.dirname(script_dir)

# 1.1. Define the log directory (e.g., "../excle/year")
log_dir = os.path.join(parent_dir, "excle", "year", "logs")  # Platform-safe path

# Create the log directory if it doesn't exist
os.makedirs(log_dir, exist_ok=True)

# Define the log file path
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

# Go into the "Excel" directory inside the parent directory
download_dir = os.path.join(parent_dir, "excle\\year")

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

# maximize_window
driver.maximize_window()

# get the address html
driver.get("https://codal.ir/ReportList.aspx?search")
wait = WebDriverWait(driver, 10)

# Click the dropdown .select2-chosen" to activate it
dropdown = driver.find_element(By.CSS_SELECTOR, ".select2-chosen")
dropdown.click()
time.sleep(1)

# Find and type in the search field that appears
# Type and press Enter to get first
search_input = driver.find_element(By.CSS_SELECTOR, ".select2-input")
search_input.send_keys("فولاد")
time.sleep(1)

# press enter to show results the site is finding for فولاد
search_input.send_keys(Keys.ENTER)
# WebDriverWait(driver, 10).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="collapse-search-1"]/div[2]/div[1]/div/div/div/ul'))
# )
time.sleep(1)

# Find and type in the search field that appears
dropdown = Select(driver.find_element(By.CSS_SELECTOR, "#reportType"))
dropdown.select_by_visible_text("اطلاعات و صورت مالی سالانه")  # Properly selects the dropdown option
selected_text = dropdown.first_selected_option.text  # Get the selected option's text
time.sleep(2)

# press search button to see result
submit_button = driver.find_element(by=By.CSS_SELECTOR, value=".btn-block")
submit_button.click()
time.sleep(2)

# from date
from_date = driver.find_element(by=By.CSS_SELECTOR, value="#txtFromDate .ng-empty")
from_date.click()

time.sleep(1)
wait=WebDriverWait(driver, 10)
wait.until(expected_conditions.visibility_of_element_located((By.XPATH, '//*[@id="adm-1"]')))
time.sleep(5)

currect_year_month = driver.find_element(By.XPATH, '//*[@id="adm-1"]/header/span')
currect_year_month.click()
time.sleep(3)

element = wait.until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="adm-1"]/div[1]/div')))


current_year= driver.find_element(By.XPATH, '//*[@id="adm-1"]/div[1]/div/div/div/p')
current_year.click()
time.sleep(3)


five=driver.find_element(By.XPATH, '//*[@id="adm-1"]/div[1]/div/div/span[3]/span')

five.click()
element = wait.until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, '#adm-1')))

time.sleep(3)

# get date from jalali calendar
today = JalaliDate.today().day
print(today)

date_class=driver.find_element(By.CLASS_NAME,"ng-binding").text
print(date_class)

persian_today = str(today).translate('0123456789'.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))
print(persian_today)


# if date_class==persian_1:
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH,
        f'//*[@class="ng-binding"][normalize-space()="{persian_today}"]'))).click()
time.sleep(2)

# to date
to_date = driver.find_element(By.CSS_SELECTOR, "#txtToDate input")
to_date.click()
time.sleep(2)

wait.until(expected_conditions.visibility_of_element_located((By.XPATH, '//*[@id="adm-2"]')))
time.sleep(3)



try:

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH,
             f'//span[contains(@class, "today") and contains(@class, "valid") and normalize-space()="{persian_today}"]')
        )
    ).click()
    time.sleep(2)
except:
    print('ame')

# press search button to see result
submit_button = driver.find_element(by=By.CSS_SELECTOR, value=".btn-block")
submit_button.click()
time.sleep(2)


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