"""
Title: XX General District Court – Automated Case Scraper
Author: Teng Zhang, PhD
Date: Apr 28, 2025.
Last Update: Dec 4, 2025 (removed sensitve info)
Description:
    This script automates the process of collecting large volumes of public 
    court case data from the a state district court website

    There is a video on my personal website visualizing the result of this code.

    The pipeline includes around 20-30 pre-defined functions saved in other files. The scraper performs:
    - Automated date-based search queries
    - Iteration through paginated results
    - Opening detailed case pages in new tabs
    - Extracting main table data and case-level details
    - Writing cleaned, structured data to local CSV files

    Note:
        This script requires external web pages and dynamic content from the 
        State X's court website. It is provided as part of a code sample to
        demonstrate my workflow to handle errors, scrape and collect info.
        It will not run outside the original scraping environment due
        to site dependencies.
"""

from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random

from selenium.common.exceptions import (
    ElementNotInteractableException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    NoSuchElementException,
    WebDriverException
)

from utils import append_case_data  # Custom helper for writing scraped data

# Base URL for General District Court website
st_url = 'https://xxxx/caseSearch.do'

# Local output path where CSV data is stored
out_path = "C:/Users/XX/Raw_Data"


def scrape_all_case_pages(driver, max_pages=20):
    """
    Main workflow to scrape case records across multiple dates.

    Process:
        1. Generate a list of weekday search dates.
        2. Loop through each date and submit a search query.
        3. For each date, loop through paginated results (up to 100 pages).
        4. Open each case link in a new tab and extract:
            - Case summary data
            - Hearing information
            - Service information
        5. Save data to CSV after each page.

    Parameters:
        driver (selenium.webdriver): Browser automation driver.
        max_pages (int): Max number of pages to attempt per date.
    """

    # test codes:
    # Generate weekday lists for multiple years
    #weekdays_2018 = get_weekdays_in_year_formatted(2018)
    #weekdays_2019 = get_weekdays_in_year_formatted(2019)

    # Example reference date (mostly for debugging)
    #date_code = weekdays_2019[1]
    # end of test codes

    driver.switch_to.window(main_page_handle)

    # Reinitialize driver (non-headless for reliability)
    driver = setup_driver(headless=False)

    # Create a unified list of all weekday dates across multiple years
    date_bag = []
    for yr in [2019, 2020, 2021, 2022, 2023, 2024, 2025]:
        date_bag.extend(get_weekdays_in_year_formatted(yr))

    # ------------------------------
    # Main date loop
    # ------------------------------
    for i in range(1, 1561):
        date_code = date_bag[i]

        main_page_handle = driver.window_handles[0]
        driver.switch_to.window(main_page_handle)

        # Submit search request for the given date
        submit_date_search(driver, date_code)

        # Load existing case IDs if main_table exists
        if pd.io.common.file_exists(f"{out_path}\\main_table.csv"):
            existing_cases = pd.read_csv(f"{out_path}\\main_table.csv")['case_number'].values

        # Identify which court we are currently scraping
        header_court_name = driver.find_element(By.ID, "headerCourtName")
        court_code = 43 if header_court_name.text == "Chesterfield General District Court " else 999

        # ------------------------------
        # Page loop for the same date
        # ------------------------------
        for page in range(100):

            # Fresh dataframes for each page
            case_info_df = pd.DataFrame()
            hearing_df = pd.DataFrame()
            service_df = pd.DataFrame()
            main_df = pd.DataFrame()

            print(f"Processing search result page {page}...")

            driver.switch_to.window(main_page_handle)

            # Extract the result table with retries
            for _ in range(10):
                try:
                    link_table = driver.find_element(By.CLASS_NAME, "tableborder")
                    table_html = link_table.get_attribute("outerHTML")
                    break
                except (WebDriverException, StaleElementReferenceException, NoSuchElementException):
                    time.sleep(1)

            # Parse table via BeautifulSoup
            table_soup = BeautifulSoup(table_html, "html.parser")
            main_page_infos = parse_main_page(table_soup)

            n_links = main_page_infos[1]

            # Check if "Next Page" button is usable
            interactable = check_next_button_interactable(driver)
            if interactable == 'disabled':
                continue  # No more pages for this date

            append_case_main_data(
                main_page_infos,
                page_num=page,
                court_code=court_code,
                interactable=interactable,
                nrows=n_links
            )
            save_main_data()

            # ------------------------------
            # Open detailed case pages
            # ------------------------------
            links = table_soup.find_all("tr")[1:]
            n_additional_tabs = 0

            for link in links:
                link_tag = link.find("a")
                if not link_tag:
                    continue

                case_number = link_tag.text.strip()
                relative_url = link_tag["href"]

                # Skip cases already collected
                if case_number in existing_cases:
                    print(f'{case_number} skipped (already collected).')
                    continue

                # Open case detail in new tab
                full_url = "https://eapps.courts.state.va.us/gdcourts/" + relative_url
                driver.execute_script("window.open(arguments[0]);", full_url)
                time.sleep(random.uniform(0.1, 0.5))
                n_additional_tabs += 1

            # If no new tabs were opened, try to move to next page
            if len(driver.window_handles) < 2:
                result = handle_next_button()
                if result == 'break':
                    break
                continue

            # ------------------------------
            # Scrape each detailed page
            # ------------------------------
            for handle in driver.window_handles[1:]:
                driver.switch_to.window(handle)

                cleaned_html = normalize_html(driver.page_source)
                detail_soup = BeautifulSoup(cleaned_html, "html.parser")

                case_number = get_text_after_label(detail_soup, "Case Number:")

                if rate_limit() or case_number == "":
                    driver.close()
                    continue

                # Extract details
                case_detail_info = parse_case_details(detail_soup, case_number)
                hearing_info = parse_hearing_info(detail_soup, case_number)
                service_info = parse_service_info(detail_soup, case_number)

                print("Scraping:", case_number)
                append_case_detail_data(case_detail_info, hearing_info, service_info)

                # Close detail tab
                driver.close()
                time.sleep(0.1)

            # Save all case detail data before navigating
            save_case_detail_data()

            result = handle_next_button()
            if result == 'break':
                break
            elif result == 'continue':
                continue


# -----------------------------------------------------------
# Script Execution
# -----------------------------------------------------------

if __name__ == "__main__":
    """
    Entry point for running the scraper end-to-end.
    Note:
        - This will launch a Selenium browser session.
    """
    try:
        # Initialize driver (set headless=True if running on server)
        driver = setup_driver(headless=False)

        # Navigate to the main search page
        driver.get(st_url)
        time.sleep(2)

        # Record initial window handle
        main_page_handle = driver.current_window_handle

        # Execute full scraping pipeline
        scrape_all_case_pages(driver)

    except Exception as e:
        print(f"An error occurred during execution: {e}")

    finally:
        # Always close the browser session cleanly
        try:
            driver.quit()
        except:
            pass

        print("Scraping session finished.")