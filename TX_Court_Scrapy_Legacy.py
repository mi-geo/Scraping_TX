#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Requests for crawling attorney page (LEGACY DEMO)
=================================================
Author: Teng
Date: 2024-05-29 - 2024-06-03
License: Personal

NOTE
----
- This script is a **legacy demonstration** of the workflow I used in the past.
- It **does NOT** reflect the current production code I am using.
- Cleaning, refactoring, and annotations were added with help from OpenAI's ChatGPT.

Description
-----------
A simple web crawler that drives Internet Explorer (via Selenium), opens a
Texas court reporting URL for each (county, year, month) combination, triggers
an AutoIt script to save the exported Excel report, and renames the downloaded
files to meaningful names.

Usage
-----
    python main.py

Requirements
-----------
    - selenium
    - pandas
    - AutoIt (external executable invoked via subprocess)
"""

import os
import re
import glob
import time
import subprocess

import pandas as pd
from selenium import webdriver
from selenium.webdriver.ie.service import Service

# ---------------------------------------------------------------------------
# Configuration / Paths
# ---------------------------------------------------------------------------

# Path to IE WebDriver (IE is deprecated; this is kept for historical reasons)
IE_DRIVER_PATH = r"C:\ProgramData\anaconda3\Lib\site-packages\IEDriverServer.exe"

# AutoIt script that clicks through the browser UI and saves the XLS file
AUTOIT_SCRIPT_PATH = r"C:\Users\PycharmProjects\TexasCourt\save_xls.exe"

# Initial URL to "warm up" the driver
INITIAL_URL = "https://www.google.com/"

# Default download directory for the browser
DOWNLOAD_PATH = r"C:\Users\46798566\Downloads"

# County list file (must contain columns like CT (name) and ID (numeric ID))
COUNTY_LIST_PATH = r"C:\Users\46798566\PycharmProjects\TexasCourt\County_list.xlsx"

# Load county list once. Expected columns:
#   CT: county name (string)
#   ID: county numeric ID (1-based index used by the website)
COUNTY_LIST = pd.read_excel(COUNTY_LIST_PATH)


# ---------------------------------------------------------------------------
# WebDriver setup
# ---------------------------------------------------------------------------

def create_driver():
    """
    Create and return an IE WebDriver instance with "eager" page load strategy.

    NOTE: IE / IE driver are deprecated technologies. This function is kept
    here to show how the original pipeline worked.
    """
    service = Service(IE_DRIVER_PATH)
    options = webdriver.IeOptions()
    options.page_load_strategy = "eager"
    options.attach_to_edge_chrome = True  # IE mode in Edge
    driver = webdriver.Ie(service=service, options=options)
    driver.set_page_load_timeout(2.5)
    return driver


# ---------------------------------------------------------------------------
# Core helper functions
# ---------------------------------------------------------------------------

def get_y_download(driver, url_use: str) -> None:
    """
    Open a URL with the webdriver and call the AutoIt script to save the report.

    Parameters
    ----------
    driver : selenium.webdriver
        Active Selenium WebDriver instance.
    url_use : str
        Full report URL to open.

    Notes
    -----
    - This uses an external AutoIt script that interacts with the browser GUI.
    - Timeout is deliberately short to avoid hanging.
    """
    try:
        driver.get(url_use)
    except Exception as e:
        print(f"Error loading page: {e}")

    try:
        print("AutoIt STARTS, do NOT move the mouse/keyboard...")
        result = subprocess.run(AUTOIT_SCRIPT_PATH, timeout=1.5)
        print("AutoIt ENDS (for now)")

        if result.returncode != 0:
            print("AutoIt script failed with exit code", result.returncode)
        else:
            print("Download handled successfully")
    except subprocess.TimeoutExpired:
        print("AutoIt script timed out.")
    except Exception as e:
        print(f"Error running AutoIt script: {e}")


def url_generate(mm: int,
                 yyyy: int,
                 countyID: int = 57,
                 case: str = "DSC_Felony_Activity_Detail_N") -> str | None:
    """
    Generate the report URL for a given month, year, county, and case type.

    Parameters
    ----------
    mm : int
        Month (1–12).
    yyyy : int
        Year (e.g., 2011–2023).
    countyID : int, default 57
        County ID as expected by the TX courts site.
    case : str, default "DSC_Felony_Activity_Detail_N"
        Report type identifier.

    Returns
    -------
    str | None
        Fully formatted URL or None if the case type is not handled.
    """
    if case == "DSC_Felony_Activity_Detail_N":
        # note the URL here is twisted for legal reasons
        url_const = (
            "https://xxxxx /xxxxx"
            "?ReportName=xxx/DSC_Felony_Activity_Detail_N.rpt"
            "&ddlFromMonth={fmm}&ddlFromYear={fyyyy}&txtFromMonthField=@FromMonth&txtFromYearField=@FromYear"
            "&ddlToMonth={tmm}&ddlToYear={tyyyy}&txtToMonthField=@ToMonth&txtToYearField=@ToYear"
            "&ddlCountyPostBack={county}&txtCountyPostBackField=@CountyID"
            "&ddlCourtAfterPostBack=0&txtCourtAfterPostBackField=@CourtID"
            "&chkAggregateMonthlyReport=0&export=1625"
        ).format(
            fmm=mm,
            fyyyy=yyyy,
            tmm=mm,
            tyyyy=yyyy,
            county=countyID,
        )
    else:
        print(f"Unhandled case type: {case}")
        return None

    return url_const


def new_file_name(mm: int,
                  yyyy: int,
                  countyID: int = 57,
                  case: str = "DSC_Felony_Activity_Detail_N") -> str:
    """
    Construct a standardized output filename for a downloaded report.

    Example
    -------
    DSC_Felony_Activity_Detail_N-Travis-2016-12.xls
    """
    county_name = COUNTY_LIST.CT[countyID - 1]
    month_str = f"{mm:02d}"
    year_str = str(yyyy)

    nname = f"{case}-{county_name}-{year_str}-{month_str}.xls"
    return nname


def rename_downloaded_file(nname: str,
                           case: str = "DSC_Felony_Activity_Detail_N") -> None:
    """
    Rename the raw downloaded XLS file to the standardized filename.

    Parameters
    ----------
    nname : str
        New file name to use.
    case : str
        Case type, used to determine the original filename pattern.
    """
    if case == "DSC_Felony_Activity_Detail_N":
        sname = "District_and_Statutory_County_Court_DSC_Felony_Activity_Detail_N.rpt.xls"
    else:
        print(f"Unhandled case type in rename_downloaded_file: {case}")
        return

    old_name = os.path.join(DOWNLOAD_PATH, sname)
    new_name = os.path.join(DOWNLOAD_PATH, nname)

    try:
        os.rename(old_name, new_name)
        print(f"File successfully renamed to {nname}")
    except FileNotFoundError:
        print("Downloaded file not found, could not rename.")
    except Exception as e:
        print(f"Error renaming file: {e}")


def partial_file_removal() -> None:
    """
    Remove any *.partial files left in the download folder.

    These are typically leftover artifacts from interrupted downloads.
    """
    partial_files = glob.glob(os.path.join(DOWNLOAD_PATH, "*.partial"))
    for file_path in partial_files:
        try:
            os.remove(file_path)
            print(f"{file_path} has been deleted successfully.")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")


def regenerate_names_based_on_missing_names(missed: str):
    """
    Given a missing filename, extract (county, year, month, countyID).

    Parameters
    ----------
    missed : str
        Expected format:
        'DSC_Felony_Activity_Detail_N-<CountyName>-YYYY-MM.xls'

    Returns
    -------
    tuple | None
        (county, year, month, countyID) or None if pattern does not match.
    """
    pattern = r"^DSC_Felony_Activity_Detail_N-(.*)-(\d{4})-(\d{2})\.xls$"
    match = re.search(pattern, missed)

    if not match:
        print("The filename does not match the expected pattern:", missed)
        return None

    county = match.group(1)
    year = match.group(2)
    month = match.group(3)

    try:
        countyID = COUNTY_LIST.loc[COUNTY_LIST["CT"] == county, "ID"].iloc[0]
    except IndexError:
        print(f"County '{county}' not found in COUNTY_LIST.")
        return None

    return county, year, month, int(countyID)


# ---------------------------------------------------------------------------
# Main crawling logic
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    case_used = "DSC_Felony_Activity_Detail_N"
    MISSING_FILES: list[str] = []

    try:
        driver = create_driver()
        driver.get(INITIAL_URL)

        # Example ranges: counties 1–254, years 2011–2023, months 1–12
        for countyID_used in range(1, 255):
            for y in range(2011, 2024):
                for m in range(1, 13):
                    nname = new_file_name(m, y, countyID=countyID_used, case=case_used)
                    full_path = os.path.join(DOWNLOAD_PATH, nname)

                    # Skip if we already have this file
                    if os.path.exists(full_path):
                        continue

                    print(
                        "Month", m,
                        "Year", y,
                        "at", COUNTY_LIST.CT[countyID_used - 1],
                        "County, ID:", countyID_used,
                    )

                    url_use = url_generate(m, y, countyID=countyID_used, case=case_used)
                    if url_use is None:
                        continue

                    get_y_download(driver, url_use)
                    rename_downloaded_file(nname, case=case_used)
                    time.sleep(2)

                    # Check if rename/download was successful
                    if not os.path.exists(full_path):
                        partial_file_removal()
                        MISSING_FILES.append(nname)

    except Exception as e:
        print(f"An error occurred in the main loop: {e}")
    finally:
        try:
            driver.quit()
            print("WebDriver closed.")
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Second pass: retry downloads for missing files
    # -----------------------------------------------------------------------

    SECOND_MISSING_LIST: list[str] = []

    # If there are no missing files, nothing to do.
    if MISSING_FILES:
        driver = create_driver()

        try:
            for MISSED in MISSING_FILES:
                missed_path = os.path.join(DOWNLOAD_PATH, MISSED)
                if os.path.exists(missed_path):
                    # Already downloaded in the meantime
                    continue

                x = regenerate_names_based_on_missing_names(MISSED)
                if x is None:
                    # Could not parse; skip
                    SECOND_MISSING_LIST.append(MISSED)
                    continue

                county, year_str, month_str, countyID = x
                year_int = int(year_str)
                month_int = int(month_str)

                url_use = url_generate(month_int, year_int, countyID=countyID, case=case_used)
                if url_use is None:
                    SECOND_MISSING_LIST.append(MISSED)
                    continue

                get_y_download(driver, url_use)

                nname = new_file_name(month_int, year_int, countyID=countyID, case=case_used)
                if nname == MISSED:
                    rename_downloaded_file(nname, case=case_used)
                    time.sleep(2.5)
                else:
                    SECOND_MISSING_LIST.append(nname)
                    continue

                final_path = os.path.join(DOWNLOAD_PATH, nname)
                if not os.path.exists(final_path):
                    partial_file_removal()
                    SECOND_MISSING_LIST.append(nname)

        except Exception as e:
            print(f"An error occurred in the second pass: {e}")
        finally:
            try:
                driver.quit()
                print("WebDriver closed for second pass.")
            except Exception:
                pass

    print("First missing list length:", len(MISSING_FILES))
    print("Second missing list length:", len(SECOND_MISSING_LIST))