from RPA.Browser.Selenium import Selenium
from selenium.common.exceptions import NoSuchElementException

import pandas as pd
import time

SITE_URL = "https://itdashboard.gov/"
DEP_TO_VISIT = "Department of Agriculture"

browser = Selenium()
browser.set_download_directory("./output")


def open_website(url):
    browser.open_available_browser(url)


def get_agencies_overviews_divs():
    browser.click_link("#home-dive-in")
    agency_div_css_selector = "css:div#agency-tiles-container div#agency-tiles-widget div.wrapper div.row div.col-sm-4"
    browser.wait_until_element_is_visible(agency_div_css_selector)
    agencies_divs = browser.get_webelements(locator=agency_div_css_selector)
    return agencies_divs


def write_agencies_to_excel(writer, agencies_divs):
    print("Start writing agencies")
    agencies = [agency_div.find_element_by_css_selector("span.h4").text for agency_div in agencies_divs]
    amounts = [agency_div.find_element_by_css_selector("span.h1").text for agency_div in agencies_divs]

    agencies_df = get_agencies_df(agencies, amounts)

    agencies_df.to_excel(writer, sheet_name="Agencies", index=False)
    print("Finish writing agencies")


def get_agencies_df(agencies, amounts):
    agencies_df = pd.DataFrame(
        {
            "Agnecy": agencies,
            "Spend amount": amounts
        }
    )
    return agencies_df


def go_to_agency_details_page(agencies_divs):
    for agency_div in agencies_divs:
        agency_name = agency_div.find_element_by_css_selector("span.h4").text
        if agency_name == DEP_TO_VISIT:
            agency_info_btn = agency_div.find_element_by_css_selector("a.btn")
            browser.click_element(agency_info_btn)
            return


def write_agency_investments_to_excel(writer):
    print("Start writing agency investments")
    show_all_investments_in_table()

    investments_rows = browser.get_webelements("css:div.dataTables_scrollBody table#investments-table-object tbody tr")
    td_list = investment_row.find_elements_by_css_selector("td")

    uiis = [td_list[0].text for investment_row in investments_rows]
    bureaus = [td_list[1].text for investment_row in investments_rows]
    investment_titles = [td_list[2].text for investment_row in investments_rows]
    total_spendings = [td_list[3].text for investment_row in investments_rows]
    types = [td_list[4].text for investment_row in investments_rows]
    cio_ratings = [td_list[5].text for investment_row in investments_rows]
    num_of_projects = [td_list[6].text for investment_row in investments_rows]

    investments_df = get_investments_df(uiis, bureaus, investment_titles, total_spendings, types, cio_ratings, num_of_projects)

    investments_df.to_excel(writer, sheet_name="Investments", index=False)
    print("Finish writing agency investments")


def show_all_investments_in_table():
    browser.wait_until_page_contains_element("css:div#investments-table-object_length label select", timeout=10)
    select_options = browser.get_webelements("css:div#investments-table-object_length label select option")
    for option in select_options:
        if option.get_attribute("value") == "-1":
            browser.click_element(option)
    time.sleep(10)


def get_investments_df(uiis, bureaus, investment_titles, total_spendings, types, cio_ratings, num_of_projects):
    investments_df = pd.DataFrame(
            {
                "UII": uiis,
                "Bureaus": bureaus,
                "Investment Title": investment_titles,
                "Total FY2021 Spending ($M)": total_spendings,
                "Type": types,
                "CIO Rating": cio_ratings,
                "# of Projects": num_of_projects
            }
        )
    return investments_df


def get_investment_details_link(investment_row):
    tds = investment_row.find_elements_by_css_selector("td")
    try:
        investment_link = tds[0].find_element_by_css_selector("a")
        return investment_link
        print(f"get investment link {investment_link.get_attribute('href')}")
    except NoSuchElementException:
        return None


def download_investment_details_pdf(investment_details_link):
    browser.click_link(investment_details_link, modifier="ctrl")

    browser.switch_window("new")
    browser.wait_until_element_is_enabled("css:div#business-case-pdf a")
    download_link = browser.get_webelement("css:div#business-case-pdf a")
    browser.click_link(download_link)
    browser.wait_until_page_does_not_contain("Generating PDF...", timeout=15)
    browser.close_window()
    browser.get_window_handles()
    browser.switch_window("main")
    print("downloaded pdf")


def download_agency_investments_details_pdfs():
    print("Start downloading")
    investments_rows = browser.get_webelements("css:div.dataTables_scrollBody table#investments-table-object tbody tr")
    for investment_row in investments_rows:
        investment_details_link = get_investment_details_link(investment_row)
        if investment_details_link is not None:
            print("I have link")
            download_investment_details_pdf(investment_details_link)


def main():
    try:
        writer = pd.ExcelWriter('./output/stats.xlsx', engine='xlsxwriter')

        open_website(SITE_URL)

        agencies_overviews_divs = get_agencies_overviews_divs()

        write_agencies_to_excel(writer, agencies_overviews_divs)

        go_to_agency_details_page(agencies_overviews_divs)

        write_agency_investments_to_excel(writer)

        download_agency_investments_details_pdfs()
        writer.save()
    finally:
        browser.close_all_browsers()


if __name__ == "__main__":
    main()