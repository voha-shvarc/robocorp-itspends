from RPA.Browser.Selenium import Selenium
from selenium.common.exceptions import NoSuchElementException

import pandas as pd
import PyPDF2 as pdf
import time
import os

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


def get_investments_rows():
    investments_rows = browser.get_webelements("css:div.dataTables_scrollBody table#investments-table-object tbody tr")
    return investments_rows


def get_agency_investments_fields():
        investments_rows = get_investments_rows()

        uiis = [investment_row.find_elements_by_css_selector("td")[0].text for investment_row in investments_rows]
        bureaus = [investment_row.find_elements_by_css_selector("td")[1].text for investment_row in investments_rows]
        investment_titles = [investment_row.find_elements_by_css_selector("td")[2].text for investment_row in investments_rows]
        total_spendings = [investment_row.find_elements_by_css_selector("td")[3].text for investment_row in investments_rows]
        types = [investment_row.find_elements_by_css_selector("td")[4].text for investment_row in investments_rows]
        cio_ratings = [investment_row.find_elements_by_css_selector("td")[5].text for investment_row in investments_rows]
        num_of_projects = [investment_row.find_elements_by_css_selector("td")[6].text for investment_row in investments_rows]
        return [uiis, bureaus, investment_titles, total_spendings, types, cio_ratings, num_of_projects]


def write_agency_investments_to_excel(writer):
    agency_investments_fields = get_agency_investments_fields()

    print("Start writing agency investments")
    investments_df = get_investments_df(*agency_investments_fields)
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
    print(f"{investment_details_link.text}.pdf - downloaded")


def download_agency_investments_details_pdfs():
    print("Start downloading")
    investments_rows = browser.get_webelements("css:div.dataTables_scrollBody table#investments-table-object tbody tr")

    for investment_row in investments_rows:
        investment_details_link = get_investment_details_link(investment_row)
        if investment_details_link is not None:
            download_investment_details_pdf(investment_details_link)


def compare_investment_titles(investment_title, pdf_investment_title, file_path):
    if investment_title != pdf_investment_title:
        print(f"\nInvestment title on the site is not equal to investment title in pdf report in the file - {file_path}.\n")
    else:
        print(f"\nInvestment title on the site is equal to investment title in pdf report in the file - {file_path}.\n")


def compare_unique_identifiers(uii, pdf_uii, file_path):
    if uii != pdf_uii:
        print(f"UII on the site is not equal to UII in pdf report in the file - {file_path}.\n")
    else:
        print(f"UII on the site is equal to UII in pdf report in the file - {file_path}.\n")


def get_pdf_investment_title(text):
    start_name_investment = text.find("Name of this Investment:")
    end_name_investment = text.find("2.")

    pdf_investment_title = text[start_name_investment+25:end_name_investment].strip()
    return pdf_investment_title


def get_pdf_unique_identifier(text):
    end_name_investment = text.find("2.")
    end_identifier = text.find("Section B:")

    pdf_unique_identifier = text[end_name_investment+39:end_identifier].strip()
    return pdf_unique_identifier



def compare():
    investments_rows = get_investments_rows()
    uiis = [investment_row.find_elements_by_css_selector("td")[0].text for investment_row in investments_rows]
    investment_titles = [investment_row.find_elements_by_css_selector("td")[2].text for investment_row in investments_rows]


    output_dir_files = os.listdir("./output")
    pdf_file_paths = [file_path for file_path in output_dir_files if ".pdf" in file_path]

    for pdf_file_path in pdf_file_paths:
        with open(f"./output/{pdf_file_path}", "rb") as pdf_file:
            pdf_reader = pdf.PdfFileReader(pdf_file)
            first_page = pdf_reader.getPage(0)
            text = first_page.extractText()

            pdf_investment_title = get_pdf_investment_title(text)
            pdf_unique_identifier = get_pdf_unique_identifier(text)

            for index, uii in enumerate(uiis):
                if uii == pdf_file_path[:-4]:
                    compare_investment_titles(investment_titles[index], pdf_investment_title, pdf_file_path)
                    compare_unique_identifiers(uii, pdf_unique_identifier, pdf_file_path)

                    print(f"site title - {investment_titles[index]}")
                    print(f"pdf title - {pdf_investment_title}")
                    print(f"site uii - {uii}")
                    print(f"pdf uii - {pdf_unique_identifier}\n\n")
                    break


def main():
    try:
        writer = pd.ExcelWriter('./output/stats.xlsx', engine='xlsxwriter')

        open_website(SITE_URL)

        agencies_overviews_divs = get_agencies_overviews_divs()

        write_agencies_to_excel(writer, agencies_overviews_divs)

        go_to_agency_details_page(agencies_overviews_divs)

        show_all_investments_in_table()

        write_agency_investments_to_excel(writer)

        download_agency_investments_details_pdfs()

        compare()
        
        writer.save()
    finally:
        browser.close_all_browsers()


if __name__ == "__main__":
    main()
