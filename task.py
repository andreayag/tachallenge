from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import time
import re
import os


URL = "https://itdashboard.gov/"
LINK_DIVEIN = "#home-dive-in"
OUTPUT_PATH = "output"
DEPARTMENT = "U.S. Agency for International Development"


class AgenciesProcess:
    agencies_data = {}
    uii_links = {}

    def __init__(self):
        self.pdf = PDF()
        self.sysfile = FileSystem()
        self.lib = Selenium()
        self.excel = Files()

        self.sysfile.create_directory(OUTPUT_PATH, False, True)
        self.lib.set_download_directory(
            os.path.join(os.getcwd(), f"{OUTPUT_PATH}/"))

    def get_agencies(self):

        found_agencies = []
        amounts = []
        try:
            self.lib.open_available_browser(URL)

            self.lib.wait_until_element_is_enabled(
                "//a[@class='btn btn-default btn-lg-2x trend_sans_oneregular']", 10, "Error")
            self.lib.click_link(LINK_DIVEIN)

            self.lib.wait_until_element_is_visible(
                '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]', 10, "Error")
            agencies = self.lib.find_elements(
                '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')

            for agency in agencies:
                agency_split = agency.text.split("\n")
                found_agencies.append(agency_split[0])
                amounts.append(agency_split[2])
            self.agencies_data = {
                'Agencies': found_agencies, 'Amounts': amounts}
        except:
            raise Exception(
                "Failed to get the agencies information from: " + URL)

    def agencies_to_excel(self):
        try:
            self.excel.create_workbook(
                OUTPUT_PATH + "/Agencies.xlsx").append_worksheet("Sheet", self.agencies_data, True)
            self.excel.rename_worksheet("Sheet", "Agencies")
            self.excel.save_workbook()

        except:
            raise Exception(
                "Failed to fill the excel file with the agencies information")
        finally:
            self.excel.close_workbook()

    def get_department_info(self):
        try:
            self.lib.wait_until_element_is_visible(
                "//div[@id='agency-tiles-widget']//img[@alt='Seal of the " + DEPARTMENT + "']", 10, "Error")

            self.lib.find_element(
                "//div[@id='agency-tiles-widget']//img[@alt='Seal of the " + DEPARTMENT + "']").click()

            self.excel.open_workbook(OUTPUT_PATH + "/Agencies.xlsx")
            self.lib.wait_until_element_is_visible(
                "//table[@id='investments-table-object']", 10, "Error")

            table_id = self.lib.find_element(
                "//table[@id='investments-table-object']")
            content = [["UII", "Bureau", "Investment Title",
                        "Total FY2021 Spending ($M)", "Type", "CIO Rating", "# Of Projects"]]

            
            self.lib.wait_until_element_is_visible("//select[@name='investments-table-object_length']//option[@value='10']",10,"Error")
            self.lib.select_from_list_by_value("//select[@name='investments-table-object_length']", "-1")
            self.lib.wait_until_element_is_visible("//a[@class='paginate_button next disabled']", 10, "Error")
            rows = table_id.find_element_by_tag_name(
                "tbody").find_elements_by_tag_name("tr")

            for row in rows:
                scrapedtable = []
                scrapedrow = row.find_elements_by_tag_name('td')
                uii = scrapedrow[0]
                title = scrapedrow[2].text
                try:
                    a_element = uii.find_element_by_tag_name('a').get_attribute("href")
                
                except:
                    a_element = ""
                self.uii_links[uii.text] = [a_element, title]

                for cell in scrapedrow:
                    scrapedtable.append(cell.text)
                content.append(scrapedtable)

            return content
        except:
            raise Exception("Failed to get " + DEPARTMENT + " information")
        finally:
            self.excel.close_workbook()

    def department_to_sheet(self, content):
        self.excel.open_workbook(OUTPUT_PATH + "/Agencies.xlsx")
        self.excel.create_worksheet("Individual Investment", content, True)
        self.excel.save_workbook()
        self.excel.close_workbook()

    def compare_pdf(self, title, filepath, name):

        pdf_text = self.pdf.get_text_from_pdf(filepath, 1)
        time.sleep(2)
        investment_title = re.search(
            r'Name of this Investment:(.*)2\.', pdf_text[1]).group(1).strip()
        uii_text = re.search(
            r'Unique Investment Identifier \(UII\):(.*)Section B', pdf_text[1]).group(1).strip()
        if name == uii_text:
            print(
                f'Unique Investment Identifier (UII): {name} found in PDF ({filepath}).')
        else:
            print(
                f'Unique Investment Identifier (UII) not found in PDF ({filepath}).')
        if title == investment_title:
            print(
                f'Name of this Investment: {title} found in PDF ({filepath}).')
        else:
            print(f'Name of this Investment not found in PDF ({filepath}).')

    def download_pdfs(self):
        try:
            for uii_name in self.uii_links:
                link = self.uii_links[uii_name][0]
                title = self.uii_links[uii_name][1]
                if link != "":
                    self.lib.go_to(link)

                    self.lib.wait_until_element_is_enabled(
                        '//*[contains(@id,"business-case-pdf")]//a', 10, "Error")
                    pdf_link = self.lib.find_element(
                        '//*[contains(@id,"business-case-pdf")]//a').get_attribute("href")

                    if pdf_link:
                        self.lib.find_element(
                            '//div[@id="business-case-pdf"]').click()

                    self.lib.wait_until_element_is_visible("//div[@id='business-case-pdf']//span[1]",5,"Error")
                    self.lib.wait_until_element_is_not_visible("//div[@id='business-case-pdf']//span[1]",10,"Error")

                    seconds_towait = 0
                    
                    filepath = OUTPUT_PATH + "/" + uii_name + ".pdf"

                    while seconds_towait < 10:
                        if not self.sysfile.does_file_exist(filepath):
                            time.sleep(1)
                            seconds_towait+= 1
                        else:
                            seconds_towait = 10
                    self.compare_pdf(title, filepath, uii_name)
        except:
            raise Exception(
                "Failed to download and proccess PDF file:" + filepath)

    def close_browsers(self):
        self.lib.close_all_browsers()


if __name__ == "__main__":
    agencies_process = AgenciesProcess()
    try:
        print("1: Getting agencies information.")
        agencies_process.get_agencies()
        print("1: Done.")
        print("2: Extracting data to excel file.")
        agencies_process.agencies_to_excel()
        print("2: Done.")
        print("3: Getting." + DEPARTMENT + " infomartion.")
        content = agencies_process.get_department_info()
        print("3: Done.")
        print("4: Extracting data from department to excel sheet.")
        agencies_process.department_to_sheet(content)
        print("4: Done.")
        print("5: Processing PDFs.")
        agencies_process.download_pdfs()
        print("5: Done.")

    finally:
        agencies_process.close_browsers()
