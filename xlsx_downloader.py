from io import BytesIO
import requests
from bs4 import BeautifulSoup
import openpyxl as xls


class XlsxDownloader:

    BASE_URL: str = "https://energy.ec.europa.eu"
    URL: str = f"{BASE_URL}/data-and-analysis/weekly-oil-bulletin_en"

    def download(self, all_: bool) -> xls.Workbook:
        download_link = self._find_download_link(all_=all_)
        xlsx_data = self._download(download_link)
        workbook = self._create_workbook(xlsx_data)
        return workbook

    def _create_workbook(self, xlsx_data: bytes) -> xls.Workbook:
        xlsx_file = BytesIO(xlsx_data)
        workbook = xls.load_workbook(xlsx_file)
        return workbook

    def _download(self, download_link: str) -> bytes:
        response = requests.get(download_link)
        response.raise_for_status()

        xlsx_data = response.content

        return xlsx_data

    def _find_download_link(self, all_: bool) -> str:
        response = requests.get(self.URL)
        response.raise_for_status()

        html = response.text

        soup = BeautifulSoup(html, "html.parser")

        ecl_div = soup.find("div", class_="ecl")
        assert ecl_div

        if not all_:
            ecl_file_div = ecl_div.find("div", class_="ecl-file")

        else:
            ecl_file_div = None

            ecl_file_divs = soup.find_all("div", class_="ecl-file")
            assert ecl_file_divs

            for div in ecl_file_divs:
                container = div.find("div", class_="ecl-file__container")
                if not container:
                    continue

                info = container.find("div", class_="ecl-file__info")
                if not info:
                    continue

                if "onwards" in info.text:
                    ecl_file_div = div
                    break

        assert ecl_file_div

        ecl_file_action_div = ecl_file_div.find("div", class_="ecl-file__action")
        assert ecl_file_action_div

        file_download_a = ecl_file_action_div.find(
            "a",
            class_="ecl-link ecl-link--standalone ecl-link--icon ecl-file__download",
        )
        assert file_download_a

        donwload_link = f"{self.BASE_URL}{file_download_a.get('href')}"

        return donwload_link


print(XlsxDownloader()._find_download_link(all_=False))
