from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from datetime import datetime
from typing import Dict, Tuple
from dateutil.relativedelta import relativedelta
import requests
from robocorp import log
import re
from openpyxl import Workbook
import os


class SitePage:

    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.selectors: Dict[str, Tuple[str, str]] = {
            "search_button": (By.CSS_SELECTOR, '[class*="search-trigger"]'),
            "search_field": (By.CSS_SELECTOR, '[title="Type search term here"]'),
            "submit_search_button": (By.CSS_SELECTOR, 'button[aria-label*="Search"]'),
            "sort_dropdown": (By.ID, 'search-sort-option'),
            "articles_container": (By.CSS_SELECTOR, 'div[class*="search-result"] article.gc'),
            "article_title_element": (By.CSS_SELECTOR, 'h3[class*="title"]'),
            "article_description_element": (By.CSS_SELECTOR, 'div[class*="excerpt"]'),
            "article_date_element": (By.CSS_SELECTOR, 'div.date-simple span[aria-hidden="true"]'),
            "article_image_element": (By.CSS_SELECTOR, 'div[class*=article-card] img'),
            "show_more_button": (By.CSS_SELECTOR, 'button[class*="show-more-button"]'),
        }

    def __find_element(self, name: str, driver_or_element: WebElement | WebDriver) -> WebDriverWait:
        by, value = self.selectors[name]
        return WebDriverWait(driver_or_element, 10).until(EC.element_to_be_clickable((by, value)))

    def __find_elements(self, name: str, driver_or_element: WebElement | WebDriver) -> WebDriverWait:
        by, value = self.selectors[name]
        return WebDriverWait(driver_or_element, 10).until(EC.presence_of_all_elements_located((by, value)))

    def __sort_results(self) -> None:
        log.info('Sorting results by date.')

        try:
            sort_dropdown = self.__find_element('sort_dropdown', self.driver)
            sort_dropdown_select = Select(sort_dropdown)
            sort_dropdown_select.select_by_value('date')
        except Exception as e:
            log.critical(f'Error sorting results: {e}')

        log.info('Results sorted.')

    def __download_image(self, url: str, file_path: str) -> None:
        log.debug(f'Downloading image at {url}.')

        try:
            response = requests.get(url)
            if response.status_code == 200:
                with open(file_path, 'wb') as file:
                    file.write(response.content)
        except Exception as e:
            log.critical(f'Error downloading image: {e}')

        log.debug('Image downloaded.')

    def __take_screenshot(self, url: str, file_path: str) -> None:
        log.debug(f'Taking screenshot at {url}.')

        try:
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.get(url)
            self.driver.get_screenshot_as_file(filename=file_path)
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
        except Exception as e:
            log.critical(f'Error taking screenshot: {e}')

        log.debug('Screenshot taken.')

    def __extract_date(self, string: str) -> str:
        match = re.search(r'\d{1,2} \w{3} \d{4}', string)  # Format 18 JUL 2024
        if match:
            return match.group(0)
        return ""

    def __contains_money(self, string: str) -> bool:
        money_patterns = [
            r'\$\d{1,3}(,\d{3})*(\.\d{1,2})?',  # Format $111,111.11
            r'\d{1,3}(,\d{3})*(\.\d{1,2})? dollars',  # Format 11 dollars
            r'\d{1,3}(,\d{3})*(\.\d{1,2})? USD'  # Format 11 USD
        ]
        for pattern in money_patterns:
            if re.search(pattern, string):
                return True
        return False

    def __parse_date(self, date_str: str) -> datetime:
        try:
            return datetime.strptime(date_str, '%d %b %Y')
        except Exception as e:
            log.critical(f'Error parsing date: {e}')

    def __clean_description(self, description: str) -> str:
        match = re.search(r'\.\.\.\s*(.*)', description)
        if match:
            return match.group(1)
        return description

    def __count_search_phrase(self, text: str, search_phrase: str) -> int:
        return text.lower().count(search_phrase.lower())

    def __is_date_within_filter(self, month_filter: int, article_date: datetime) -> bool:
        first_day_of_month = datetime(
            datetime.now().year, datetime.now().month, 1)
        if month_filter == 0:
            min_datetime = first_day_of_month - relativedelta(months=0)
        else:
            min_datetime = first_day_of_month - \
                relativedelta(months=month_filter-1)

        if article_date >= min_datetime:
            return True

        return False

    def __save_to_excel(self, articles_data: list) -> None:
        log.info("Generating excel.")

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Articles Data"

            headers = ["Title", "Date", "Description",
                       "Screenshot Path", "Contains Money", "Search Phrase Count"]

            ws.append(headers)

            for data in articles_data:
                ws.append(data)

            output_dir = "output"
            excel_filename = os.path.join(output_dir, "articles_data.xlsx")
            wb.save(excel_filename)
        except Exception as e:
            log.critical(f'Error saving to excel: {e}')

        log.info("Excel generated.")

    def search(self, query: str) -> None:
        log.info(f'Beginning search for {query}.')

        try:
            search_button = self.__find_element("search_button", self.driver)
            search_button.click()
            search_field = self.__find_element("search_field", self.driver)
            search_field.clear()
            search_field.send_keys(query)
            search_field.send_keys(Keys.RETURN)
            self.__sort_results()
        except Exception as e:
            log.critical(f'Error searching for {query}: {e}')

        log.info(f'Search for {query} successful.')

    def get_articles_info(self, month_filter: int, query: str) -> None:
        log.info(f'Getting articles info.')
        finish = False
        articles_count = 0
        articles_data = []

        while not finish:
            try:
                show_more_button = self.__find_element(
                    'show_more_button', self.driver)
                self.driver.execute_script(
                    "arguments[0].click();", show_more_button)
            except TimeoutException:
                finish = True
                log.info('No more articles to load.')

        articles_container = self.__find_elements(
            "articles_container", self.driver)

        for i in range(articles_count, len(articles_container)):
            article = articles_container[i]
            try:
                title = self.__find_element(
                    "article_title_element", article).text
                if "Today's latest from Al" in title:
                    articles_count += 1
                    continue

                date_text = self.__find_element(
                    "article_date_element", article).text
                date = self.__extract_date(date_text)
                parsed_date = self.__parse_date(date)

                within_filter = self.__is_date_within_filter(
                    month_filter=month_filter, article_date=parsed_date)
                if not within_filter:
                    articles_count += 1
                    continue

                description = self.__find_element(
                    "article_description_element", article).text
                cleaned_description = self.__clean_description(
                    description=description)

                image = self.__find_element("article_image_element", article)
                image_url = image.get_attribute("src")

                contains_money = self.__contains_money(description)

                count_search_phrase = self.__count_search_phrase(
                    f"{title} {description}", query)

                articles_data.append([
                    title,
                    parsed_date.strftime('%d %b %Y'),
                    cleaned_description,
                    image_url,
                    contains_money,
                    count_search_phrase
                ])

                articles_count += 1
            except (NoSuchElementException, TimeoutException, StaleElementReferenceException):
                articles_count += 1
                continue
            except Exception as e:
                log.critical(f'Error getting articles infos: {e}')
                articles_count += 1
                continue

        self.__save_to_excel(articles_data=articles_data)
