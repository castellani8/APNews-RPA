from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os
import time
import re
import urllib.request
import xlwt
from urllib.parse import unquote
from RPA.Robocorp.WorkItems import WorkItems


class NewsScraper:
    def __init__(self):
        wi = WorkItems()
        wi.get_input_work_item()
        self.search_phrase = wi.get_work_item_variable('search_phrase')
        self.category = wi.get_work_item_variable('category')
        
        options = webdriver.FirefoxOptions()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument('--disable-web-security')
        options.add_argument("--start-maximized")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36")
        self.driver = webdriver.Firefox(options=options)
        
        self.output_dir = "output"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
    def search_news(self):
        url = "https://apnews.com/"
        self.driver.get(url)
        
        WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".icon-magnify"))
        )

        self.driver.find_element(By.CSS_SELECTOR, ".icon-magnify").click()
        search_box = self.driver.find_element(By.CSS_SELECTOR, 'input[name="q"]')
        search_box.send_keys(self.search_phrase)
        search_box.submit()
        
    def filter_by_category(self):
        if self.category != 'None':
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".SearchFilter"))
            )
            self.driver.find_element(By.CSS_SELECTOR, ".SearchFilter").click()

            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//label[contains(span, '${category}')]/input[@type='checkbox']"))
                )
            except:
                return
        
            self.driver.find_element(By.XPATH, "//label[contains(span, '${category}')]/input[@type='checkbox']").click()
    
    def sort_by_recent(self):
        time.sleep(2)
        close_button = self.driver.find_element(By.CSS_SELECTOR, "button.ot-close-icon")
        close_button.click()
        sort_by_select = self.driver.find_element(By.XPATH, '//select[@name="s"]')
        Select(sort_by_select).select_by_visible_text('Newest')

    def extract_data(self):
        self.driver.refresh()
        articles = self.driver.find_elements(By.CSS_SELECTOR, ".SearchResultsModule-results .PageList-items-item")
        data = []

        count = 1
        for article in articles:
            title = article.find_element(By.CSS_SELECTOR, " .PagePromo-title span").text

            try:
                img_element = self.driver.find_element(By.CSS_SELECTOR, f".PageList-items .PageList-items-item:nth-of-type({count}) .PagePromo-media a picture img")
                if img_element:
                    try:
                        img_url = img_element.get_attribute("srcset").split(',')[0].strip().split(' ')[0]
                        img_name = title + '.jpg'
                        self.download_image(img_url, img_name)
                    except (IndexError, AttributeError):
                        pass
            except NoSuchElementException:
                img_name = ""
            
            try:
                description = article.find_element(By.CSS_SELECTOR, ".PagePromo-description a span").text
            except NoSuchElementException:
                description = ""

            try:
                date = self.driver.find_element(By.CSS_SELECTOR, f".PageList-items .PageList-items-item:nth-of-type({count}) .PagePromo-date span span").text
            except NoSuchElementException:
                date = ""

            count_phrase = title.lower().count(self.search_phrase.lower()) + description.lower().count(self.search_phrase.lower())
            contains_money = bool(re.search(r'\$\s?\d+[\.,]\d+', title) or re.search(r'\$\s?\d+[\.,]\d+', description))

            data.append({
                'title': title,
                'description': description,
                'date': date,
                'image_file': img_name,
                'phrase_count': count_phrase,
                'contains_money': contains_money
            })
            count = count+1
        return data

    def download_image(self, img_url, img_name):
        try:
            if not os.path.exists('output/images'):
                os.makedirs('output/images')

            opener = urllib.request.build_opener()
            opener.addheaders = [('User-agent', 'Mozilla/5.0')]

            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(img_url, 'output/'+img_name)
        except:
            pass

    def save_to_file(self, data):
        output_file = os.path.join(self.output_dir, "news_data.xls")
        
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Sheet1")

        headers = ["Title", "Date", "Description", "Image File", "Phrase Count", "Contains Money"]
        for col, header in enumerate(headers):
            sheet.write(0, col, header)

        for row_index, item in enumerate(data, start=1):
            sheet.write(row_index, 0, item['title'])
            sheet.write(row_index, 1, item['date'])
            sheet.write(row_index, 2, item['description'])
            sheet.write(row_index, 3, item['image_file'])
            sheet.write(row_index, 4, item['phrase_count'])
            sheet.write(row_index, 5, str(item['contains_money']).upper())

        workbook.save(output_file)
    
    def run(self):
        self.search_news()
        self.filter_by_category()
        # self.sort_by_recent()
        data = self.extract_data()
        self.save_to_file(data)
        self.driver.quit()

if __name__ == "__main__":
    scraper = NewsScraper()
    scraper.run()
