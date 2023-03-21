import openpyxl
from undetected_chromedriver import Chrome
from undetected_chromedriver import By
from fake_useragent import UserAgent
import time


user_agent = UserAgent().random
driver = Chrome()
data = []
def open_url():
    try:
        for page in range(7,29):
            url = f'https://www.g2.com/categories/crm?order=g2_score&page={page}#product-list'
            options = driver.options
            options.add_argument(f'user-agent={user_agent}')
            driver.maximize_window()
            driver.get(url=url)
            time.sleep(15)
            names_company = driver.find_elements(By.CSS_SELECTOR, ".d-ib.c-midnight-100.js-log-click")
            urls_company = []
            for name in names_company:
                url_company = name.get_attribute('href')
                urls_company.append(url_company)
            for url_company in urls_company:
                driver.get(url=url_company)
                try:
                    site = driver.find_element(By.CSS_SELECTOR, '.paper.paper--nestable.border-top').find_element(By.CSS_SELECTOR, '.link.js-log-click').get_attribute('href')
                except:
                    site= None
                name_company = driver.find_element(By.CSS_SELECTOR, '.c-midnight-100').text
                data.append([site,name_company])
    except Exception as ex:
        print(ex)

    finally:
        driver.close()
        driver.quit()

if __name__ == '__main__':
    open_url()
    # write data to xlsx file
    wb = openpyxl.load_workbook('data.xlsx')
    ws = wb.active
    for row in data:
        ws.append(row)
    wb.save('data.xlsx')