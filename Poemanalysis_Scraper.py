from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
import pandas as pd
import time
import csv
import sys
import numpy as np

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options  = webdriver.ChromeOptions()
    # suppressing output messages from the driver
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--window-size=1920,1080')
    # adding user agents
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument("--incognito")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # running the driver with no browser window
    chrome_options.add_argument('--headless')
    # disabling images rendering 
    prefs = {"profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    driver.set_page_load_timeout(60)
    driver.maximize_window()

    return driver

def scrape_poemanalysis(path):

    start = time.time()
    print('-'*75)
    print('Scraping poemanalysis.com ...')
    print('-'*75)
    # initialize the web driver
    driver = initialize_bot()

    # initializing the dataframe
    data = pd.DataFrame()

    # if no books links provided then get the links
    if path == '':
        name = 'poemanalysis_data.xlsx'
        # getting the books under each category
        links = []
        nbooks, npages = 0, 0
        homepage = 'https://poemanalysis.com/poem-explorer/?_paged='
        while True:
            npages += 1
            url = homepage + str(npages)
            driver.get(url)
            # scraping books urls
            titles = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.fwpl-item.el-fz703r")))
            for title in titles:
                try:
                    nbooks += 1
                    print(f'Scraping the url for book {nbooks}')
                    link = wait(title, 5).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                    links.append(link)
                except Exception as err:
                    print('The below error occurred during the scraping from poemanalysis.com, retrying ..')
                    print('-'*50)
                    print(err)
                    continue

            # checking the next page
            try:
                button = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.facetwp-page.next")))
            except:
                break
                    
        # saving the links to a csv file
        print('-'*75)
        print('Exporting links to a csv file ....')
        with open('poemanalysis_links.csv', 'w', newline='\n', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Link'])
            for row in links:
                writer.writerow([row])

    scraped = []
    if path != '':
        df_links = pd.read_csv(path)
        name = path.split('\\')[-1][:-4]
        name = name + '_data.xlsx'
    else:
        df_links = pd.read_csv('poemanalysis_links.csv')

    links = df_links['Link'].values.tolist()

    try:
        data = pd.read_excel(name)
        scraped = data['Title Link'].values.tolist()
    except:
        pass

    # scraping books details
    print('-'*75)
    print('Scraping Books Info...')
    print('-'*75)
    n = len(links)
    for i, link in enumerate(links):
        try:
            if link in scraped: continue
            driver.get(link)           
            details = {}
            print(f'Scraping the info for book {i+1}\{n}')

            # title and title link
            title_link, title = '', ''              
            try:
                title_link = link
                title = wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).get_attribute('textContent').replace('\n', '').strip().title() 
            except:
                print(f'Warning: failed to scrape the title for book: {link}')               
                
            details['Title'] = title
            details['Title Link'] = title_link                          
            # Author and author link
            author, author_link = '', ''
            try:
                a = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.category-style")))
                author = a.get_attribute('textContent').replace('\n', '').strip().title() 
                author_link = a.get_attribute('href')
            except:
                try:
                    p = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "p[class='gb-headline gb-headline-7bab6c20 gb-headline-text dynamic-term-class']")))
                    a = wait(p, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a")))
                    author = a.get_attribute('textContent').replace('\n', '').strip().title() 
                    author_link = a.get_attribute('href')
                except:
                    pass
                    
            details['Author'] = author            
            details['Author Link'] = author_link            

            # summary
            summary = ''
            try:
                summary = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//h2[contains(.,'Summary')]/following-sibling::p"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Summary'] = summary             
            
            # Analysis
            analysis = ''
            try:
                analysis = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//h2[contains(.,'Analysis')]/following-sibling::p"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Analysis'] = analysis                          
            # appending the output to the datafame        
            data = data.append([details.copy()])
            # saving data to csv file each 100 links
            if np.mod(i+1, 100) == 0:
                print('Outputting scraped data ...')
                data.to_excel(name, index=False)
        except:
            pass

    # optional output to excel
    data.to_excel(name, index=False)
    elapsed = round((time.time() - start)/60, 2)
    print('-'*75)
    print(f'poemanalysis.com scraping process completed successfully! Elapsed time {elapsed} mins')
    print('-'*75)
    driver.quit()

    return data

if __name__ == "__main__":
    
    path = ''
    if len(sys.argv) == 2:
        path = sys.argv[1]
    data = scrape_poemanalysis(path)

