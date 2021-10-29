
#Program Warning:
print("""Program Warning:
1) Never change the filename, sheetname, values, positioning or layout of 'Crypto Database.xlsx', as we referenced their rows/columns for data extraction or appendation
2) If you obeyed the above but the program suddenly returns an error or doesn't run as you expect, it is likely due to your internet connection, a mismatch in this program's global variable (ex: wrong link, etc.), or that CMC's webpage design has been changed (causing the location of elements and x-paths to change as well). Start debugging from this, as our program's logic has been reviewed multiple times and hence won't be the culprit""")

import time, re
import pyinputplus as pyip
import openpyxl as opx
import pandas as pd
from tqdm import tqdm
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException

#Function: inputs our CMC screening criteria, and shows its results
def CMC_ScreenShow(min_marketCap, min_volume):
    filter_icon = browser.find_element_by_css_selector("button.x0o17e-0.ewuQaX.sc-36mytl-0.bBafzO.table-control-filter")
    browser.execute_script("arguments[0].click();", filter_icon)       #the JavaScript way to 'force' the click if normal Selenium python way doesn't work for whatever reason (hence why we're forced to type this 'messy code', if not we can just use ActionChains)
    time.sleep(0.5)
    addFilter_icon = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/ul/li[5]/button')
    browser.execute_script("arguments[0].click();", addFilter_icon)
    time.sleep(0.5)
    
    allCryptocurrencies_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div[1]/div[1]/button')
    browser.execute_script("arguments[0].click();", allCryptocurrencies_button)
    Coins_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div[1]/div[2]/div[2]/button')
    browser.execute_script("arguments[0].click();", Coins_button)
    time.sleep(1)         #to give time so that changes to the CMC filter is applied properly before we move on and apply more criteria to it

    marketCap_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div[1]/div[2]/button')
    browser.execute_script("arguments[0].click();", marketCap_button)
    min_marketCap_input = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div/div[3]/div[1]/div[1]/input[1]')
    min_marketCap_input.clear()
    min_marketCap_input.send_keys(min_marketCap)
    applyFilter1_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div/div[3]/div[2]/div/button[1]')
    browser.execute_script("arguments[0].click();", applyFilter1_button)
    time.sleep(1)

    volume_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div/div[5]/button')
    browser.execute_script("arguments[0].click();", volume_button)
    min_volume_input = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div/div[6]/div[1]/div[1]/input[1]')
    min_volume_input.clear()
    min_volume_input.send_keys(min_volume)
    applyFilter2_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div/div[6]/div[2]/div/button[1]')
    browser.execute_script("arguments[0].click();", applyFilter2_button)
    time.sleep(1)

    showResults_button = browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[6]/div/div/div[2]/div[2]/button[1]')
    browser.execute_script("arguments[0].click();", showResults_button)
    time.sleep(3)       #we give time up to 3 secs for the results to load properly before continuing code execution

#Function: returns the "totalResults_formatted" figure (to determine "num_of_pages")
def CMC_ScreenResults():
    totalResults_raw = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[7]/p')))        #Will return an error if the element doesn't load within 5s. Syntax: by_css_selector = By.CSS_SELECTOR, etc.
    totalResults_text = totalResults_raw.text       #returns the text form of CMC's displayed screen results to python
    totalResults_Regex = re.compile(r'\d+$')        #to return only the digit in the end of the str "totalResults_text"
    totalResults = totalResults_Regex.findall(totalResults_text)        #results will be in form of a list
    totalResults_formatted = int(totalResults[-1])          #formats "totalResults" in int form
    return totalResults_formatted

#Global scope: setting up the webdriver's filepath, and "ignored_exceptions"
chromedriver_filepath = '/Users/delvinkennedy/sandbox/Python/chromedriver'
browser = wd.Chrome(chromedriver_filepath)
ignored_exceptions = (NoSuchElementException, StaleElementReferenceException, TimeoutException)
#Opens up link in browser and gets the global crypto market cap figure to auto-input it to screening criteria (min market cap and min volume)
link = 'https://coinmarketcap.com/'
browser.implicitly_wait(5); browser.get(link)
GlobalCryptoMarketCap_rawtext = browser.find_element_by_xpath('//*[@id="__next"]/div[1]/div[1]/div[1]/div[2]/div/div/div/div[2]/div/span[3]').get_attribute('textContent')          #use "get_attribute('textContent')" if ".text" doesn't work (likely because the text is 'hidden' inside the WebElement)
startPos = GlobalCryptoMarketCap_rawtext.find('$'); text = GlobalCryptoMarketCap_rawtext[startPos:]         #to locate only the market cap figure out of the whole raw text (by getting the position of '$')
GlobalCryptoMarketCap = text.replace('$', '').replace(',', '')      #removes '$' and ',', so that it returns only numbers (which will be converted from str to int later)
min_marketCap = int(round(0.0005 * int(GlobalCryptoMarketCap), 0))
min_volume = int(round(0.1 * min_marketCap, 0))

#Starts time (for performance tracking purposes only)
startTime = time.time()
print('Please wait for your results...')
#Runs "CMC_ScreenShow" based on min_marketCap and min_volume to show screen results, and "CMC_ScreenResults" to get the total number of screen results
CMC_ScreenShow(min_marketCap, min_volume)
print(f'''Screener Details:
- Min Market Cap: USD{min_marketCap:,}
- Min Volume: USD{min_volume:,}
NB: This automated screener will return coins ONLY. Reconfigure code if you want to include tokens.''')
totalResults_formatted = CMC_ScreenResults()
print(f"\nResults Found: {totalResults_formatted}")

#The Main Program (extract screening data into lists)
filteredCrypto_name = []; filteredCrypto_symbol = []
num_of_pages = (totalResults_formatted - 1) // 100 + 1
main_page = browser.find_element_by_tag_name("html")
for x in range(num_of_pages):
    remaining_results = totalResults_formatted - (x * 100)
    pages_parsed = x + 1        #since by default "for" loops start the count from 0, we need to add it by 1 to reflect the 'accurate' "pages_parsed" figure
    crypto_table = browser.find_element_by_tag_name('tbody')
    print(f"Parsing page {pages_parsed} of {num_of_pages}. Please wait...")
    for i in range(20):
        main_page.send_keys(Keys.DOWN)      #presses the down key 20x, so that our browser's direct view is only on the crypto table (to reduce the probability of stale element exception)
    for y in tqdm(range(min(remaining_results, 100)), desc='Progress'):       #CMC can only display a maximum of 100 results per page, "result" starts from 0 ("for" loops by default)
        result = y + 1      #"{result}" is derived from a pattern on the table's x-path for each crypto iteration
        attempts = 0
        while attempts <= 3:
            try:
                crypto_name = crypto_table.find_element_by_xpath(f'//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[5]/table/tbody/tr[{result}]/td[3]/div/a/div/div/p')
                crypto_symbol = crypto_table.find_element_by_xpath(f'//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[5]/table/tbody/tr[{result}]/td[3]/div/a/div/div/div/p')
            except ignored_exceptions:          #to handle "ignored_exceptions" (if triggered, will wait up to 2s before retrying the current "y" for loop)
                attempts += 1
                if attempts <= 3:
                    time.sleep(2)
                    continue
            else:
                try:
                    filteredCrypto_name.append(crypto_name.text)
                    filteredCrypto_symbol.append(crypto_symbol.text)
                except ignored_exceptions:          #to handle "ignored_exceptions" if element is detected and code execution has passed through our first try-except block, but for some reason we can't extract its text (if triggered, will wait up to 2s before retrying the current "y" for loop)
                    attempts += 1
                    if attempts <= 3:
                        time.sleep(2)
                        continue
                else:
                    for z in range(2):         #presses the down key 2x after each crypto's successful data extraction (after trial-error, we found out that each row in CMC's table is 'worth' 2x down key); this is because CMC won't load the data if our browser is not in direct view of it
                        main_page.send_keys(Keys.DOWN)
                    break
        else:
            raise Exception("Cannot fetch data from CoinMarketCap. Make sure CoinMarketCap's server isn't down, or check your internet connection")
    if pages_parsed < num_of_pages:         #clicks the "next_button" only if there are still results yet to be displayed
        remaining_pages = num_of_pages - pages_parsed
        if num_of_pages > 7:
            if pages_parsed <= 4 or remaining_pages == 3:
                nextButton_xpath = 10
            elif remaining_pages <= 2:
                nextButton_xpath = 9
            else:
                nextButton_xpath = 11
        else:
            nextButton_xpath = num_of_pages + 2
        next_button = browser.find_element_by_xpath(f'//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[4]/div[1]/div/ul/li[{nextButton_xpath}]/a')      #"nextButton_xpath" is derived from a pattern on CMC's next button's x-path (based on "num_of_pages")
        browser.execute_script("arguments[0].click();", next_button)
        time.sleep(3)       #gives time up to 3 secs after pressing 'next' so that "crypto_table" is properly loaded (and can be referenced) and specific elements inside it can be extracted without problems later on (ex: NoSuchElementException, missing data on extraction, skipped rows, etc.)
browser.quit()

#Ends time (for performance tracking purposes only)
endTime = time.time(); runTime = round(endTime - startTime, 2)
print(f"\nScreener has run successfully in {runTime} second(s).\n")
#Organizes screened data into a single dataframe
directory = '/Users/delvinkennedy/sandbox/Personal Investments/Cryptos/'
series_networkName = pd.Series(filteredCrypto_name)
series_cryptoSymbol = pd.Series(filteredCrypto_symbol)
df_filteredCrypto = pd.DataFrame([])        #creates a 'blank' dataframe to be filled (its like doing "[]" for lists, or "{}" for dicts)
df_filteredCrypto['Network Name'] = series_networkName
df_filteredCrypto['Crypto Symbol'] = series_cryptoSymbol
print('Last 3 rows of extracted data:')
print(df_filteredCrypto.tail(3))

#Loads existing crypto database, appends new crypto data, remove duplicates and missing values, and appends to "Crypto Database.xlsx"
startingRowIndex = 23; startingColIndex = 1         #the indexes where our table starts from (note that unlike in openpyxl, in pandas indexing starts from 0)
df_cryptodatabase = pd.read_excel(f"{directory}Crypto Database.xlsx", skiprows=range(0, startingRowIndex-1), usecols=range(startingColIndex, startingColIndex+2))
df_cryptocombined = df_cryptodatabase.append(df_filteredCrypto, ignore_index=True)
df_cryptocombined = df_cryptocombined.astype(str)           #converts all data to the same str form so that "drop_duplicates" run without issues; it won't work if the data types are different (ex: UTF-8 str and UTF-4 str, etc.)
df_cryptocombined = df_cryptocombined.drop_duplicates(['Crypto Symbol'])
df_nan_filter = (df_cryptocombined['Network Name'] == 'nan')        #as we converted all to str, "np.nan" will be converted as 'nan'
df_nan_index = df_cryptocombined[df_nan_filter].index          #gets the row index of the row where the column 'Network Name' has the value 'nan'
df_cryptocombined = df_cryptocombined.drop(index=df_nan_index)      #deletes the row based on the specified index (in this case, "df_nan_index")
book = opx.load_workbook(f'{directory}Crypto Database.xlsx')
writer = pd.ExcelWriter(f'{directory}Crypto Database.xlsx', engine='openpyxl')        #we save on a new file in case if our program doesn't run as intended (so that we still have the 'original' file uncorrupted)
writer.book = book; writer.sheets = {ws.title: ws for ws in book.worksheets}
for sheetname in writer.sheets:
    df_cryptocombined.to_excel(writer, sheet_name='Crypto Database', startrow=startingRowIndex, startcol=startingColIndex, index=False, header=False)
writer.save()
print('\n"Crypto Database.xlsx" has been updated.')
