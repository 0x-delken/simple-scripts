
#Program Warning:
print("""Program Warning:
1) Never change the filename, sheetname, values, positioning or layout of 'Stock Database.xlsx', as we referenced their rows/columns for data extraction or appendation 
2) If you obeyed the above but the program suddenly returns an error or doesn't run as you expect, it is likely due to your internet connection, a mismatch in this program's global variable (ex: wrong link, change in WSJ's country code or investing.com's country numbers, etc.), or that investing.com's webpage design has been changed (ex: new pop-ups, change in element locations, etc.). Start debugging from this, as our program's logic has been reviewed multiple times and hence won't be the culprit""")

import time
import openpyxl as opx
import numpy as np
import pandas as pd
import concurrent.futures
from tqdm import tqdm
from selenium import webdriver as wd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

#Specifying the save directory, the webdriver's filepath, and 'headless' option
directory = '/Users/delvinkennedy/sandbox/Personal Investments/Stocks/Stock Database & Valuation Models/'
chromedriver_filepath = '/Users/delvinkennedy/sandbox/Python/chromedriver'
options = wd.ChromeOptions(); options.add_argument("--headless")

#The Main Screening Function: Returns stock data for cash generation and deep value screen types based on the "country" and "number" (investing.com's code to represent a country) argument
#NB: We enclose everything inside a single function as we want to use threading on this later.
def Screener_by_Country(country, number):
	def SignUp_PopUp_Handler():			#Function: to handle investing.com's sign-up pop-ups
		SignUp_PopUp = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="PromoteSignUpPopUp"]/div[2]/i')))		#the "X" button to close the pop-up
		browser.execute_script("arguments[0].click();", SignUp_PopUp)
	def Screen_NumberOfResults():			#Function: to get the total results of a particular screen
		totalResults_raw = browser.find_element_by_xpath('//*[@id="fullColumn"]/div[11]/div[3]/span')		#the html text showing the number of our screen results
		totalResults = int(totalResults_raw.text)
		return totalResults
	#Setting up browser, screening links (cash generation and deep value), exchanges specifications, and category-separated lists to hold each country's screening results
	browser = wd.Chrome(chromedriver_filepath, options=options)			#we must declare "browser" inside the main function (remember, functions are local scope and can't access global variables)
	if country == 'US':
		exchange = '1'
	elif country == 'US2':
		country = 'US'			#resets the temporary 'US2' variable (only for denoting NASDAQ) back to 'US'; we need to denote NYSE ("US") and NASDAQ ("US2") because we want to exclude OTC stocks from being included (if we simply put 'a' as "exchange", OTC stocks will be included)
		exchange = '2'
	elif country == 'VN':
		exchange = '122'		#specially for VN, we want to screen only for Ho Chi Minh stocks (exclude Hanoi)
	else:
		exchange = 'a'
	cashGeneration_link = f'https://www.investing.com/stock-screener/?sp=country::{number}|sector::a|industry::a|equityType::ORD|exchange::{exchange}|qtotd2eq_us::0,100|ttmpr2rev_us::0,2|ttmastturn_us::0.5,5|margin5yr_us::5,100|opmgn5yr_us::5,100%3Ename_trans;1'
	deepValue_link = f'https://www.investing.com/stock-screener/?sp=country::{number}|sector::a|industry::a|equityType::ORD|exchange::{exchange}|qtotd2eq_us::0,50|qcurratio_us::2,20|ttmpr2rev_us::0,1|opmgn5yr_us::1,100|margin5yr_us::1,100|ttmastturn_us::0.25,5%3Ename_trans;1'
	screenerType = {'Cash Generation': cashGeneration_link, 'Deep Value': deepValue_link}
	stockData_by_parsedCountry = []
	#Loop to separate cash generation and deep value screens
	for screener in screenerType:
		while True:
			try:
				browser.get(screenerType[screener]); time.sleep(5)		#we give 5 secs for the screens to load properly (remember that we run multiple threads at once, hence the time to load screen on initial startup may be longer compared to if we only run a single thread)
			except:
				continue
			else:
				break
		try:
			SignUp_PopUp_Handler()		#if upon screener loading the pop-up shows up
		except (ElementNotInteractableException, TimeoutException):			#these two exceptions indicate that the pop-up didn't show up, hence "pass"
			pass
		else:
			pass 		#after handling the pop-up successfully, we "pass" and proceed to below code execution
		#Setting up pre-requisite variables for 'parsing loop'
		totalResults = Screen_NumberOfResults()
		num_of_pages = (totalResults - 1) // 50 + 1
		stockData_by_parsedScreener = []
		#Loop to parse and extract data for each screen type
		print(f'Country: {country}. Screener Type: {screener}.')
		for x in tqdm(range(num_of_pages), desc='Parsing'):
			remaining_results = totalResults - (x * 50)
			pages_parsed = x + 1
			stockData_by_parsedPage = []
			time.sleep(5)		#gives time up to 5 secs (remember we have multiple threads running, so internet traffic may be congested) until "stock_table" is properly loaded (and can be referenced), so that the specific elements inside it can be extracted without problems later on (ex: NoSuchElementException, missing data on extraction, skipped rows, etc.)
			stock_table = browser.find_element_by_tag_name('tbody')
			for y in range(min(remaining_results, 50)):       #investing.com can only display a maximum of 50 results per page, "result" starts from 0 ("for" loops by default)
				result = y + 1      #"{result}" is derived from a pattern on the table's x-path for each crypto iteration
				hundred_strCharacters = '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890'
				stockData_by_parsedRow = np.array([[hundred_strCharacters, hundred_strCharacters, hundred_strCharacters]])			#acts as a temporary placeholder; we need to specify the max characters we want each column to hold, if not your str will get cut off (ex: if you specify only 10 characters, any str characters above 10 will not get included)
				try:
					stock_name = stock_table.find_element_by_xpath(f'//*[@id="resultsTable"]/tbody/tr[{result}]/td[2]')			#we put them on "try" block in case if when extracting data suddenly the annoying pop-up shows up
					stock_ticker = stock_table.find_element_by_xpath(f'//*[@id="resultsTable"]/tbody/tr[{result}]/td[3]')
				except (StaleElementReferenceException, NoSuchElementException):		#these two exceptions indicate that most likely the pop-up indeed shows up (hence why we can't access the element 'behind' the pop-up)
					try:
						SignUp_PopUp_Handler()
					except (ElementNotInteractableException, TimeoutException):			#if for some reason we can't access the element not because of the pop-up, we can just refresh and retry the current loop (which will re-wait for "stock table" to properly load)
						browser.refresh()
						continue
					else:
						continue
				else:
					try:
						stockName = str(stock_name.text)
						stockTicker = str(stock_ticker.text)		#to make sure the format of tickers returned will be in str (even on tickers like 2045, etc.)
						if country == 'CN':			#IBKR only provides access to CN stocks where its ticker starts from 0 or 6 (see comments on "Stock Database.xlsx")
							hasAccess_filter = stockTicker[0] not in ('0', '6')
							stockCountry = 'SKIP' if hasAccess_filter else country
						elif country == 'US':			#to get rid of all non-USD denominated cross-listings (the pattern: ticker length is more than 4, and ends with 'F')
							crossListing_filter = (len(stockTicker) > 4) and (stockTicker[-1] == 'F')
							stockCountry = 'SKIP' if crossListing_filter else country
						elif country == 'KR':			#to get rid of all CNY denominated cross-listings (the pattern: starts with '9')
							CN_crossListing_filter = (stockTicker[0] == '9')
							stockCountry = 'SKIP' if CN_crossListing_filter else country
						elif country == 'EU':			#to get rid of "EU" stocks not in specified regions like XE, FR, etc. (the pattern: ends with the specified lower case letter; see "EU_regions" for the complete list)
							EU_region = stockTicker[-1]
							try:
								EU_regions[EU_region]
							except KeyError:
								stockCountry = 'SKIP'
							else:
								stockCountry = EU_regions[EU_region]
								stockTicker = stockTicker[0:len(stockTicker)-1]
						else:
							stockCountry = country
						stockData_by_parsedRow[0,0] = stockName
						stockData_by_parsedRow[0,1] = stockTicker
						stockData_by_parsedRow[0,2] = stockCountry
						stockData_by_parsedPage.append(stockData_by_parsedRow)
					except (StaleElementReferenceException, NoSuchElementException):		#same thing as the above, handling pop-ups
						try:
							SignUp_PopUp_Handler()
						except (ElementNotInteractableException, TimeoutException):
							browser.refresh()
							continue
						else:
							continue
			stockData_by_parsedScreener.extend(stockData_by_parsedPage)			#extends the stockData numpy arrays into "stockData_by_parsedScreener" list
			if pages_parsed < num_of_pages:			#clicks the "next_button" only if there are still results yet to be displayed
				while True:
					try:
						next_button = browser.find_element_by_xpath(f'//*[@id="paginationWrap"]/div[3]/a')
					except (StaleElementReferenceException, NoSuchElementException):
						try:
							SignUp_PopUp_Handler()
						except (ElementNotInteractableException, TimeoutException):			#same thing as the above, handling pop-ups
							browser.refresh()
							continue
						else:
							continue
					else:
						browser.execute_script("arguments[0].click();", next_button)
						break
		stockData_by_parsedCountry.extend(stockData_by_parsedScreener)
	numpy_stockData_by_parsedCountry = np.vstack(stockData_by_parsedCountry)			#v-stacks the list of arrays into a single array
	stockData.append(numpy_stockData_by_parsedCountry); time.sleep(5)			#we give 5 secs for each thread to append their data to global variable (appending that much data will definitely consume time, and we don't want to prematurely close the thread when the data is still 'in transit', as it can result in missing data)
	browser.quit()

#Setting up global variables: WSJ's country codes vs investing.com's country numbers, WSJ's EU region codes, and "stockData" to hold all extracted data in one place
country_numbers = {'US': 5, 'US2': 5, 'JP': 35, 'EU': 72,			#'US2' is a temporary variable to denote NASDAQ (will be adjusted back to 'US' inside the function)
				'CA': 6, 'AU': 25, 'SG': 36, 'KR': 11,
				'MY': 42, 'ID': 48, 'CN': 37, 'HK': 39, 'TW': 46,
				'TH': 41, 'PH': 45, 'VN': 178}
EU_regions = {'a':'NL', 'b':'BE', 'd':'XE', 'm':'IT', 'l':'UK', 'u':'PT', 'p':'FR', 'o':'NO', 'e':'ES', 's':'SE', 'v':'AT', 'z':'CH'}
stockData = []
#Start time (for performance tracking purposes only)
startTime = time.time()
#Setting up the threads and running them
print('Please wait...')
with concurrent.futures.ThreadPoolExecutor() as threader:			#by default the number of threads will be determined based on your CPU power (i5, i7, etc.), where the higher the spec, the more the number of threads allowed; you can also manually specify the max allowed number of threads by the kwarg "max_workers="
	threads = []
	for country, number in country_numbers.items():
		x = threader.submit(Screener_by_Country, country, number)
		threads.append(x)
	for thread in concurrent.futures.as_completed(threads):
		thread.result()
#Ends time (for performance tracking purposes only)
endTime = time.time(); runTime = round(endTime - startTime, 2)
print(f'\n\nAll screeners has run successfully in {runTime} second(s).')
#Organizes screened data so that it can be put into a single DataFrame
numpy_stockData = np.vstack(stockData)
df_stockData = pd.DataFrame(numpy_stockData)
df_stockData.columns = ['Stock Name', 'Stock Ticker', 'Country']
totalScreenedStocks = df_stockData.shape[0]			#returns the tuple "(rows, cols)", and extracts rows out of it (to be printed out for visual aid)
print(f'Total Screened Stocks: {totalScreenedStocks}')

#Loads existing stock database, appends new stockData, and remove duplicates, "SKIP", and missing values
startingRowIndex = 32; startingColIndex = 1; endingColIndex = 3         #the indexes where our table starts from (note that unlike in openpyxl, in pandas indexing starts from 0)
df_existingstockData = pd.read_excel(f"{directory}Stock Database.xlsx", skiprows=range(0, startingRowIndex-1), usecols=range(startingColIndex, endingColIndex+1))
df_combinedstockData = df_existingstockData.append(df_stockData, ignore_index=True)
df_combinedstockData = df_combinedstockData.astype(str)			#converts the whole dataframe's data type to the same str form so that "drop" will work without any issues
df_combinedstockData = df_combinedstockData.drop_duplicates(['Stock Name'])			#sometimes stocks can crosslist on different exchanges, hence why we remove duplicates of same name
df_combinedstockData = df_combinedstockData.drop_duplicates(['Stock Ticker', 'Country'])			#sometimes stocks can undergo name change, hence why we remove duplicates of stock ticker on the same country
df_skip_filter = (df_combinedstockData['Country'] == 'SKIP')
df_skip_index = df_combinedstockData[df_skip_filter].index
df_combinedstockData = df_combinedstockData.drop(index=df_skip_index)
df_nan_filter = (df_combinedstockData['Stock Name'] == 'nan')        #as we converted all to str, "np.nan" will be converted as 'nan'
df_nan_index = df_combinedstockData[df_nan_filter].index          #gets the row index of the row where the column 'Stock Name' has the value 'nan'
df_combinedstockData = df_combinedstockData.drop(index=df_nan_index)      #deletes the row based on the specified index (in this case, "df_nan_index")
#Appends updated stock database into "Stock Database.xlsx"
print('\nUpdating "Stock Database.xlsx"...')
book = opx.load_workbook(f'{directory}Stock Database.xlsx')
writer = pd.ExcelWriter(f'{directory}Stock Database.xlsx', engine='openpyxl')
writer.book = book; writer.sheets = {ws.title: ws for ws in book.worksheets}
for sheetname in writer.sheets:
	df_combinedstockData.to_excel(writer, sheet_name='Stock Database', startrow=startingRowIndex, startcol=startingColIndex, index=False, header=False)
writer.save()
print('"Stock Database.xlsx" has been updated.')
