#######################################################################
# Check Warranty Status
#######################################################################
# Automates checking the warranty status of Dell, HP and
# Lenovo products from an Excel file utilizing asyncio, the 
# Playwright API and Chromium.
#######################################################################
# Author: Jeffrey Frederick
# Date: 10/08/22
#######################################################################
import time
import asyncio
import pandas as pd
from datetime import timedelta
from timeit import default_timer as timer
from playwright.async_api import async_playwright

expiration_list = []

# open excel file and convert it to a data frame
df = pd.read_excel("input.xlsx")

# change data types to lists for iterationc
uniques = df['Unique #'].values.tolist()
serials = df['Serial #'].values.tolist()


async def run(unique, serial, index):
    index = index + 1

    # launch chromium browser driver
    print (f"starting web driver #{index} for unique #{unique}...")
    time.sleep(.5)
    async with async_playwright() as driver:
        print(f"(web driver #{index}) opening browser...")
        browser = await driver.chromium.launch(headless=True)
        page = await browser.new_page()

        # go to Dell's warranty page
        print(f"(web driver #{index}) [Dell] loading warranty page...")
        await page.goto("https://www.dell.com/support/home/en-us?app=warranty/")
        await page.wait_for_load_state('networkidle')

        # enter the serial number and click on search button
        print(f"(web driver #{index}) [Dell] looking up serial #{serial}...")
        page.on("dialog", lambda dialog: dialog.accept())
        await page.locator('#inpEntrySelection').type(serial)
        await page.locator('#btn-entry-select').click()

        # Dell has some kind of bot detection so have to wait for a timer to end
        await asyncio.sleep(32)

        # click on search button again
        await page.locator('#btn-entry-select').click()
        await page.wait_for_load_state('networkidle')

        # click on product details to view warranty
        await page.locator('#viewDetailsRedesign').click()
        await page.wait_for_load_state('networkidle')

        # grab warranty expiration date from product details
        expiration_date = await page.locator('#dsk-expirationDt').inner_text()
        print(f"(web driver #{index}) [Dell] warranty status found ({expiration_date}) for serial #{serial}.")

        # append expiration date to list for saving to excel sheet
        print(f"(web driver #{index}) [Dell] appending '{expiration_date}' to excel sheet for unique #{unique}...")  
        expiration_list.append(expiration_date)
        print(expiration_list)
            
        # close browser driver
        print(f"(web driver #{index}) [Dell] finished. closing driver...")
        await browser.close()
        print(f"(web driver #{index}) driver closed.")

async def main():
    unique_length = (len(uniques))
    print(f"looking up {unique_length} uniques...")
    
    #start loop
    for i in range(0, unique_length):
        await run(uniques[i], serials[i], i,)

# timer start
start = timer()

# asyncio start
asyncio.run(main())

# add expiration list to data frame column
df.insert(5, 'Expiration', expiration_list)

# output results to excel file
df.to_excel('output.xlsx')

# timer end
end = timer()
print(f"execution time: {timedelta(seconds=end-start)}")
