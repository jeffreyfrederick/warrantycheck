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
from numpy.lib.shape_base import row_stack
import openpyxl
from openpyxl.worksheet import worksheet
import pandas as pd
from playwright.async_api import async_playwright

# open excel file and convert it to a data frame
df = pd.read_excel("Book17_small.xlsx")
workbook = openpyxl.load_workbook('Book17_small.xlsx')
worksheet = workbook.active

expiration_list = []


# change data types to lists for iteration
uniques = df['Unique #'].values.tolist()
manufacturers = df['OEM'].values.tolist()
locations = df['Location'].values.tolist()
serials = df['Serial #'].values.tolist()
locations = df['Model'].values.tolist()


async def run(unique, manufacturer, location, serial, index, sleep):
    index = index + 1
    sleep = sleep * 5

        # launch chromium browser driver
    print("starting web driver #" + str(index) + " for unique #" +str(unique) + "...")
    time.sleep(.5)
    async with async_playwright() as driver:
        print("(web driver #" + str(index) + ") opening browser...")
        browser = await driver.chromium.launch(headless=True)
        page = await browser.new_page()

        # if the unit manufacturer is a "Dell" then we go to Dell's warranty page
        print("(web driver #" + str(index) + ") manufacturer 'Dell' found, loading warranty page...")
        await page.goto("https://www.dell.com/support/home/en-us?app=warranty/")
        await page.wait_for_load_state('networkidle')

            # enter the serial number and click on search button
        print("(web driver #" + str(index) + ") [Dell] page loaded, looking up serial #" + serial + "...")
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
        print("(web driver #" + str(index) + ") [Dell] warranty status found (" + expiration_date + ") for serial #" + serial + ".")

            # need to code! (create a new excel sheet column for expiration date and add value)
        print("(web driver #" + str(index) + ") [Dell] appending warranty expiration date (" + expiration_date + ") to excel sheet for unique #" + str(unique) + "...")
            
        expiration_list.append(expiration_date)
        print(expiration_list)
            

            # close browser driver
        print("(web driver #" + str(index) + ") [Dell] finshed. closing driver...")
        await browser.close()
        print("(web driver #" + str(index) + ") driver closed.")

async def main():

    unique_length = (len(uniques))
    print(unique_length)
    # do you know how we could improve the output of this engine? lets add a turbo!
    for i in range(0, unique_length):
        await run(uniques[i], manufacturers[i], locations[i], serials[i], i, i)

asyncio.run(main())

df.insert(5, 'Expiration', expiration_list)
print(df.to_excel('output.xlsx'))