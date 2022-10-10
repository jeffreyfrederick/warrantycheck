#######################################################################
# Check Warranty Status
#######################################################################
# Automates checking the warranty status of Dell, HP and
# Lenovo products from an Excel file utilizing asyncio, the 
# Playwright API and Chromium
#######################################################################
# Author: Jeffrey Frederick
# Date: 10/08/22
#######################################################################

import asyncio
import pandas as pd
from playwright.async_api import async_playwright

# open excel file and convert it to a data frame
df = pd.read_excel("products.xlsx")

# change data types to lists for iteration
uniques = df['Unique'].values.tolist()
manufacturers = df['Manufacturer'].values.tolist()
locations = df['Location'].values.tolist()
serials = df['Serial'].values.tolist()

async def run(unique, manufacturer, location, serial, index, sleep):
    index = index + 1
    sleep = sleep * 2

    # launch chromium browser driver
    print("starting web driver #" + str(index) + " for unique #" +str(unique) + "...")
    await asyncio.sleep(sleep)
    async with async_playwright() as driver:
        print("(web driver #" + str(index) + ") opening browser...")
        browser = await driver.chromium.launch(headless=True)
        page = await browser.new_page()

        # if the unit manufacturer is a "Dell" then we go to Dell's warranty page
        if manufacturer.lower() == "Dell".lower():
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

            # close browser driver
            print("(web driver #" + str(index) + ") [Dell] finshed. closing driver...")
            await browser.close()
            print("(web driver #" + str(index) + ") driver closed.")

         # if the unit manufacturer is a "Lenovo" then we go to Lenovo's warranty page
        elif manufacturer.lower() == "Lenovo".lower():
            # need to code!
            print("(web driver #" + str(index) + ") manufacturer 'Lenovo' found, loading warranty page...")

            # close browser driver
            print("(web driver #" + str(index) + ") [Lenovo] finshed. closing driver...")
            await browser.close()
            print("(web driver #" + str(index) + ") driver closed.")

        # if the unit manufacturer is an "HP" then we go to HP's warranty page
        elif manufacturer.lower() == "HP".lower():
            print("(web driver #" + str(index) + ") manufacturer 'HP' found, loading warranty page...")
            await page.goto("https://support.hp.com/us-en/checkwarranty/")
            await page.wait_for_load_state('networkidle')

            # enter the serial number and click on search button
            print("(web driver #" + str(index) + ") [HP] page loaded, looking up serial #" + serial + "...")
            page.on("dialog", lambda dialog: dialog.accept())
            await page.locator('#wFormSerialNumber').type(serial)
            await page.locator('#btnWFormSubmit').click()
            await page.wait_for_load_state('networkidle')
            await asyncio.sleep(8)

            # if the serial number is not valid then throw an error
            if await page.locator('h2:has-text("HP was unable to find a match")').is_visible() == True:
                # throw error
                print("(web driver #" + str(index) + ") [HP] ERROR: unable to find match for serial #" + serial + ".")

                # close browser driver
                print("(web driver #" + str(index) + ") [ERROR:2] error found! closing driver...")
                await browser.close()
                print("(web driver #" + str(index) + ") driver closed.")
                
            # if the serial number is valid then continue to grab warranty information
            else:
                # need to code!
                print("(web driver #" + str(index) + ") [HP] warranty status found () for serial #" + serial + ".")

                # close browser driver
                print("(web driver #" + str(index) + ") [HP] finshed. closing driver...")
                await browser.close()
                print("(web driver #" + str(index) + ") driver closed.")

        # if the unit manufacturer is not a "Dell", "Lenovo" or "HP" then throw an error
        else:
            # throw error
            print("(web driver #" + str(index) + ") [ERROR]: manufacturer not known. (are you sure it's in the correct column?)")

            # close browser driver
            print("(web driver #" + str(index) + ") [ERROR:1] error found! closing driver...")
            await browser.close()
            print("(web driver #" + str(index) + ") driver closed.")

async def main():
    # do you know how we could improve the output of this engine? lets add a turbo!
    turbo = []

    # build a web driver engine by asynchronously creating a list of tasks to run
    for i in range(0, len(uniques)):
            turbo.append(asyncio.create_task(run(uniques[i], manufacturers[i], locations[i], serials[i], i, i)))
    while True:
        turbo = [t for t in turbo if not t.done()]
        if len(turbo) == 0:
            return
        await turbo[0]

asyncio.run(main())
