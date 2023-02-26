import asyncio
import pandas as pd
from httpx import AsyncClient
from datetime import timedelta
from aiolimiter import AsyncLimiter
from timeit import default_timer as timer
from playwright.async_api import async_playwright

expiration_list = []

# open Excel file and convert it to a data frame
df = pd.read_excel("input.xlsx")

# change data types to lists for iteration
uniques = df['Unique #'].values.tolist()
serials = df['Serial #'].values.tolist()

async def scrape(index, unique, serial, throttler):
    index += 1
    async with throttler:

        # launch Chromium browser driver
        print(f"starting web driver #{index}")
        async with async_playwright() as driver:
            print(f"(web driver #{index}) opening browser...")
            browser = await driver.firefox.launch(headless=True)

            # load Dell's warranty page
            print(f"(web driver #{index}) [Dell] loading warranty page...")
            page = await browser.new_page()
            await page.goto("https://www.dell.com/support/home/en-us?app=warranty/")
            await page.wait_for_load_state()

            # enter the serial number and click on search button
            print(f"(web driver #{index}) [Dell] looking up serial #{serial}...")
            page.on("dialog", lambda dialog: dialog.accept())
            await page.locator('#inpEntrySelection').type(serial)
            await asyncio.sleep(.5)
            await page.locator('#btn-entry-select').click()

            # Dell has some kind of bot detection so have to wait for a timer to end
            await asyncio.sleep(32)

            # click on search button again
            await page.locator('#btn-entry-select').click()

            # Dell has a survey popup that blocks viewing the warranty information, close the popup
            print(f"(web driver #{index}) [Dell] checking for survey popup...")
            await asyncio.sleep(4)
            try:
                close_button = page.frame_locator("iframe#iframeSurvey").locator("#noButtonIPDell")
                await close_button.click(timeout=1000)
                print(f"(web driver #{index}) [Dell] survey popup closed!")
            except:
                print(f"(web driver #{index}) [Dell] survey popup not found.")

            # click on product details to view warranty information
            await asyncio.sleep(.5)
            await page.locator('#viewDetailsWarranty').click()
            await asyncio.sleep(.5)

            # grab warranty expiration date from product details
            expiration_date = await page.locator('#dsk-expirationDt').inner_text()
            print(f"(web driver #{index}) [Dell] warranty status found ({expiration_date}) for serial #{serial}.")

            # append expiration date to list for saving to Excel sheet
            print(f"(web driver #{index}) [Dell] appending '{expiration_date}' to excel sheet for unique #{unique}...")  
            expiration_list.append(expiration_date)
            print(expiration_list)

            # close browser driver
            print(f"(web driver #{index}) [Dell] finished. closing driver...")
            await browser.close()
            print(f"driver #{index} closed.")

async def run():
    # start timer and run engine
    start = timer()
    unique_length = (len(uniques))
    print(f"looking up {unique_length} uniques...")

    # slow down tasks using a limiter so we don't overload our system  (8 max chromium browsers open at once)
    throttler = AsyncLimiter(max_rate=1, time_period=10)

    # start loop
    async with AsyncClient() as session:
        tasks = []
        for i in range(0, unique_length):
            tasks.append(scrape(i, uniques[i], serials[i], throttler=throttler))
        results = await asyncio.gather(*tasks)

    # add expiration list to data frame column
    df.insert(5, 'Expiration', expiration_list)

    # output results to excel file
    df.to_excel('output.xlsx')

    #end timer
    end = timer()
    print(f"execution time: {timedelta(seconds=end-start)}")

if __name__ == "__main__":
    asyncio.run(run())
