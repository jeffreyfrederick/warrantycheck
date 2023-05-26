import asyncio
import pandas as pd
from httpx import AsyncClient
from datetime import timedelta
from aiolimiter import AsyncLimiter
from timeit import default_timer as timer
from tenacity import retry
from playwright.async_api import async_playwright

unique_list = []
serial_list = []
oem_list = []
model_list = []
location_list = []
expiration_list = []
current_plan_list = []
popup_count = 0

# open Excel file and convert it to a data frame
df = pd.read_excel("input.xlsx")

# change data types to lists for iteration
uniques = df['Unique #'].values.tolist()
serials = df['Serial #'].values.tolist()
oem = df['OEM'].values.tolist()
models = df['Model'].values.tolist()
locations = df['Location'].values.tolist()

@retry
async def scrape(index, unique, serial, oem, model, location, throttler):
    index += 1
    global popup_count
    global serials
    
    async with throttler:
        # launch browser driver
        print(f"starting web driver #{index}")
        async with async_playwright() as driver:
            print(f"(web driver #{index}) opening browser...")
            browser = await driver.firefox.launch(headless=True)
            context = await browser.new_context(extra_http_headers={"Accept Language": "en-US",})

            # load Dell's warranty page
            print(f"(web driver #{index}) loading warranty page...")
            page = await context.new_page()
            await page.goto("https://www.dell.com/support/home/en-us?app=warranty/")
            await page.wait_for_load_state()

            # enter the serial number and click on search button
            print(f"(web driver #{index}) looking up serial #{serial}...")
            page.on("dialog", lambda dialog: dialog.accept())
            await page.locator('#inpEntrySelection').type(serial)
            await page.locator('#btn-entry-select').click()
            await asyncio.sleep(2)

            # Dell has some kind of bot detection so have to wait for a timer to end
            if await page.query_selector('#btn-entry-select'):
                print(f"(web driver #{index}) anti-bot detected, waiting...")
                await asyncio.sleep(30)
                await page.locator('#btn-entry-select').click()

            # Dell has a survey popup that blocks viewing the warranty information, close the popup
            await page.wait_for_load_state()
            await asyncio.sleep(8)   
            try:
                close_button = page.frame_locator("iframe#iframeSurvey").locator("#noButtonIPDell")
                await close_button.click(timeout=1000)
                print(f"(web driver #{index}) \033[31m\033[01m*** POPUP BLOCKED! ***\033[0m")
                popup_count += 1
            except:
                print(f"(web driver #{index}) popup not found, continuing...")

            # click on product details to view warranty information
            await asyncio.sleep(8)
            await page.locator('#viewDetailsWarranty').click()

            # grab warranty expiration date from product details and remove any special characters
            expiration_date = await page.locator('#dsk-expirationDt').inner_text()
            expiration_date = expiration_date.translate({ord(i): None for i in "!@#$%^&*()_+-=[]\{}|;':,/.\""})

            # grab warranty plan type
            if await page.query_selector('#supp-svc-status-txt-2'):
                plan_locator = ('#supp-svc-status-txt-2')
            else:
                plan_locator = ('#supp-svc-status-txt')

            plan = await page.locator(plan_locator).inner_text()
            plan_status = plan.split(':')[1].strip()

            if "Active" in plan_status:
                current_plan_locator = '#supp-svc-plan-txt-2'
                current_plan = await page.locator(current_plan_locator).inner_text()
                plan_type = current_plan.split(':')[1].strip()
                current_plan_list.append(plan_type)
            elif "Expires" in plan_status:
                if await page.query_selector('#supp-svc-plan-txt-2'):
                    current_plan = await page.locator('#supp-svc-plan-txt-2').inner_text()
                    plan_type = current_plan.split(':')[1].strip()
                    plan_status = "Active"
                    current_plan_list.append(plan_type)
                else:
                    plan_status = "Active"
                    current_plan_list.append('N/A')
            else:
                plan_type = "Expired"
                current_plan_list.append('N/A')

            print(f"(web driver #{index}) warranty status found (\033[32m{expiration_date}\033[0m) for serial #{serial}")

            # append expiration date to list for saving to Excel sheet
            print(f"(web driver #{index}) appending '{expiration_date}' to excel sheet for unique #{unique}")
            unique_list.append(unique)
            serial_list.append(serial)
            oem_list.append(oem)
            model_list.append(model)
            expiration_list.append(expiration_date)
            location_list.append(location)
            print(expiration_list)
            print(current_plan_list)

            # close browser driver
            print(f"(web driver #{index}) finished. closing driver...")
            print(f"(web driver #{index}) serials \033[36m{len(serials)}\033[0m dates \033[32m{len(expiration_list)}\033[0m popups \033[31m{popup_count}\033[0m")
            await browser.close()
            print(f"driver #{index} closed")

async def run():
    # start timer and run engine
    start = timer()
    unique_length = (len(uniques))
    print(f"looking up \033[36m{unique_length}\033[0m uniques...")

    # slow down tasks using a limiter so we don't overload our system  (8 max firefox browsers open at once)
    throttler = AsyncLimiter(max_rate=1, time_period=10)

    # start loop
    async with AsyncClient() as session:
        tasks = []
        for i in range(0, unique_length):
            tasks.append(scrape(i, uniques[i], serials[i], oem[i], models[i], locations[i], throttler=throttler))
        results = await asyncio.gather(*tasks)

    # add expiration list to data frame column
    data = {
    'Unique #': unique_list,
    'Serial #': serial_list,
    'OEM': oem_list,
    'Model': model_list,
    'Expiration': expiration_list,
    'Plan': current_plan_list,
    'Location': location_list
    }
    ndf = pd.DataFrame(data)

    # output results to excel file
    ndf.to_excel('output.xlsx')

    #end timer
    end = timer()
    print(f"execution time: {timedelta(seconds=end-start)}")

if __name__ == "__main__":
    # lets go!
    asyncio.run(run())