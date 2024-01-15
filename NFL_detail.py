import asyncio
import functools
import time
from io import StringIO
import pandas as pd
from pyppeteer import launch
from bs4 import BeautifulSoup
from pyppeteer_stealth import stealth
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

path = '/Users/trishika/PycharmProjects/pythonProject'
#wb = Workbook()


def retry(max_retries, delay=1):
    def decorator(func):
        @functools.wraps(func)
        async def wrapper(*args, **kwargs):
            for _ in range(max_retries):
                try:
                    result = await func(*args, **kwargs)
                    return result
                except Exception as e:
                    print(f"Retrying after error: {e}")
                    time.sleep(delay)
            raise Exception(f"Max retries reached for {func.__name__}")

        return wrapper

    return decorator


async def main():
    await demo3(site_1='https://www.footballdb.com/games/boxscore/detroit-lions-vs-kansas-city-chiefs-2023090701')


@retry(max_retries=3)
async def demo3(site_1, name, _date, wb):
    try:
        browser = await launch({'headless': True})
        page = await browser.newPage()
    except Exception:
        raise Exception('Error launching browser')

    await stealth(page)
    try:
        await page.goto(site_1, {'waitUntil': 'domcontentloaded', 'timeout': 30000})
        html = await page.content()
    except Exception:
        raise Exception('Error loading page')

    soup = BeautifulSoup(html, "html.parser")
    table = soup.find_all('table', class_='statistics')

    #print(table[1])
    t1 = soup.select_one('#divBox_scoring > table > tbody > tr:nth-child(1) > td:nth-child(2)').get_text()
    t2 = soup.select_one('#divBox_scoring > table > tbody > tr:nth-child(1) > td:nth-child(3)').get_text()
    strio_1 = StringIO(str(table[1]))
    df = pd.read_html(strio_1)[0]

    # Create a new Excel workbook and add a worksheet
    match = t1+'-'+t2+'-'+_date.replace('/', '_')
    w = wb.create_sheet(title=match)
    wb.active = w

    # Add the DataFrame to the worksheet
    for row in dataframe_to_rows(df, index=False, header=True):
        w.append(row)
    place = name
    f_path = '/Users/trishika/PycharmProjects/pythonProject/data/output_' + place + '.xlsx'
    wb.save(f_path)
    #await asyncio.sleep(10)
    await browser.close()


if __name__ == '__main__':
    asyncio.run(main())
