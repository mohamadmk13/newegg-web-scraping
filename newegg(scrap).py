import aiohttp
import asyncio
from bs4 import BeautifulSoup
import xlsxwriter

# Excel
workbook = xlsxwriter.Workbook("newegg.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Brand")
worksheet.write(0, 1, "Color")
worksheet.write(0, 2, "Price")
worksheet.write(0, 3, "CPU")
worksheet.write(0, 4, "Memory")
worksheet.write(0, 5, "SSD")
worksheet.write(0, 6, "External GPU")

# All Variable
row_number = 1
lock = asyncio.Lock()
sem = asyncio.Semaphore(5)
pages = 20

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

# Table
def get_spec(soup, label):
    rows = soup.select(".table-horizontal tr")
    for row in rows:
        th = row.find("th")
        td = row.find("td")
        if th and td and label.lower() in th.text.lower():
            return td.text.strip()
    return ""
# price
def get_price(soup):
    whole = soup.select_one(".price-current strong")
    fraction = soup.select_one(".price-current sup")
    if whole:
        price = whole.text
        if fraction:
            price += fraction.text
        return price.strip()
    return ""

# Product information
async def fetch_product(session, url):
    global row_number

    async with sem:
        async with session.get(url, headers=headers) as response:
            html = await response.text()

    soup = BeautifulSoup(html, "html.parser")

    brand = get_spec(soup, "Brand")
    color = get_spec(soup, "Color")
    cpu = get_spec(soup, "CPU")
    memory = get_spec(soup, "Memory")
    ssd = get_spec(soup, "SSD")
    gpu = get_spec(soup, "Graphics")
    price = get_price(soup)

    async with lock:
        worksheet.write(row_number, 0, brand)
        worksheet.write(row_number, 1, color)
        worksheet.write(row_number, 2, price)
        worksheet.write(row_number, 3, cpu)
        worksheet.write(row_number, 4, memory)
        worksheet.write(row_number, 5, ssd)
        worksheet.write(row_number, 6, gpu)

        row_number += 1

# Body
async def main():
    async with aiohttp.ClientSession() as session:
        search = "laptop"
        number_pages = 1

        while number_pages <= pages:
            url = f"https://www.newegg.com/p/pl?d={search}&page={number_pages}"
            async with session.get(url, headers=headers) as response:
                html = await response.text()

            soup = BeautifulSoup(html, "html.parser")
            links = soup.select("a.item-title")

            tasks = []
            for link in links:
                href = link.get("href")
                if href:
                    tasks.append(fetch_product(session, href))

            await asyncio.gather(*tasks)
            number_pages += 1

# Run
asyncio.run(main())
workbook.close()
print("Finished")
