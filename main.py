import asyncio
import time
import itertools
from playwright.async_api import async_playwright
import re
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO
from PIL import Image as PILImage
import httpx


def create_xlsx(data, filename, image_size=(100, 100)):
    wb = openpyxl.Workbook()
    ws = wb.active

    for word, items in data.items():
        # Добавляем слово как заголовок
        ws.append([word])
        ws.merge_cells(start_row=ws.max_row, start_column=1,
                       end_row=ws.max_row, end_column=3)
        ws.cell(row=ws.max_row,
                column=1).font = openpyxl.styles.Font(size=20, bold=True)

        for item in items:
            row = ["", item["index"], item["listing_id"]]
            ws.append(row)

            img_data = BytesIO(item["image_data"])
            pil_img = PILImage.open(img_data)
            pil_img = pil_img.resize(image_size)

            # Преобразуем PIL изображение в объект BytesIO
            img_byte_arr = BytesIO()
            pil_img.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)

            # Добавляем измененное изображение в Excel
            img = ExcelImage(img_byte_arr)
            img.anchor = f"A{ws.max_row}"
            ws.add_image(img)

            # Подгоняем высоту строки под изображение
            image_height = img.height
            row_height = image_height * 0.75
            ws.row_dimensions[ws.max_row].height = row_height

            column = list(ws.columns)[0][0].column_letter
            ws.column_dimensions[column].width = image_size[0] * 0.075 * 2
        ws.append([])

    wb.save(filename)


MAIN_URL = "https://www.etsy.com"


class EtsyClient():
    @staticmethod
    def get_images_links(srcset):
        if not srcset:
            return {}

        pattern = re.compile(r'(\d+)w')
        image_links = {}
        for entry in srcset.split(','):
            url, width = entry.strip().rsplit(' ', 1)
            match = pattern.search(width)
            if match:
                image_links[match.group(1)] = url

        return image_links

    async def get_cookie_header(self):
        cookies = await self._context.cookies()

        cookies_header = "; ".join([f"{cookie['name']}={cookie['value']}"
                                    for cookie in cookies])
        return cookies_header

    def __init__(self, shop_name, language, currency, image_size,
                 region, max_page, file, req_cooldown):
        self._max_page = max_page
        self._req_cooldown = req_cooldown
        self._shop_name = shop_name
        self._language = language
        self._region = region
        self._currency = currency
        self._image_size = image_size

        try:
            with open(file, "r") as f:
                self._words = f.readlines()
        except Exception:
            raise Exception(f"Во время загрузки файла {file} произошла ошибка")

    async def init_browser(self):
        playwright = await async_playwright().start()
        browser = await playwright.chromium.launch(headless=True)

        headers = {
            "Accept": ("text/html,application/xhtml+xml,application/xml;"
                       "q=0.9,image/avif,image/webp,image/apng,*/*;"
                       "q=0.8,application/signed-exchange;v=b3;q=0.7"),
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "ru",
            "Cache-Control": "max-age=0",
            "Downlink": "8.25",
            "Dpr": "1.25",
            "Ect": "4g",
            "Priority": "u=0, i",
            "Referer": "https://www.etsy.com/?dd_referrer=",
            "Origin": "https://www.etsy.com",
            "Rtt": "100",
            "Sec-Ch-Dpr": "1.25",
            "Sec-Ch-Ua": ('"Microsoft Edge";v="131", '
                          '"Chromium";v="131", "Not_A Brand";v="24"'),
            "Sec-Ch-Ua-Arch": '"x86"',
            "Sec-Ch-Ua-Bitness": '"64"',
            "Sec-Ch-Ua-Full-Version-List": ('"Microsoft Edge";'
                                            'v="131.0.2903.99", '
                                            '"Chromium";'
                                            'v="131.0.6778.140", '
                                            '"Not_A Brand";v="24.0.0.0"'),
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Ch-Ua-Platform-Version": '"15.0.0"',
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
        }

        self._context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 "
                "Edg/131.0.0.0"
            ),
            extra_http_headers=headers
        )

        page = await self._context.new_page()
        await page.goto(MAIN_URL)
        await page.wait_for_load_state('networkidle')
        await page.close()

        page = await self._context.new_page()
        await page.goto(MAIN_URL)
        await page.wait_for_load_state('networkidle')
        content = await page.content()

        regexp_csrf = (
            r'<meta name=["\']csrf_nonce["\'] '
            r'content=["\'](.+?)["\']'
        )
        match = re.search(regexp_csrf, content)
        if match:
            self._csrf_nonce = match.group(1)

        script_content = await page.locator("script").all_inner_texts()
        page_guid_match = None
        regexp_guid = r'page_guid"\s*[:=]\s*["\'](.+?)["\']'

        for script in script_content:
            page_guid_match = re.search(regexp_guid, script)
            if page_guid_match:
                self._page_guid = page_guid_match.group(1)
                break

        self._api_headers = {
            **headers,
            "accept": "*/*",
            "referer": f"{MAIN_URL}/search?q=pyjamas&page=2"
                       "&ref=pagination",
            "x-csrf-token": self._csrf_nonce,
            "x-page-guid": self._page_guid,
            "X-Detected-Locale": "USD|ru|UA",
            "X-Etsy-Protection": "1",
            "X-Requested-With": "XMLHttpRequest",
            "content-type": "application/json",
            "cookie": await self.get_cookie_header(),
        }

        api = self._context.request
        await api.post(f"{MAIN_URL}/api/v3/ajax/member/locale-preferences",
                       headers=self._api_headers, data={
                            "currency": self._currency,
                            "language": self._language,
                            "region": self._region
                        })

    async def get_data_from_page(self, page, word):
        locale_header = f"{self._currency}|{self._language}|{self._region}"
        self._api_headers = {
                **self._api_headers,
                "X-Detected-Locale": locale_header,
                "cookie": await self.get_cookie_header()}

        payload = {
            "log_performance_metrics": True,
            "specs": {
                "async_search_results": [
                    "Search2_ApiSpecs_WebSearch",
                    {
                        "search_request_params": {
                            "detected_locale": {
                                "language": "ru",
                                "currency_code": self._currency,
                                "region": self._region
                            },
                            "locale": {
                                "language": self._language,
                                "currency_code": self._currency,
                                "region": self._region
                            },
                            "name_map": {
                                "query": "q",
                                "query_type": "qt",
                                "results_per_page": "result_count",
                                "min_price": "min",
                                "max_price": "max"
                            },
                            "parameters": {
                                "q": word,
                                "page": page,
                                "ref": "pagination",
                                "referrer": (f"{MAIN_URL}/uk/search?q={word}&page={page}"
                                             "&ref=search_bar"),
                                "is_prefetch": True,
                                "placement": "wsg"
                            },
                            "user_id": None
                        },
                        "request_type": "pagination_preact"
                    }
                ]
            },
            "view_data_event_name": "search_single_page_app_specview_rendered",
            "runtime_analysis": False
        }

        api = self._context.request
        api_path = "/api/v3/ajax/bespoke/member/neu/specs/async_search_results"
        api_url = MAIN_URL + api_path
        response = await api.post(
            api_url,
            data=payload,
            headers=self._api_headers,

        )

        data = await response.json()
        html_data = data["output"]["async_search_results"]
        soup = BeautifulSoup(html_data, "html.parser")

        listings = soup.select("a.listing-link")
        if page == 1:
            listings = listings[:-6]

        listings_filtered = list(filter(
            lambda listing: listing.text.find(self._shop_name) != -1,
            listings))

        listings_data = []

        client = httpx.AsyncClient()

        async def get_listing_data(listing):
            num = listings.index(listing)
            row = num // 4 + 1
            listing_id = listing["data-listing-id"]
            image = listing.select_one("img.wt-image")

            if not image:
                return {}

            links = self.get_images_links(image["srcset"])
            image_link = links.get(self._image_size, list(links.values())[0])
            index = f"{page}.{row:02d}"

            image_data = await client.get(image_link)

            return {
                "image_data": image_data.content,
                "index": index,
                "listing_id": listing_id,
            }

        listings_data = await asyncio.gather(*[get_listing_data(listing)
                                             for listing in listings_filtered])

        return listings_data

    async def get_data_full(self):
        full_data = {}

        for word in self._words:
            tasks = [self.get_data_from_page(page, word)
                     for page in range(1, self._max_page+1)]

            word_data = list(itertools.chain(*await asyncio.gather(*tasks)))

            full_data = {
                **full_data,
                word: word_data
            }

        return full_data

    async def close_client(self):
        await self._context.close()


async def main():
    t = time.time()
    etsy = EtsyClient("IDlingerieUK", "en-GB", "GBP", '300',
                      "GB", 3, "test.txt", 10)
    await etsy.init_browser()
    data = await etsy.get_data_full()
    await etsy.close_client()
    create_xlsx(data, "test.xlsx")
    print(len(data), time.time()-t)

asyncio.run(main())
