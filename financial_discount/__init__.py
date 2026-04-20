from math import ceil
import curl_cffi
import json
import io
import openpyxl
from time import sleep
from random import uniform
from pypdf import PdfReader
from re import findall
from utils import clean_spreadsheet, get_xlsx_filepath, get_random_user_agent, fetch_with_backoff


def isin_from_text(text: str) -> str:
    isin_pattern = r"[A-Z]{2}[A-Z0-9]{9}[0-9]"
    isin = findall(isin_pattern, text)
    if len(isin) > 0:
        return isin[0]
    return ""


def isin_from_pdf(url: str) -> str:
    if len(url) == 0:
        return ""
    response = fetch_with_backoff(url)
    if response:
        if not response.content:
            try:
                pdf_bytes = io.BytesIO(response.content)
                reader = PdfReader(pdf_bytes)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
            except Exception as e:
                print(f"[{url}]isin_from_pdf: ", e)
                return ""

            return isin_from_text(text)
    return ""


def get_total_funds() -> int:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0',
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        # 'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Origin': 'https://investorhub.financialexpress.net',
        'Sec-GPC': '1',
        'Connection': 'keep-alive',
        'Referer': 'https://investorhub.financialexpress.net/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'cross-site',
        # Requests doesn't support trailers
        # 'TE': 'trailers',
    }

    ua = get_random_user_agent()
    headers.update(ua)
    params = {
        'jsonString': '{"FilteringOptions":{"undefined":0,"RangeId":null,"RangeName":"","CategoryId":null,"Category2Id":null,"PriipProductCode":null,"DefaultCategoryId":null,"DefaultCategory2Id":null,"ForSaleIn":null,"ShowMainUnits":false,"MPCategoryCode":null},"ProjectName":"fdd","LanguageCode":"en-gb","UserType":"","Region":"","LanguageId":"1","LocaleId":"1","Theme":"fdd","SortingStyle":"1","PageNo":1,"PageSize":25,"OrderBy":"UnitName:init","IsAscOrder":true,"OverrideDocumentCountryCode":null,"ToolId":"1","PrefetchPages":80,"PrefetchPageStart":1,"OverridenThemeName":"fdd","ForSaleIn":"","ValidateFeResearchAccess":false,"HasFeResearchFullAccess":false,"EnableSedolSearch":"false","GrsProjectId":"17200043","ShowMainUnitExpansion":false,"UseCombinedOngoingChargeTER":false}',
    }

    response = curl_cffi.get(
        'https://digitalfundservice.feprecisionplus.com/FundDataService.svc/GetRowIdList',
        params=params,
        headers=headers,
        impersonate="chrome"
    )
    if response.status_code != 200:
        return 0

    total = response.json()["TotalRows"]
    return int(total)


def get_rows_id(begin: int, end: int) -> str:
    rows = []
    for i in range(begin, end + 1):
        rows.append(f"{i}")
    return ",".join(rows)


def get_page_data(nb_page: int, page_size: int, rows_id: str) -> list[dict]:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:148.0) Gecko/20100101 Firefox/148.0',
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        # 'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Origin': 'https://investorhub.financialexpress.net',
        'Sec-GPC': '1',
        'Connection': 'keep-alive',
        'Referer': 'https://investorhub.financialexpress.net/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'cross-site',
        'Priority': 'u=0',
    }

    ua = get_random_user_agent()
    headers.update(ua)

    params = '{"FilteringOptions":{"undefined":0,"RangeId":null,"RangeName":"","CategoryId":null,"Category2Id":null,"PriipProductCode":null,"DefaultCategoryId":null,"DefaultCategory2Id":null,"ForSaleIn":null,"ShowMainUnits":false,"MPCategoryCode":null},"ProjectName":"fdd","LanguageCode":"en-gb","UserType":"","Region":"","LanguageId":"1","LocaleId":"1","Theme":"fdd","SortingStyle":"1","PageNo":2,"PageSize":25,"OrderBy":"UnitName:init","IsAscOrder":true,"OverrideDocumentCountryCode":null,"ToolId":"1","PrefetchPages":80,"PrefetchPageStart":1,"OverridenThemeName":"fdd","ForSaleIn":"","ValidateFeResearchAccess":false,"HasFeResearchFullAccess":false,"EnableSedolSearch":"false","RowCount":3600,"RowIDs":"26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50"}'

    payload = json.loads(params)
    pageInfo = {
        "PageNo": nb_page,
        "PageSize": page_size,
        "RowIDs": f"{rows_id}",
    }
    payload.update(pageInfo)
    params = {
        'jsonString': json.dumps(payload),
    }
    response = curl_cffi.get(
        'https://digitalfundservice.feprecisionplus.com/FundDataService.svc/GetUnitList',
        params=params,
        headers=headers,
        impersonate='chrome'
    )
    if response.status_code != 200:
        print(response)
        return []

    data = response.json()
    decode = json.loads(data)
    return decode["DataList"]
    # with open(f"financial_{nb_page}.json", "w+") as f:
    #    json.dump(decode, f, indent=4)


def extract_data(fund: dict) -> dict:
    base_url = "https://investorhub.financialexpress.net"
    name = fund["FundInfo"]["Name"]
    url = f'{base_url}{fund["FundInfo"]["FactsheetPdfLink"]}'
    doc = fund["Documents"]
    isin_text = doc.get("AdditionalInformationtoInvestors")
    if isin_text:
        return dict(name=name,  url=url, isin=isin_from_text(isin_text))
    return dict(name=name,  url=url)


def extract_isin(fund: dict) -> str | None:
    url = fund.get("url")
    if url:
        return isin_from_pdf(url)


def financial_discount_runner(start: int = 1) -> None:
    out_xlsx = get_xlsx_filepath("financial_discount.xlsx")
    clean_spreadsheet(out_xlsx)
    page_size = 25
    total_funds = get_total_funds()

    print(f"[financial discount] total funds = {total_funds}")
    if total_funds == 0:
        raise Exception(f"financial discount: total_funds = {total_funds}")

    total_pages = ceil(total_funds / page_size)
    end = page_size * start
    funds_arr = []
    for current_page in range(start, total_pages+1):
        print(f"page [{current_page}/{total_pages}]")
        rows = get_rows_id(start, end)
        print(rows)
        funds_json = get_page_data(current_page, page_size, rows)

        page_data = []
        for fund in funds_json:
            fund_data = extract_data(fund)
            # print(fund_data.get("name"))
            page_data.append(fund_data)
            funds_arr.append(fund_data)
        start += page_size
        end += page_size
        sleep(uniform(1, 2))
    financial_write_xlsx(out_xlsx, funds_arr)


def financial_write_xlsx(file_xlsx: str, data: list[dict]):
    wb = openpyxl.load_workbook(file_xlsx)
    ws = wb["Funds"]
    final = []

    for fund in data:
        if not fund.get("isin"):
            url = fund.get("url")
            if url:
                print(f"extracting from {url}")
                isin = isin_from_pdf(url)
                fund.update(dict(isin=isin))
        final.append(fund)
        sleep(uniform(3, 10))

    i = 2
    for row in final:
        ws.cell(i, 1, row["name"])
        ws.cell(i, 2, row["isin"])
        c = ws.cell(i, 3, row["url"])
        c.hyperlink = row["url"]
        c.style = "Hyperlink"
        i += 1
    wb.save(file_xlsx)

    for sheet in wb:
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # type: ignore

            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            adjusted_width = max_length + 2
            sheet.column_dimensions[column_letter].width = adjusted_width
    wb.save(file_xlsx)
    wb.close()
