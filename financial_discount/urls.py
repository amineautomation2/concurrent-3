from math import ceil
import curl_cffi
import json
from time import sleep
from random import uniform
from utils import get_random_user_agent, save_xlsx


def get_total_funds() -> int:
    headers = {
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


def get_funds_url() -> list[dict]:
    base_url = "https://investorhub.financialexpress.net"
    start = 1
    page_size = 25
    total_funds = get_total_funds()

    print(f"[financial discount] total funds = {total_funds}")
    if total_funds == 0:
        raise Exception(f"financial discount: total_funds = {total_funds}")

    total_pages = ceil(total_funds / page_size)
    end = page_size * start
    funds_url = []
    for current_page in range(start, total_pages+1):
        print(f"page [{current_page}/{total_pages}]")
        rows = get_rows_id(start, end)
        page_funds_json = get_page_data(current_page, page_size, rows)

        for fund in page_funds_json:
            name = fund["FundInfo"]["Name"]
            url = f'{base_url}{fund["FundInfo"]["FactsheetPdfLink"]}'
            # print(fund_data.get("name"))
            funds_url.append(dict(name=name, url=url))
        start += page_size
        end += page_size
        sleep(uniform(1, 2))
    return funds_url


def get_financial_url(xlsx: str) -> None:
    urls = get_funds_url()
    save_xlsx(
        xlsx_out=xlsx,
        funds=urls,
        cols=["name", "isin", "url"],
        sheet="Funds",
    )
