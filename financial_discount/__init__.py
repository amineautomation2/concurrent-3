from utils import delay, save_xlsx
from worker import get_data_by_worker_id, get_xlsx_data, write_csv_by_id
import io
from pypdf import PdfReader
from re import findall
from utils import get_xlsx_filepath, get_random_user_agent, fetch_with_backoff
from .urls import get_funds_url


def isin_from_text(text: str) -> str:
    isin_pattern = r"[A-Z]{2}[A-Z0-9]{9}[0-9]"
    isin = findall(isin_pattern, text)
    if len(isin) > 0:
        return isin[0]
    return ""


def isin_from_pdf(url: str) -> str:
    cookies = {
        'safariCookie': '1',
        'ASP.NET_SessionId': 'ffhz34mnbxiumr3rlfodsjaw',
    }

    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'en-US,en;q=0.9',
        'cache-control': 'max-age=0',
        'priority': 'u=0, i',
        'sec-ch-ua': '"Not(A:Brand";v="8", "Chromium";v="144"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'none',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
    }

    headers.update(get_random_user_agent())
    isin = ""
    if len(url) == 0:
        return isin

    response = fetch_with_backoff(url, headers=headers, cookies=cookies)
    if response:
        if response.content:
            try:
                pdf_bytes = io.BytesIO(response.content)
                reader = PdfReader(pdf_bytes)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
            except Exception as e:
                print(f"[{url}]isin_from_pdf: ", e)
                return isin
            return isin_from_text(text)
    return isin


def financial_discount_runner(id_worker: int, max_worker: int) -> None:
    out_xlsx = get_xlsx_filepath("financial_discount.xlsx")
    data = get_xlsx_data(out_xlsx, "Funds")
    funds_per_worker = get_data_by_worker_id(id_worker, max_worker, data)
    for fund in funds_per_worker:
        url = fund.get("url")
        if url:
            isin = isin_from_pdf(url)
            fund.update(dict(isin=isin))
        delay(2, 4)

    csv = f"financial_discount_{id_worker}.csv"
    fields = ["index", "name", "isin", "url"]
    write_csv_by_id(csv, funds_per_worker, fields)
