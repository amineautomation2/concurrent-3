from openpyxl import load_workbook
from pypdf.errors import EmptyFileError
from pypdf import PdfReader
from re import findall
import utils
from random import uniform
from os.path import join, dirname, basename
from pathlib import Path
import openpyxl
from utils import fetch_with_backoff
from selenium import webdriver
import base64
from time import sleep
from utils import setup_driver
from pprint import pprint
import json
from utils import get_xlsx_filepath
from os.path import join
from os import chdir
from os.path import dirname
import requests
from utils import get_random_user_agent
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


def create_mutual_fund_spreadsheet(data, filepath):
    """
    Create Excel spreadsheet with Mutual Fund data.

    Args:
        data: List of dictionaries containing Name, Open, and ISIN
        filepath: Output file path for the Excel file
    """
    # Create new workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Mutual Fund"

    # Style the header row
    header_font = Font(bold=True, italic=True, size=13, color='000000')
    header_fill = PatternFill(start_color='4472C4',
                              end_color='4472C4', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center')

    # Write headers
    headers = ['Name', 'Open', 'ISIN']
    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # Write data rows
    # for row in data:
    #    sheet.append([
    #        row.get('Name'),
    #        row.get('Open'),
    #        row.get('ISIN')
    #    ])

    for column in sheet.columns:  # type: ignore
        max_length = 0
        column_letter = column[0].column_letter  # type: ignore

        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        adjusted_width = max_length + 2
        # type: ignore
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Save workbook
    workbook.save(filepath)
    workbook.close()


# mutual_funds = [
#    {'Name': 'Vanguard 500 Index', 'Open': 425.32, 'ISIN': 'US9229083632'},
#    {'Name': 'Fidelity Total Market', 'Open': 115.67, 'ISIN': 'US3160928030'},
#    {'Name': 'PIMCO Income Fund', 'Open': 11.85, 'ISIN': 'US72201R8824'}
# ]

# create_mutual_fund_spreadsheet(mutual_funds, 'mutual_funds.xlsx')
# print('Spreadsheet created successfully!')


# mylogger = logging.getLogger(__name__)
# logging_config = {
#    "version": 1,
#    "disable_existing_loggers": False,
#    "formatters": {
#        "simple": {
#            "format": "%(levelname)s: %(message)s",
#        }
#    },
#    "handlers": {
#        "stderr": {
#            "class": "logging.StreamHandler",
#            "level": "WARNING",
#            "formatter": "simple",
#            "stream": "ext://sys.stderr",
#        },
#        "file": {
#            "class": "logging.handlers.RotatingFileHandler",
#            "level": "DEBUG",
#            "formatter": "simple",
#            "filename": "logs/app.log",
#            "maxBytes": 10000,
#            "backupCount": 3,
#        }
#    },
#    "loggers": {
#        "root": {
#            "level": "DEBUG",
#            "handlers": ["stdout"]
#        }
#    },
# }
#
# logging.config.dictConfig(config=logging_config)
# mylogger.addHandler(logging.StreamHandler())
# mylogger.debug("debug message")
# mylogger.info("info message")
# mylogger.warning("warning message")
# mylogger.error("error message")
# mylogger.critical("critical message")
#
#
# def calc():
#    try:
#        s = 1 / 0
#    except:
#        raise Exception(ZeroDivisionError)
#




url = "https://investorhub.financialexpress.net/Pdf/fdd/en-gb/fdd/BUH7/U/?specialunittype=ORDN&priipproductcode="


def download_pdf(url: str, name: str) -> int:
    response = fetch_with_backoff(url)
    if response:
        with open(f"{name}.pdf", "wb") as f:
            f.write(response.content)
            return 0
    return 1


def missing_funds() -> list[dict]:
    path = get_xlsx_filepath("financial_discount.xlsx")
    wb = openpyxl.load_workbook(path)
    ws = wb["Funds"]

    data = []
    i = 2
    for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
        f = {}
        if row[1].value:
            i += 1
            continue
        f["name"] = row[0].value
        f["isin"] = row[1].value
        f["url"] = row[2].value
        f["row_index"] = i
        i += 1
        data.append(f)
    return data


def retry_missing_funds():
    data = missing_funds()
    new_data = []
    new_fund = {}
    # project_root = Path(__file__).resolve().parent.parent
    project_root = Path(__file__).resolve().parent
    i = 0
    for fund in data:
        print(f"downloading [{i+1}/{len(data)}]")
        p = join(project_root, "download", f"{fund['row_index']}")
        err = download_pdf(fund["url"], p)
        sleep(uniform(2, 5))
        new_fund.update(fund)
        new_fund.update({"error": err})
        new_data.append(new_fund)
        i += 1
    return new_data


# data = retry_missing_funds()

# utils.write_json("financial_missing.json", data)


def get_pdf_files(folder_name):
    folder_path = Path(folder_name)
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_name}' does not exist")
    if not folder_path.is_dir():
        raise NotADirectoryError(f"'{folder_name}' is not a directory")
    pdf_files = []
    for file in folder_path.iterdir():
        if file.is_file() and file.suffix.lower() == '.pdf':
            pdf_files.append(file.name)
    return pdf_files


def isin_from_text(text: str) -> str:
    isin_pattern = r"[A-Z]{2}[A-Z0-9]{9}[0-9]"
    isin = findall(isin_pattern, text)
    if len(isin) > 0:
        return isin[0]
    return ""


def isin_from_pdf(file: str) -> str | None:
    try:
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return isin_from_text(text)
    except EmptyFileError:
        return None


def get_missing_isins():
    data = get_pdf_files("download")
    isins = []
    i = 0
    for pdf in data:
        print(f"[{i+1}/{len(data)}] reading {pdf} ")
        p = join("download", pdf)
        index = int(pdf[:len(pdf) - 4])
        isin = isin_from_pdf(p)
        isins.append({"index": index, "filename": pdf, "isin": isin})
        i += 1
    utils.write_json("financial_isins.json", isins)
    fix_missing_isins(isins)


def fix_missing_isins(data: list[dict]):
    path = utils.get_xlsx_filepath("financial_discount.xlsx")
    wb = load_workbook(path)
    ws = wb["Funds"]
    for fund in data:
        ws.cell(fund["index"], 2, fund["isin"])

    wb.save(path)
    wb.close()
    print(path)
    return


def count_empty_isins(filename: str, sheet: str) -> int:
    path = utils.get_xlsx_filepath(filename)
    wb = load_workbook(path)
    ws = wb[sheet]
    empty_count = 0
    for row in ws.iter_rows(min_row=1, values_only=False):
        cell_b = row[1]  # Column B is index 1
        if cell_b.value is None or str(cell_b.value).strip() == '':
            empty_count += 1
    wb.close()
    return empty_count
