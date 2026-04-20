import argparse
import time
from financial_discount import financial_discount_runner, get_financial_url
from utils import clean_spreadsheet, delay, get_xlsx_filepath
from worker import (
    merge_csv_to_xlsx,
)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--id", type=str, help="id worker")
    parser.add_argument("--save", action='store_true', help="sheet name")
    parser.add_argument("--url", action='store_true', help="sheet name")

    args = parser.parse_args()
    xlsx_out = get_xlsx_filepath("financial_discount.xlsx")
    if args.url:
        clean_spreadsheet(xlsx_out)
        get_financial_url(xlsx_out)

    elif args.id:
        # id_worker = int(sys.argv[1])
        start = time.perf_counter()
        financial_discount_runner(id_worker=int(args.id), max_worker=5)
        elapsed = time.perf_counter() - start
        print(f"Execution time: {elapsed:.2f} seconds.")
        return

    elif args.save:
        merge_csv_to_xlsx(
            xlsx_out, ["name", "isin", "url"], "Funds")
        return


if __name__ == "__main__":
    main()
