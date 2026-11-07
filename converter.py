
import argparse
import re
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests


def is_google_sheet_url(value):

    parsed_url = urlparse(value)

    return (
        parsed_url.netloc.lower() in {
            "docs.google.com",
            "sheets.google.com"
        }
        and "/spreadsheets/" in parsed_url.path
    )


def get_google_sheet_id(url):

    match = re.search(
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
        url
    )

    if not match:
        raise ValueError("Invalid Google Sheets URL")

    return match.group(1)


def download_google_sheet(url):

    sheet_id = get_google_sheet_id(url)

    download_url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=xlsx"
    )

    response = requests.get(
        download_url,
        timeout=60
    )

    response.raise_for_status()

    if not response.content.startswith(b"PK"):
        raise ValueError(
            "Unable to download Google Sheet. "
            "Check sharing permissions."
        )

    return BytesIO(response.content), sheet_id


def clean_filename(name):

    cleaned = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        str(name)
    ).strip().rstrip(".")

    return cleaned or "Sheet"


def convert_workbook(workbook, output_dir):

    output_dir = Path(output_dir)

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    excel_file = pd.ExcelFile(
        workbook,
        engine="openpyxl"
    )

    used_names = set()

    print("Total sheets:", len(excel_file.sheet_names))

    for sheet_name in excel_file.sheet_names:

        print("Processing:", sheet_name)

        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=0
        )

        filename = clean_filename(sheet_name)
        candidate = filename
        counter = 2

        while candidate.casefold() in used_names:

            candidate = f"{filename}_{counter}"
            counter += 1

        used_names.add(candidate.casefold())

        output_path = output_dir / f"{candidate}.csv"

        df.to_csv(
            output_path,
            index=False,
            encoding="utf-8-sig"
        )

        print("Saved:", output_path)

    excel_file.close()

    print("All sheets converted successfully!")


def convert_file(input_path):

    input_path = Path(input_path).expanduser()

    if not input_path.exists():
        raise FileNotFoundError(
            f"File not found: {input_path}"
        )

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError("Only XLSX files are supported")

    output_dir = input_path.parent / input_path.stem

    convert_workbook(
        input_path,
        output_dir
    )


def convert_google_sheet(url):

    workbook, sheet_id = download_google_sheet(url)

    output_dir = Path.cwd() / f"GoogleSheet_{sheet_id}"

    convert_workbook(
        workbook,
        output_dir
    )



def main():
    parser = argparse.ArgumentParser(
        description="Convert XLSX files or Google Sheets to CSV"
    )

    parser.add_argument(
        "input",
        help="Local XLSX file path or Google Sheets URL"
    )

    args = parser.parse_args()

    try:
        if is_google_sheet_url(args.input):
            convert_google_sheet(args.input)

        elif args.input.startswith(("http://", "https://")):
            raise ValueError(
                "Invalid URL. Only Google Sheets URLs are supported."
            )

        else:
            convert_file(args.input)

    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()



