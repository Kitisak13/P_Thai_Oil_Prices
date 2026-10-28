import argparse
from datetime import date, datetime, timedelta
import logging
from pathlib import Path
import re
import time
from typing import Optional

from bs4 import BeautifulSoup
import requests

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

BASE_DIR = Path(__file__).resolve().parent
LEGACY_DATA_DIR = BASE_DIR / 'Data'
DEFAULT_RAW_DIR = BASE_DIR / 'data' / 'raw'
EPPO_URL = 'https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx'


def get_default_output_dir() -> Path:
    """Return default output directory, keeping compatibility with legacy Data/ if used."""
    if DEFAULT_RAW_DIR.exists():
        return DEFAULT_RAW_DIR
    if LEGACY_DATA_DIR.exists():
        return LEGACY_DATA_DIR
    return DEFAULT_RAW_DIR


def extract_tokens(html_content: str) -> dict[str, str]:
    """Extract ASP.NET form state tokens from HTML page."""
    soup = BeautifulSoup(html_content, 'html.parser')
    viewstate = soup.find(id='__VIEWSTATE')
    viewstategen = soup.find(id='__VIEWSTATEGENERATOR')
    eventval = soup.find(id='__EVENTVALIDATION')

    if not viewstate or not viewstate.get('value'):
        raise ValueError("Could not find '__VIEWSTATE' in HTML response.")

    return {
        '__VIEWSTATE': viewstate['value'],
        '__VIEWSTATEGENERATOR': viewstategen['value'] if viewstategen else '',
        '__EVENTVALIDATION': eventval['value'] if eventval else '',
    }


def find_latest_downloaded_date(search_dirs: list[Path]) -> Optional[date]:
    """Find the latest date among already downloaded Excel files in given directories."""
    latest_dt: Optional[date] = None
    pattern = re.compile(r'_on_(\d{4})(\d{2})(\d{2})\.xls$', re.IGNORECASE)

    for directory in search_dirs:
        if not directory.exists():
            continue
        for file in directory.glob('*.xls'):
            match = pattern.search(file.name)
            if match:
                try:
                    d = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
                    if latest_dt is None or d > latest_dt:
                        latest_dt = d
                except ValueError:
                    continue
    return latest_dt


def download_eppo_oil_prices(
    start_date: date,
    end_date: date,
    output_dir: Path,
    timeout: int = 20,
    delay: float = 1.0,
    max_retries: int = 3,
) -> None:
    """Download daily oil price Excel reports from EPPO within the date range."""
    output_dir.mkdir(parents=True, exist_ok=True)

    with requests.Session() as session:
        session.headers.update({
            'User-Agent': (
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
            )
        })

        logging.info("Requesting initial page to obtain ASP.NET tokens...")
        try:
            res = session.get(EPPO_URL, timeout=timeout)
            res.raise_for_status()
            tokens = extract_tokens(res.text)
            logging.info("ASP.NET tokens successfully extracted.")
        except Exception as e:
            logging.error(f"Failed to initialize session from EPPO: {e}")
            return

        current_date = start_date
        success_count = 0
        skip_count = 0
        no_data_count = 0
        fail_count = 0

        while current_date <= end_date:
            date_str = current_date.strftime('%d/%m/%Y')
            filename = f"EPPO_RetailOilPrice_on_{current_date.strftime('%Y%m%d')}.xls"
            filepath = output_dir / filename

            # Check existing file (both in output_dir and legacy data dir)
            legacy_file = LEGACY_DATA_DIR / filename
            if (filepath.exists() and filepath.stat().st_size > 0) or (
                legacy_file.exists() and legacy_file.stat().st_size > 0
            ):
                logging.info(f"File {filename} already exists. Skipping...")
                skip_count += 1
                current_date += timedelta(days=1)
                continue

            logging.info(f"Downloading data for {date_str} ({current_date.strftime('%Y-%m-%d')})...")
            payload = {
                **tokens,
                'TbxToDate': date_str,
                'BtnGenerate.x': '10',
                'BtnGenerate.y': '10',
            }

            downloaded = False
            has_no_data = False
            request_error = False

            for attempt in range(1, max_retries + 1):
                try:
                    post_res = session.post(EPPO_URL, data=payload, timeout=timeout)
                    post_res.raise_for_status()

                    content_type = post_res.headers.get('Content-Type', '')
                    content_disp = post_res.headers.get('Content-Disposition', '')
                    is_excel_content = post_res.content.startswith(b'\xd0\xcf\x11\xe0') or (
                        len(post_res.content) > 1000 and b'Workbook' in post_res.content
                    )

                    if (
                        'application/vnd.ms-excel' in content_type
                        or content_disp
                        or is_excel_content
                    ):
                        filepath.write_bytes(post_res.content)
                        logging.info(f"Successfully saved {filename} ({len(post_res.content):,} bytes)")
                        success_count += 1
                        downloaded = True
                        break
                    else:
                        logging.warning(
                            f"No Excel file returned for {date_str} (no price announcement). (Response size: {len(post_res.content)} bytes)"
                        )
                        has_no_data = True
                        # Update tokens from response HTML in case server rotated ViewState
                        try:
                            tokens = extract_tokens(post_res.text)
                        except ValueError:
                            pass
                        break

                except requests.exceptions.RequestException as e:
                    logging.warning(f"Attempt {attempt}/{max_retries} failed for {date_str}: {e}")
                    if attempt == max_retries:
                        request_error = True
                    time.sleep(2)

            if not downloaded:
                if has_no_data:
                    no_data_count += 1
                elif request_error:
                    fail_count += 1

            if delay > 0:
                time.sleep(delay)
            current_date += timedelta(days=1)

        logging.info(
            f"Download complete: {success_count} downloaded, {skip_count} skipped, "
            f"{no_data_count} no data, {fail_count} failed."
        )


def parse_date(date_string: str) -> date:
    """Parse date string in YYYY-MM-DD or YYYYMMDD format."""
    cleaned = date_string.strip()
    if '-' in cleaned:
        return datetime.strptime(cleaned, '%Y-%m-%d').date()
    if len(cleaned) == 8 and cleaned.isdigit():
        return datetime.strptime(cleaned, '%Y%m%d').date()
    raise argparse.ArgumentTypeError(f"Invalid date format: '{date_string}'. Expected YYYY-MM-DD.")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape daily retail oil prices from Thai EPPO into Excel files."
    )
    parser.add_argument(
        '--start-date',
        type=parse_date,
        default=None,
        help="Start date (YYYY-MM-DD). If omitted, auto-detects latest date and runs incrementally.",
    )
    parser.add_argument(
        '--end-date',
        type=parse_date,
        default=None,
        help="End date (YYYY-MM-DD). Default is today.",
    )
    parser.add_argument(
        '--output-dir',
        type=Path,
        default=None,
        help="Output directory for downloaded Excel files (default: data/raw or Data).",
    )
    parser.add_argument(
        '--incremental',
        '-i',
        action='store_true',
        help="Force incremental mode starting from the latest downloaded date + 1 day.",
    )
    parser.add_argument(
        '--timeout',
        type=int,
        default=20,
        help="HTTP request timeout in seconds (default: 20).",
    )
    parser.add_argument(
        '--delay',
        type=float,
        default=1.0,
        help="Delay in seconds between requests (default: 1.0).",
    )

    args = parser.parse_args()

    output_dir = args.output_dir or get_default_output_dir()
    end_date = args.end_date or date.today()

    # Determine start date
    if args.start_date is None or args.incremental:
        latest = find_latest_downloaded_date([DEFAULT_RAW_DIR, LEGACY_DATA_DIR, output_dir])
        if latest:
            start_date = latest + timedelta(days=1)
            logging.info(f"Latest downloaded file found on {latest}. Auto-syncing from {start_date}...")
        else:
            # If no existing files found, default to beginning of 2015
            start_date = date(2015, 1, 1)
            logging.info(f"No existing files found. Defaulting start date to {start_date}...")
    else:
        start_date = args.start_date

    if start_date > end_date:
        logging.info(
            f"Already up to date. Start date ({start_date}) is after end date ({end_date}). Nothing to download."
        )
        return

    logging.info(
        f"Starting download pipeline from {start_date} to {end_date} (Target folder: {output_dir})"
    )
    download_eppo_oil_prices(
        start_date=start_date,
        end_date=end_date,
        output_dir=output_dir,
        timeout=args.timeout,
        delay=args.delay,
    )


if __name__ == '__main__':
    main()
