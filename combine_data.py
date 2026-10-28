import argparse
import logging
from pathlib import Path
import re
from typing import Optional

import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIRS = [BASE_DIR / 'data' / 'raw', BASE_DIR / 'Data']
DEFAULT_OUTPUT_DIR = BASE_DIR / 'data' / 'processed'

COMPANY_ORDER = [
    'PTT',
    'BCP',
    'Shell',
    'Esso',
    'Chevron',
    'IRPC',
    'PT',
    'Susco',
    'Pure',
    'SUSCO Dealers',
]

# Check multi-word first when matching header text to avoid 'Susco' matching 'SUSCO Dealers'
DETECTION_COMPANIES = [
    'SUSCO Dealers',
    'PTT',
    'BCP',
    'Shell',
    'Esso',
    'Chevron',
    'IRPC',
    'PT',
    'Susco',
    'Pure',
]

PRODUCT_MAPPING = {
    'Gasohol 95-E10': 'Gasohol 95-E10',
    'Gasohol 95-E20': 'Gasohol 95-E20',
    'Gasohol 95-E85': 'Gasohol 95-E85',
    'Gasohol 91-E10': 'Gasohol 91-E10',
    'ULG 95 RON': 'ULG 95 RON',
    'HSD B7': 'HSD B7',
    'HSD B10': 'HSD B10',
    'HSD B20': 'HSD B20',
}


def clean_product_name(raw_name: str) -> str:
    """Normalize raw product name into standard English naming."""
    cleaned = str(raw_name).replace('\xa0', ' ').strip()

    # Exact product mapping
    for key, val in PRODUCT_MAPPING.items():
        if key in cleaned:
            return val

    # Premium products detection (strictly checks actual 'พรีเมียม' or 'Premium' keywords)
    if 'พรีเมียม' in cleaned or 'Premium' in cleaned:
        if any(k in cleaned for k in ('ดีเซล', 'HSD')):
            return 'HSD Premium'
        if any(k in cleaned for k in ('เบนซิน', 'ULG')):
            return 'ULG 95 Premium'
        if any(k in cleaned for k in ('แก๊สโซฮอล ออกเทน 99', 'Gasohol 99', '99')):
            return 'Gasohol 99 Premium'
        if any(k in cleaned for k in ('แก๊สโซฮอล์ 95', 'แก๊สโซฮอล ออกเทน 95', 'Gasohol 95', 'แก๊สโซฮอล์', 'แก๊สโซฮอล', '95')):
            return 'Gasohol 95 Premium'

    match = re.search(r'\((.*?)\)', cleaned)
    if match:
        return match.group(1).strip()
    return cleaned


def detect_company_columns(df: pd.DataFrame) -> dict[int, str]:
    """Dynamically detect company name for each column index by inspecting header rows 8-12."""
    col_to_company: dict[int, str] = {}
    max_row = min(12, len(df))

    for col in range(1, len(df.columns)):
        text_parts = [
            str(df.iloc[r, col])
            for r in range(min(8, max_row), max_row)
            if pd.notna(df.iloc[r, col])
        ]
        combined_text = ' '.join(' '.join(text_parts).split()).lower()

        matched_company = None
        for comp in DETECTION_COMPANIES:
            if comp == 'PT':
                if re.search(r'\bpt\b', combined_text):
                    matched_company = 'PT'
                    break
            elif comp.lower() in combined_text:
                matched_company = comp
                break

        if matched_company:
            col_to_company[col] = matched_company

    # Fallback to positional indexing if dynamic detection fails or finds too few
    if len(col_to_company) < 5:
        logging.debug("Header detection returned fewer than 5 companies; using positional fallback.")
        col_to_company = {i + 1: comp for i, comp in enumerate(COMPANY_ORDER) if (i + 1) < len(df.columns)}

    return col_to_company


def parse_excel_file(file_path: Path) -> pd.DataFrame:
    """Read and parse single EPPO daily retail oil price Excel file."""
    # Extract date from filename: EPPO_RetailOilPrice_on_YYYYMMDD.xls
    match = re.search(r'_on_(\d{4})(\d{2})(\d{2})\.xls$', file_path.name, re.IGNORECASE)
    if match:
        formatted_date = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    else:
        formatted_date = file_path.stem.split('_on_')[-1]

    # Try reading as BIFF8 .xls, fallback to HTML table if disguised
    df = None
    try:
        df = pd.read_excel(file_path, engine='xlrd')
    except Exception as e_xlrd:
        try:
            html_tables = pd.read_html(file_path)
            if html_tables:
                df = html_tables[0]
        except Exception:
            logging.warning(f"Failed to read {file_path.name} as Excel or HTML: {e_xlrd}")
            return pd.DataFrame()

    if df is None or df.empty or len(df.columns) < 2:
        return pd.DataFrame()

    col_to_company = detect_company_columns(df)

    start_idx = -1
    end_idx = -1

    for idx, row in df.iterrows():
        val = str(row.iloc[0])
        if any(
            k in val
            for k in (
                'Gasohol',
                'ULG',
                'HSD',
                'ดีเซล',
                'เบนซิน',
                'แก๊สโซฮอล์',
                '͡෹',
                'ع',
                'ู',
            )
        ):
            if start_idx == -1:
                start_idx = idx
        if '(Effective Date)' in val or 'มีผลตั้งแต่' in val or 'Effective Date' in val:
            end_idx = idx
            break

    if start_idx == -1:
        start_idx = 12
    if end_idx == -1:
        end_idx = min(start_idx + 12, len(df))

    rows_data = []
    for idx in range(start_idx, end_idx):
        if idx >= len(df):
            break

        row = df.iloc[idx]
        raw_product = str(row.iloc[0])

        if (
            pd.isna(raw_product)
            or raw_product.strip() == ''
            or raw_product == 'nan'
            or 'Effective Date' in raw_product
            or 'Unit : Baht' in raw_product
        ):
            continue

        product_name = clean_product_name(raw_product)
        if not product_name or product_name in ('Effective Date', 'Unit : Baht/Litre'):
            continue

        row_dict = {
            'Date': formatted_date,
            'Product': product_name,
        }

        # Initialize all known companies with None
        for comp in COMPANY_ORDER:
            row_dict[comp] = None

        # Fill detected companies
        for col_idx, comp in col_to_company.items():
            if col_idx < len(row):
                val = row.iloc[col_idx]
                if pd.isna(val) or str(val).strip() in ('', '-', 'nan', 'NaN'):
                    row_dict[comp] = None
                else:
                    try:
                        price = float(val)
                        row_dict[comp] = price if price > 0 else None
                    except (ValueError, TypeError):
                        row_dict[comp] = None

        rows_data.append(row_dict)

    return pd.DataFrame(rows_data)


def collect_input_files(input_dir: Optional[Path]) -> list[Path]:
    """Find input files from specified directory or default search directories."""
    if input_dir:
        files = sorted(input_dir.glob('*.xls'))
        if files:
            return files
        logging.warning(f"No .xls files found in specified directory: {input_dir}")

    # Check defaults: data/raw, then Data
    for candidate in DEFAULT_INPUT_DIRS:
        if candidate.exists():
            files = sorted(candidate.glob('*.xls'))
            if files:
                logging.info(f"Using {len(files)} files from {candidate}")
                return files

    return []


def combine_data(
    input_dir: Optional[Path] = None,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    output_format: str = 'both',
    export_parquet: bool = True,
    legacy_csv: bool = True,
) -> None:
    """Process all Excel files and export structured datasets."""
    output_dir.mkdir(parents=True, exist_ok=True)
    files = collect_input_files(input_dir)

    if not files:
        logging.error("No .xls files found to process. Please run main.py first.")
        return

    logging.info(f"Processing {len(files)} Excel files...")

    all_dfs: list[pd.DataFrame] = []
    for i, file_path in enumerate(files, start=1):
        parsed = parse_excel_file(file_path)
        if not parsed.empty:
            all_dfs.append(parsed)

        if i % 500 == 0 or i == len(files):
            logging.info(f"Progress: {i}/{len(files)} files parsed ({(i / len(files) * 100):.1f}%)")

    if not all_dfs:
        logging.error("No valid data could be extracted from files.")
        return

    logging.info("Concatenating and deduplicating data...")
    wide_df = pd.concat(all_dfs, ignore_index=True)
    wide_df.drop_duplicates(subset=['Date', 'Product'], inplace=True)
    wide_df.sort_values(by=['Date', 'Product'], inplace=True)

    # Reorder columns: Date, Product, followed by COMPANY_ORDER
    col_order = ['Date', 'Product'] + [c for c in COMPANY_ORDER if c in wide_df.columns]
    remaining_cols = [c for c in wide_df.columns if c not in col_order]
    wide_df = wide_df[col_order + remaining_cols]

    # Generate Long Format (Tidy Data)
    company_cols = [c for c in COMPANY_ORDER if c in wide_df.columns]
    long_df = wide_df.melt(
        id_vars=['Date', 'Product'],
        value_vars=company_cols,
        var_name='Company',
        value_name='Price',
    )
    # Drop rows without prices and sort
    long_df = long_df.dropna(subset=['Price'])
    long_df.sort_values(by=['Date', 'Product', 'Company'], inplace=True)

    logging.info(f"Extracted Wide Format: {wide_df.shape[0]:,} rows x {wide_df.shape[1]} columns")
    logging.info(f"Extracted Long Format: {long_df.shape[0]:,} rows x {long_df.shape[1]} columns")

    # Export Wide Format
    if output_format in ('wide', 'both'):
        wide_csv = output_dir / 'Combined_Oil_Prices_Wide.csv'
        wide_df.to_csv(wide_csv, index=False, encoding='utf-8-sig')
        logging.info(f"Saved: {wide_csv}")

        if export_parquet:
            try:
                wide_parquet = output_dir / 'Combined_Oil_Prices_Wide.parquet'
                wide_df.to_parquet(wide_parquet, index=False)
                logging.info(f"Saved: {wide_parquet}")
            except Exception as e_pq:
                logging.warning(f"Could not save Parquet (Wide): {e_pq}")

        if legacy_csv:
            legacy_path = BASE_DIR / 'Combined_Oil_Prices.csv'
            wide_df.to_csv(legacy_path, index=False, encoding='utf-8-sig')
            logging.info(f"Saved legacy copy: {legacy_path}")

    # Export Long Format
    if output_format in ('long', 'both'):
        long_csv = output_dir / 'Combined_Oil_Prices_Long.csv'
        long_df.to_csv(long_csv, index=False, encoding='utf-8-sig')
        logging.info(f"Saved: {long_csv}")

        if export_parquet:
            try:
                long_parquet = output_dir / 'Combined_Oil_Prices_Long.parquet'
                long_df.to_parquet(long_parquet, index=False)
                logging.info(f"Saved: {long_parquet}")
            except Exception as e_pq:
                logging.warning(f"Could not save Parquet (Long): {e_pq}")

    logging.info("Data combination pipeline completed successfully.")


def main():
    parser = argparse.ArgumentParser(
        description="Parse and combine daily retail oil price Excel files into Wide and Long format datasets."
    )
    parser.add_argument(
        '--input-dir',
        type=Path,
        default=None,
        help="Directory containing daily .xls files (default: data/raw or Data).",
    )
    parser.add_argument(
        '--output-dir',
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory to save combined outputs (default: data/processed).",
    )
    parser.add_argument(
        '--format',
        choices=['wide', 'long', 'both'],
        default='both',
        help="Output dataset format: 'wide', 'long', or 'both' (default: both).",
    )
    parser.add_argument(
        '--no-parquet',
        action='store_true',
        help="Disable Parquet file export.",
    )
    parser.add_argument(
        '--no-legacy-csv',
        action='store_true',
        help="Do not generate root Combined_Oil_Prices.csv.",
    )

    args = parser.parse_args()

    combine_data(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        output_format=args.format,
        export_parquet=not args.no_parquet,
        legacy_csv=not args.no_legacy_csv,
    )


if __name__ == '__main__':
    main()
