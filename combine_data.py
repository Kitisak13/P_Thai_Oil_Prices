import os
import glob
import pandas as pd
import re

def clean_product_name(raw_name):
    raw_name = str(raw_name).replace('\xa0', ' ').strip()
    if 'Gasohol 95-E10' in raw_name: return 'Gasohol 95-E10'
    if 'Gasohol 95-E20' in raw_name: return 'Gasohol 95-E20'
    if 'Gasohol 95-E85' in raw_name: return 'Gasohol 95-E85'
    if 'Gasohol 91-E10' in raw_name: return 'Gasohol 91-E10'
    if 'ULG 95 RON' in raw_name: return 'ULG 95 RON'
    if 'HSD B7' in raw_name: return 'HSD B7'
    if 'HSD B10' in raw_name: return 'HSD B10'
    if 'HSD B20' in raw_name: return 'HSD B20'
    if 'พรีเมียม' in raw_name or 'Premium' in raw_name:
        if 'แก๊สโซฮอล์ 95' in raw_name or 'Gasohol 95' in raw_name: return 'Gasohol 95 Premium'
        if 'เบนซิน' in raw_name or 'ULG' in raw_name: return 'ULG 95 Premium'
        if 'ดีเซล' in raw_name or 'HSD' in raw_name: return 'HSD Premium'
    match = re.search(r'\((.*?)\)', raw_name)
    if match: return match.group(1).strip()
    return raw_name

def parse_excel_file(file_path):
    filename = os.path.basename(file_path)
    # Extract date from filename: EPPO_RetailOilPrice_on_20260102.xls
    date_str = filename.split('_on_')[-1].replace('.xls', '')
    if len(date_str) == 8:
        formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    else:
        formatted_date = date_str # Fallback

    try:
        df = pd.read_excel(file_path, engine='xlrd')
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return pd.DataFrame()

    companies = ['PTT', 'BCP', 'Shell', 'Esso', 'Chevron', 'IRPC', 'PT', 'Susco', 'Pure', 'SUSCO Dealers']
    
    rows_data = []
    
    start_idx = -1
    end_idx = -1
    for idx, row in df.iterrows():
        val = str(row.iloc[0])
        if '(Gasohol' in val or 'Gasohol' in val or '(ULG' in val or 'HSD' in val or 'ดีเซล' in val or 'เบนซิน' in val or 'แก๊สโซฮอล์' in val:
            if start_idx == -1:
                start_idx = idx
        if '(Effective Date)' in val or 'มีผลตั้งแต่' in val:
            end_idx = idx
            break

    if start_idx == -1:
        start_idx = 12
    if end_idx == -1:
        end_idx = start_idx + 11

    for idx in range(start_idx, end_idx):
        if idx >= len(df):
            break
            
        row = df.iloc[idx]
        raw_product = str(row.iloc[0])
        
        if raw_product == 'nan' or raw_product.strip() == '':
            continue
            
        product_name = clean_product_name(raw_product)
        
        row_dict = {
            'Date': formatted_date,
            'Product': product_name
        }
        
        for i, comp in enumerate(companies):
            col_idx = i + 1
            if col_idx < len(row):
                val = row.iloc[col_idx]
                if pd.isna(val) or val == '' or val == ' ' or str(val).strip() == '':
                    row_dict[comp] = None
                else:
                    try:
                        row_dict[comp] = float(val)
                    except:
                        row_dict[comp] = None
            else:
                row_dict[comp] = None
                
        rows_data.append(row_dict)
        
    return pd.DataFrame(rows_data)

def main():
    data_dir = r'D:\Project\Thai-Oil-Price\Data'
    files = glob.glob(os.path.join(data_dir, '*.xls'))
    print(f"Found {len(files)} files to process.")
    
    all_dfs = []
    for i, file in enumerate(files):
        df = parse_excel_file(file)
        if not df.empty:
            all_dfs.append(df)
        if (i+1) % 100 == 0:
            print(f"Processed {i+1} files...")
            
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        # Sort by Date and Product
        final_df.sort_values(by=['Date', 'Product'], inplace=True)
        
        output_path = r'D:\Project\Thai-Oil-Price\Combined_Oil_Prices.csv'
        final_df.to_csv(output_path, index=False, encoding='utf-8')
        print(f"Successfully combined into {output_path}")
        print("Sample Data:")
        print(final_df.head(10))
    else:
        print("No data extracted.")

if __name__ == '__main__':
    main()
