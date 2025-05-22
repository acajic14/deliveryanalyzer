import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
from rapidfuzz import fuzz, process

# ==============================
# Helper Functions (Keep These)
# ==============================

def normalize_diacritics(text):
    diacritic_map = {'č':'c', 'š':'s', 'ž':'z', 'Č':'c', 'Š':'s', 'Ž':'z'}
    return ''.join(diacritic_map.get(c, c) for c in text)

def clean_city_name(city):
    city = str(city).split('-')[0].split('–')[0].strip()
    return normalize_diacritics(city).lower()

def clean_street_name(address):
    irrelevant_words = {
        'slovenia', 'slovenija', 'slo', 'ljubljana', 'lj', 'avenija',
        'ulica', 'cesta', 'ul', 'street', 'road', 'd.o.o.', 'd.d.', 'eu', 'skl', 'vh', 'naselje', 'mesto'
    }
    address = normalize_diacritics(str(address))
    address = re.sub(r'[^a-zA-Z0-9 ]', ' ', address)
    address = re.sub(r'\s+', ' ', address).lower()
    address = re.sub(r'([a-z])(\d)', r'\1 \2', address)
    address = re.sub(r'(\d)([a-z])', r'\1 \2', address)
    parts = address.split()
    while parts and (
        re.match(r'^\d{4}$', parts[-1]) or
        parts[-1] in irrelevant_words
    ):
        parts.pop()
    cleaned_parts = []
    for part in parts:
        if part in irrelevant_words or len(part) <= 2:
            continue
        cleaned_parts.append(part)
        if re.match(r'^\d+[a-z]?$', part):
            break
    return ' '.join(cleaned_parts).strip()

def normalize_consignee_name(name):
    if not isinstance(name, str):
        return ''
    name = name.lower().strip()
    name = re.sub(r'[.,]', '', name)
    suffixes = [
        'doo', 'd o o', 'd.o.o', 'd.o.o.', 'd d', 'd.d.', 'd.d', 'dno', 'd.n.o.', 'd.n.o', 'ddoo', 'ltd', 'limited', 'llc', 'gmbh', 'inc'
    ]
    for suffix in suffixes:
        pattern = r'\s*' + re.escape(suffix) + r'$'
        name = re.sub(pattern, '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def extract_house_number(address):
    match = re.search(r'\b(\d+[a-z]?)\b', str(address).lower())
    return match.group(1) if match else None

def house_number_to_float(hn):
    if not hn:
        return None
    hn = str(hn).lower()
    num_part = re.sub(r'[^0-9]', '', hn)
    letter_part = re.sub(r'[^a-z]', '', hn)
    if not num_part:
        return None
    base = float(num_part)
    if letter_part:
        base += (ord(letter_part) - ord('a') + 1) * 0.01
    return base

def load_street_city_routes(path):
    try:
        df = pd.read_excel(path)
        df = df.rename(columns={
            df.columns[0]: 'ROUTE',
            df.columns[1]: 'STREET',
            df.columns[2]: 'CITY'
        })
        df['CITY_CLEAN'] = df['CITY'].apply(clean_city_name)
        df['STREET_CLEAN'] = df['STREET'].apply(clean_street_name)
        return df
    except Exception as e:
        st.error(f"Street-city routes file error: {str(e)}")
        return pd.DataFrame()

def load_zip_routes_xls(path):
    try:
        xls = pd.ExcelFile(path)
        zip_sheets = {}
        for sheet in xls.sheet_names:
            if not sheet.isdigit() or len(sheet) != 4:
                continue
            df = pd.read_excel(xls, sheet_name=sheet, header=1)
            df = df.rename(columns={
                df.columns[0]: 'STREET',
                df.columns[1]: 'LOW',
                df.columns[2]: 'HIGH',
                df.columns[4]: 'ROUTE'
            })
            df['STREET_CLEAN'] = df['STREET'].apply(clean_street_name)
            df['LOW_FLOAT'] = df['LOW'].apply(lambda x: house_number_to_float(str(x)) if pd.notna(x) else None)
            df['HIGH_FLOAT'] = df['HIGH'].apply(lambda x: house_number_to_float(str(x)) if pd.notna(x) else None)
            zip_sheets[sheet] = df
        return zip_sheets
    except Exception as e:
        st.error(f"ZIP routes file error: {str(e)}")
        return {}

def load_fallback_routes(path):
    try:
        df = pd.read_excel(path)
        df.columns = ['ROUTE', 'ZIP']
        df['ZIP'] = df['ZIP'].astype(str).str.zfill(4)
        return df
    except Exception as e:
        st.error(f"Fallback routes error: {str(e)}")
        return pd.DataFrame()

def parse_pieces(value):
    try:
        value_str = str(value).replace('\\', '/')
        parts = value_str.split('/')
        first_part = parts[0].strip()
        numbers = re.findall(r'\d+', first_part)
        return int(numbers[0]) if numbers else 1
    except Exception:
        return 1

def clean_nan_from_address(addr):
    if isinstance(addr, str):
        return re.sub(r'\s*nan\s*$', '', addr).strip()
    return addr

def process_manifest(file):
    try:
        if file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
        elif file.name.endswith('.csv'):
            df = pd.read_csv(file, delimiter=',', quotechar='"', on_bad_lines='warn', dtype=str, engine='python')
        else:
            st.error("Unsupported file format")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return pd.DataFrame()

    column_mapping = {
        'HWB': 'HWB',
        '# Pcs\\Tot Pcs': 'PIECES',
        'Wt': 'WEIGHT',
        'Vol Wt': 'VOLUMETRIC_WEIGHT',
        'Cnee Nm': 'CONSIGNEE_NAME',
        'Cnee Addr Ln 1': 'CONSIGNEE_STREET1',
        'Cnee Addr Ln 2': 'CONSIGNEE_STREET2',
        'Cnee Zip': 'CONSIGNEE_ZIP',
        'Cnee City': 'CONSIGNEE_CITY',
        'Cnee Str #': 'HOUSE_NUMBER'
    }

    new_df = pd.DataFrame()
    for orig_col, new_col in column_mapping.items():
        if orig_col in df.columns:
            new_df[new_col] = df[orig_col]
        else:
            if new_col in ['HWB', 'CONSIGNEE_ZIP', 'WEIGHT', 'PIECES']:
                new_df[new_col] = None

    street1 = new_df['CONSIGNEE_STREET1'] if 'CONSIGNEE_STREET1' in new_df.columns else pd.Series('', index=new_df.index)
    street2 = new_df['CONSIGNEE_STREET2'] if 'CONSIGNEE_STREET2' in new_df.columns else pd.Series('', index=new_df.index)
    new_df['CONSIGNEE_STREET'] = street1.astype(str).fillna('') + ' ' + street2.astype(str).fillna('')
    new_df['CONSIGNEE_STREET'] = new_df['CONSIGNEE_STREET'].str.strip().replace('nan', '')

    new_df['HOUSE_NUMBER'] = new_df['CONSIGNEE_STREET'].apply(extract_house_number)
    new_df['HOUSE_NUMBER_FLOAT'] = new_df['HOUSE_NUMBER'].apply(house_number_to_float)
    new_df['STREET_NAME'] = new_df['CONSIGNEE_STREET'].apply(clean_street_name)

    if 'PIECES' in new_df.columns:
        new_df['PIECES'] = new_df['PIECES'].apply(parse_pieces)

    if 'CONSIGNEE_ZIP' in new_df.columns:
        new_df['CONSIGNEE_ZIP'] = (
            new_df['CONSIGNEE_ZIP']
            .astype(str)
            .apply(lambda x: ''.join(re.findall(r'\d+', x))[:4].zfill(4))
        )

    numeric_cols = ['WEIGHT', 'VOLUMETRIC_WEIGHT', 'PIECES']
    for col in numeric_cols:
        if col in new_df.columns:
            new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0 if col != 'PIECES' else 1)

    new_df['MATCHED_ROUTE'] = None
    new_df['MATCH_SCORE'] = 0.0
    new_df['MATCH_METHOD'] = None
    new_df['CONSIGNEE_ADDRESS'] = new_df['CONSIGNEE_STREET'].fillna('').apply(clean_nan_from_address)
    new_df['CONSIGNEE_NAME_NORM'] = new_df['CONSIGNEE_NAME'].apply(normalize_consignee_name)
    return new_df

def match_address_to_route(manifest_df, street_city_routes, zip_sheets, fallback_routes):
    manifest_df['MATCHED_ROUTE'] = None
    manifest_df['MATCH_METHOD'] = None
    manifest_df['MATCH_SCORE'] = 0.0

    for idx, row in manifest_df.iterrows():
        zip_code = row['CONSIGNEE_ZIP']
        street_name = clean_street_name(row['CONSIGNEE_STREET'])
        city_name = clean_city_name(row['CONSIGNEE_CITY'])
        house_num = row['HOUSE_NUMBER_FLOAT']
        matched = False

        # 1. Street-City matching
        if not street_city_routes.empty:
            city_matches = street_city_routes[street_city_routes['CITY_CLEAN'] == city_name]
            if not city_matches.empty:
                matches = process.extract(
                    street_name,
                    city_matches['STREET_CLEAN'],
                    scorer=fuzz.token_set_ratio,
                    score_cutoff=70,
                    limit=3
                )
                for matched_street, score, _ in matches:
                    best_match = city_matches[city_matches['STREET_CLEAN'] == matched_street].iloc[0]
                    manifest_df.at[idx, 'MATCHED_ROUTE'] = best_match['ROUTE']
                    manifest_df.at[idx, 'MATCH_METHOD'] = 'Street-City'
                    manifest_df.at[idx, 'MATCH_SCORE'] = float(score)
                    matched = True
                    break

        # 2. ZIP-specific sheets
        if not matched and zip_code in zip_sheets:
            sheet = zip_sheets[zip_code]
            matches = process.extract(street_name, sheet['STREET_CLEAN'], scorer=fuzz.token_set_ratio, score_cutoff=65, limit=5)
            for matched_street, score, _ in matches:
                street_rows = sheet[sheet['STREET_CLEAN'] == matched_street]
                for _, r in street_rows.iterrows():
                    if (pd.isna(r['LOW_FLOAT']) or pd.isna(r['HIGH_FLOAT'])) or \
                       (r['LOW_FLOAT'] <= house_num <= r['HIGH_FLOAT']):
                        manifest_df.at[idx, 'MATCHED_ROUTE'] = r['ROUTE']
                        manifest_df.at[idx, 'MATCH_METHOD'] = 'Street+House'
                        manifest_df.at[idx, 'MATCH_SCORE'] = float(score)
                        matched = True
                        break
                if matched:
                    break

        # 3. ZIP fallback
        if not matched and zip_code in fallback_routes['ZIP'].values:
            manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[fallback_routes['ZIP'] == zip_code, 'ROUTE'].values[0]
            manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
            manifest_df.at[idx, 'MATCH_SCORE'] = 100.0

    return manifest_df

def fuzzy_group_names(df, group_col='CONSIGNEE_NAME_NORM', threshold=90):
    grouped_data = []
    processed = set()
    for name in df[group_col].unique():
        if name in processed:
            continue
        matches = process.extract(name, df[group_col], scorer=fuzz.ratio, score_cutoff=threshold)
        matched_names = {match[0] for match in matches}
        group_rows = df[df[group_col].isin(matched_names)]
        non_null_names = group_rows['CONSIGNEE_NAME'].dropna()
        if non_null_names.empty:
            canonical_name = "Unknown Consignee"
        else:
            mode_result = non_null_names.mode()
            canonical_name = mode_result.iloc[0] if not mode_result.empty else non_null_names.iloc[0]
        grouped_data.append({
            'MATCHED_ROUTE': group_rows['MATCHED_ROUTE'].iloc[0] if not group_rows.empty else '',
            'CANONICAL_NAME': canonical_name,
            'shipment_count': group_rows['shipment_count'].sum(),
            'total_pieces': group_rows['total_pieces'].sum(),
            'total_weight': group_rows['total_weight'].sum()
        })
        processed.update(matched_names)
    return pd.DataFrame(grouped_data)

def generate_reports(manifest_df, output_path, weight_thr=70, vol_weight_thr=150, pieces_thr=6):
    hwb_aggregated = manifest_df.groupby(['HWB', 'MATCHED_ROUTE']).agg({
        'CONSIGNEE_NAME_NORM': 'first',
        'CONSIGNEE_NAME': 'first',
        'CONSIGNEE_ZIP': 'first',
        'CONSIGNEE_ADDRESS': 'first',
        'WEIGHT': 'sum',
        'VOLUMETRIC_WEIGHT': 'sum',
        'PIECES': 'sum'
    }).reset_index()

    route_summary = hwb_aggregated.groupby('MATCHED_ROUTE').agg(
        total_shipments=('HWB', 'count'),
        unique_consignees=('CONSIGNEE_NAME_NORM', 'nunique'),
        total_weight=('WEIGHT', 'sum'),
        total_pieces=('PIECES', 'sum')
    ).reset_index()
    route_summary['total_weight'] = route_summary['total_weight'].round().astype(int)
    route_summary['total_pieces'] = route_summary['total_pieces'].astype(int)

    consignee_summary = hwb_aggregated.groupby(
        ['MATCHED_ROUTE', 'CONSIGNEE_NAME_NORM']
    ).agg(
        consignee_display=('CONSIGNEE_NAME', 'first'),
        shipment_count=('HWB', 'count'),
        total_pieces=('PIECES', 'sum'),
        total_weight=('WEIGHT', 'sum')
    ).reset_index().rename(columns={'consignee_display': 'CONSIGNEE_NAME'})

    consignee_summary_fuzzy = consignee_summary.groupby('MATCHED_ROUTE', group_keys=False).apply(
        lambda x: fuzzy_group_names(x, group_col='CONSIGNEE_NAME_NORM', threshold=90)
    ).reset_index(drop=True)

    route_details = hwb_aggregated[[
        'MATCHED_ROUTE', 'HWB', 'CONSIGNEE_NAME',
        'CONSIGNEE_ZIP', 'CONSIGNEE_ADDRESS',
        'WEIGHT', 'PIECES'
    ]].sort_values(
        by=['MATCHED_ROUTE', 'CONSIGNEE_ZIP', 'CONSIGNEE_NAME'],
        ascending=[True, True, True]
    )

    def get_trigger_reason(row):
        reasons = []
        if row['WEIGHT'] > weight_thr: reasons.append(f'Weight >{weight_thr}kg')
        if row['VOLUMETRIC_WEIGHT'] > vol_weight_thr: reasons.append(f'Volumetric >{vol_weight_thr}kg')
        if row['PIECES'] > pieces_thr: reasons.append(f'Pieces >{pieces_thr}')
        return ', '.join(reasons) if reasons else None

    special_cases = manifest_df[
        (manifest_df['WEIGHT'] > weight_thr) |
        (manifest_df['VOLUMETRIC_WEIGHT'] > vol_weight_thr) |
        (manifest_df['PIECES'] > pieces_thr)
    ].copy()
    special_cases = special_cases.groupby('HWB').agg({
        'CONSIGNEE_NAME': 'first',
        'CONSIGNEE_ZIP': 'first',
        'MATCHED_ROUTE': 'first',
        'WEIGHT': 'max',
        'VOLUMETRIC_WEIGHT': 'max',
        'PIECES': 'sum',
    }).reset_index()
    special_cases['TRIGGER_REASON'] = special_cases.apply(get_trigger_reason, axis=1)
    special_cases['WEIGHT_PER_PIECE'] = special_cases.apply(
        lambda x: round(x['WEIGHT']/x['PIECES'], 2) if x['PIECES'] > pieces_thr else None,
        axis=1
    )
    special_cases = special_cases[[
        'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ZIP', 'MATCHED_ROUTE',
        'WEIGHT', 'VOLUMETRIC_WEIGHT', 'PIECES', 'TRIGGER_REASON', 'WEIGHT_PER_PIECE'
    ]].sort_values(
        by=['CONSIGNEE_ZIP', 'MATCHED_ROUTE', 'CONSIGNEE_NAME'],
        ascending=[True, True, True]
    )

    matching_details = manifest_df[[
        'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ZIP', 'CONSIGNEE_ADDRESS',
        'MATCHED_ROUTE', 'MATCH_METHOD'
    ]].copy()
    matching_details['MATCHED_ROUTE'] = matching_details['MATCHED_ROUTE'].fillna('UNMATCHED')
    matching_details['MATCH_METHOD'] = matching_details['MATCH_METHOD'].fillna('UNMATCHED')
    matching_details.rename(columns={
        'MATCH_METHOD': 'MATCHING_METHOD'
    }, inplace=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    with pd.ExcelWriter(f"{output_path}/route_summary_{timestamp}.xlsx") as writer:
        route_summary.to_excel(writer, sheet_name='Summary', index=False)
        consignee_summary_fuzzy.rename(columns={'CANONICAL_NAME': 'CONSIGNEE_NAME'}, inplace=True)
        consignee_summary_fuzzy.to_excel(writer, sheet_name='Consignee_Summary', index=False)
        route_details.to_excel(writer, sheet_name='Route_Details', index=False)

    special_cases.to_excel(f"{output_path}/special_cases_{timestamp}.xlsx", index=False)
    matching_details.to_excel(f"{output_path}/matching_details_{timestamp}.xlsx", index=False)
    return timestamp

# ==============================
# Streamlit UI
# ==============================

def main():
    st.title("🚚 Delivery Route Analyzer")
    st.sidebar.header("Settings")
    
    # Threshold controls
    weight_thr = st.sidebar.number_input("Weight Threshold (kg)", value=70)
    vol_weight_thr = st.sidebar.number_input("Volumetric Weight Threshold (kg)", value=150)
    pieces_thr = st.sidebar.number_input("Pieces Threshold", value=6)
    
    # File upload
    uploaded_file = st.file_uploader("Upload Manifest File", type=["xlsx", "xls", "csv"])
    
    if uploaded_file:
        st.info("Processing manifest...")
        
        # Load route databases
        street_city_routes = load_street_city_routes('input/route_street_city.xlsx')
        zip_routes = load_zip_routes_xls('input/zip_routes.xls')
        fallback_routes = load_fallback_routes('input/routes_database.xlsx')
        
        # Process data
        manifest = process_manifest(uploaded_file)
        matched_manifest = match_address_to_route(manifest, street_city_routes, zip_routes, fallback_routes)
        
        # Generate reports
        output_path = "output"
        os.makedirs(output_path, exist_ok=True)
        timestamp = generate_reports(matched_manifest, output_path, weight_thr, vol_weight_thr, pieces_thr)
        
        # Show success
        st.success("Processing complete! 🎉")
        
        # Download buttons
        st.subheader("Download Reports")
        col1, col2, col3 = st.columns(3)
        with col1:
            with open(f"{output_path}/route_summary_{timestamp}.xlsx", "rb") as f:
                st.download_button("Route Summary", f, f"route_summary_{timestamp}.xlsx")
        with col2:
            with open(f"{output_path}/special_cases_{timestamp}.xlsx", "rb") as f:
                st.download_button("Special Cases", f, f"special_cases_{timestamp}.xlsx")
        with col3:
            with open(f"{output_path}/matching_details_{timestamp}.xlsx", "rb") as f:
                st.download_button("Matching Details", f, f"matching_details_{timestamp}.xlsx")
        
        # Preview
        st.subheader("Preview of Processed Data")
        st.dataframe(matched_manifest.head(10))

if __name__ == "__main__":
    main()
