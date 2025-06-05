import sys
import os

if getattr(sys, 'frozen', False):
    sys.argv = [sys.argv[0], "run"]
    os.environ['STREAMLIT_RUNNING_VIA_PYINSTALLER'] = 'true'
    os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVE'] = 'true'
    os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

import streamlit as st
import pandas as pd
from datetime import datetime
import re
from rapidfuzz import fuzz, process
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import CellIsRule

# Diacritic normalization
def normalize_diacritics(text):
    diacritic_map = {'č':'c', 'š':'s', 'ž':'z', 'Č':'c', 'Š':'s', 'Ž':'z'}
    return ''.join(diacritic_map.get(c, c) for c in text)

# City name cleaning
def clean_city_name(city):
    city = str(city).split('-')[0].split('–')[0].strip()
    return normalize_diacritics(city).lower()

# Street address cleaning
def clean_street_name(address):
    irrelevant_words = {
        'slovenia', 'slovenija', 'slo', 'ljubljana', 'lj', 'avenija',
        'ulica', 'cesta', 'ul', 'street', 'road', 'd.o.o.', 'd.d.', 'eu', 'skl', 'vh', 'naselje', 'mesto'
    }
    address = normalize_diacritics(str(address))
    address = re.sub(r'[^\w\s]', ' ', address)
    address = re.sub(r'\s+', ' ', address).lower()
    address = re.sub(r'([a-z])(\d)', r'\1 \2', address)
    address = re.sub(r'(\d)([a-z])', r'\1 \2', address)
    parts = address.split()
    while parts and (re.match(r'^\d{4}$', parts[-1]) or parts[-1] in irrelevant_words):
        parts.pop()
    cleaned_parts = []
    for part in parts:
        if part in irrelevant_words or len(part) <= 2: continue
        cleaned_parts.append(part)
        if re.match(r'^\d+[a-z]?$', part): break
    return ' '.join(cleaned_parts).strip()

# Consignee name normalization
def normalize_consignee_name(name):
    if not isinstance(name, str): return ''
    name = name.lower().strip()
    name = re.sub(r'[.,]', '', name)
    suffixes = [
        'doo', 'd o o', 'd.o.o', 'd.o.o.', 'd d', 'd.d.', 'd.d', 'dno', 
        'd.n.o.', 'd.n.o', 'ddoo', 'ltd', 'limited', 'llc', 'gmbh', 'inc'
    ]
    for suffix in suffixes:
        name = re.sub(r'\s*' + re.escape(suffix) + r'$', '', name)
    return re.sub(r'\s+', ' ', name).strip()

# House number extraction
def extract_house_number(address):
    match = re.search(r'\b(\d+[a-z]?)\b', str(address).lower())
    return match.group(1) if match else None

# Convert house number to float
def house_number_to_float(hn):
    if not hn: return None
    hn = str(hn).lower()
    num_part = re.sub(r'[^0-9]', '', hn)
    letter_part = re.sub(r'[^a-z]', '', hn)
    if not num_part: return None
    base = float(num_part)
    if letter_part: base += (ord(letter_part) - ord('a') + 1) * 0.01
    return base

# Load street-city routes
def load_street_city_routes(path):
    try:
        df = pd.read_excel(path)
        df = df.rename(columns={df.columns[0]: 'ROUTE', df.columns[1]: 'STREET', df.columns[2]: 'CITY'})
        df['CITY_CLEAN'] = df['CITY'].apply(clean_city_name)
        df['STREET_CLEAN'] = df['STREET'].apply(clean_street_name)
        return df
    except Exception as e:
        st.error(f"Street-city routes error: {str(e)}")
        return pd.DataFrame()

# Load fallback routes
def load_fallback_routes(path):
    try:
        df = pd.read_excel(path)
        df.columns = ['ROUTE', 'ZIP']
        df['ZIP'] = df['ZIP'].astype(str).str.zfill(4)
        return df
    except Exception as e:
        st.error(f"Fallback routes error: {str(e)}")
        return pd.DataFrame()

# Load targets
def load_targets(path):
    try:
        return pd.read_excel(path, usecols="A:C", header=0, names=['ROUTE', 'Average PU stops', 'Target stops'])
    except Exception as e:
        st.warning(f"Couldn't load targets.xlsx: {str(e)}")
        return pd.DataFrame()

# Parse pieces value
def parse_pieces(value):
    try:
        value_str = str(value).replace('\\', '/')
        parts = value_str.split('/')
        numbers = re.findall(r'\d+', parts[0].strip())
        return int(numbers[0]) if numbers else 1
    except Exception:
        return 1

# Clean NaN from addresses
def clean_nan_from_address(addr):
    if isinstance(addr, str): return re.sub(r'\s*nan\s*$', '', addr).strip()
    return addr

# Process manifest file
def process_manifest(file):
    try:
        df = pd.read_excel(file) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, delimiter=',', quotechar='"', engine='python')
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
        'Cnee Str #': 'HOUSE_NUMBER',
        'PCC': 'PCC'
    }

    new_df = pd.DataFrame()
    for orig_col, new_col in column_mapping.items():
        new_df[new_col] = df[orig_col] if orig_col in df.columns else None

    if 'PIECES' in new_df.columns: new_df['PIECES'] = new_df['PIECES'].apply(parse_pieces)
    if 'PCC' in new_df.columns: new_df['PCC'] = new_df['PCC'].astype(str).str.strip().str.upper()

    street1 = new_df['CONSIGNEE_STREET1'] if 'CONSIGNEE_STREET1' in new_df.columns else pd.Series('')
    street2 = new_df['CONSIGNEE_STREET2'] if 'CONSIGNEE_STREET2' in new_df.columns else pd.Series('')
    new_df['CONSIGNEE_STREET'] = (street1.astype(str) + ' ' + street2.astype(str)).str.strip().replace('nan', '')
    
    new_df['HOUSE_NUMBER'] = new_df['CONSIGNEE_STREET'].apply(extract_house_number)
    new_df['HOUSE_NUMBER_FLOAT'] = new_df['HOUSE_NUMBER'].apply(house_number_to_float)
    new_df['STREET_NAME'] = new_df['CONSIGNEE_STREET'].apply(clean_street_name)
    
    if 'CONSIGNEE_ZIP' in new_df.columns:
        new_df['CONSIGNEE_ZIP'] = new_df['CONSIGNEE_ZIP'].astype(str).apply(lambda x: re.sub(r'\D', '', x)[:4].zfill(4))
    
    for col in ['WEIGHT', 'VOLUMETRIC_WEIGHT', 'PIECES']:
        if col in new_df.columns: new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0 if col != 'PIECES' else 1)
    
    new_df['MATCHED_ROUTE'] = None
    new_df['MATCH_SCORE'] = 0.0
    new_df['MATCH_METHOD'] = None
    new_df['CONSIGNEE_ADDRESS'] = new_df['CONSIGNEE_STREET'].fillna('').apply(clean_nan_from_address)
    new_df['CONSIGNEE_NAME_NORM'] = new_df['CONSIGNEE_NAME'].apply(normalize_consignee_name)
    return new_df

# Address matching logic
def match_address_to_route(manifest_df, street_city_routes, fallback_routes):
    manifest_df['MATCHED_ROUTE'] = None
    manifest_df['MATCH_METHOD'] = None
    manifest_df['MATCH_SCORE'] = 0.0
    special_zips = {'2000', '3000', '4000', '5000', '6000', '8000'}

    for idx, row in manifest_df.iterrows():
        zip_code = row['CONSIGNEE_ZIP']
        street_name = clean_street_name(row['CONSIGNEE_STREET'])
        city_name = clean_city_name(row['CONSIGNEE_CITY'])
        matched = False

        # Matching logic for special ZIP codes
        if zip_code in special_zips:
            if not street_city_routes.empty:
                city_matches = street_city_routes[street_city_routes['CITY_CLEAN'] == city_name]
                if not city_matches.empty:
                    matches = process.extract(street_name, city_matches['STREET_CLEAN'], scorer=fuzz.token_set_ratio, score_cutoff=70, limit=3)
                    for matched_street, score, _ in matches:
                        best_match = city_matches[city_matches['STREET_CLEAN'] == matched_street].iloc[0]
                        manifest_df.at[idx, 'MATCHED_ROUTE'] = best_match['ROUTE']
                        manifest_df.at[idx, 'MATCH_METHOD'] = 'Street-City'
                        manifest_df.at[idx, 'MATCH_SCORE'] = float(score)
                        matched = True
                        break
                    if matched: continue

            if zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[fallback_routes['ZIP'] == zip_code, 'ROUTE'].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
                continue

        else:
            if zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[fallback_routes['ZIP'] == zip_code, 'ROUTE'].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
                continue

            if not street_city_routes.empty:
                city_matches = street_city_routes[street_city_routes['CITY_CLEAN'] == city_name]
                if not city_matches.empty:
                    matches = process.extract(street_name, city_matches['STREET_CLEAN'], scorer=fuzz.token_set_ratio, score_cutoff=70, limit=3)
                    for matched_street, score, _ in matches:
                        best_match = city_matches[city_matches['STREET_CLEAN'] == matched_street].iloc[0]
                        manifest_df.at[idx, 'MATCHED_ROUTE'] = best_match['ROUTE']
                        manifest_df.at[idx, 'MATCH_METHOD'] = 'Street-City'
                        manifest_df.at[idx, 'MATCH_SCORE'] = float(score)
                        matched = True
                        break
                    if matched: continue

    return manifest_df

# Fuzzy grouping of names
def fuzzy_group_names(df, group_col='CONSIGNEE_NAME_NORM', threshold=90):
    grouped_data = []
    processed = set()
    for name in df[group_col].unique():
        if name in processed: continue
        matches = process.extract(name, df[group_col], scorer=fuzz.ratio, score_cutoff=threshold)
        matched_names = {match[0] for match in matches}
        group_rows = df[df[group_col].isin(matched_names)]
        non_null_names = group_rows['CONSIGNEE_NAME'].dropna()
        canonical_name = non_null_names.mode().iloc[0] if not non_null_names.empty else "Unknown Consignee"
        grouped_data.append({
            'MATCHED_ROUTE': group_rows['MATCHED_ROUTE'].iloc[0] if not group_rows.empty else '',
            'CANONICAL_NAME': canonical_name,
            'shipment_count': group_rows['shipment_count'].sum(),
            'total_pieces': group_rows['total_pieces'].sum(),
            'total_weight': group_rows['total_weight'].sum()
        })
        processed.update(matched_names)
    return pd.DataFrame(grouped_data)

# Conditional formatting for targets
def add_target_conditional_formatting(sheet, col_letter, start_row, end_row):
    red_font = Font(color='FF0000')
    yellow_font = Font(color='FFA500')
    green_font = Font(color='008000')
    sheet.conditional_formatting.add(f'{col_letter}{start_row}:{col_letter}{end_row}',
        CellIsRule(operator='greaterThan', formula=['5'], font=red_font))
    sheet.conditional_formatting.add(f'{col_letter}{start_row}:{col_letter}{end_row}',
        CellIsRule(operator='lessThan', formula=['-5'], font=red_font))
    sheet.conditional_formatting.add(f'{col_letter}{start_row}:{col_letter}{end_row}',
        CellIsRule(operator='between', formula=['-5', '0'], font=yellow_font))
    sheet.conditional_formatting.add(f'{col_letter}{start_row}:{col_letter}{end_row}',
        CellIsRule(operator='between', formula=['0', '5'], font=green_font))

# Generate all reports
def generate_reports(manifest_df, output_path, weight_thr=70, vol_weight_thr=150, pieces_thr=6,
                    vehicle_weight_thr=70, vehicle_vol_thr=150, vehicle_pieces_thr=12,
                    vehicle_kg_per_piece_thr=10, vehicle_van_max_pieces=20):
    # Aggregated data processing
    hwb_aggregated = manifest_df.groupby(['HWB', 'MATCHED_ROUTE']).agg({
        'CONSIGNEE_NAME_NORM': 'first',
        'CONSIGNEE_NAME': 'first',
        'CONSIGNEE_ZIP': 'first',
        'CONSIGNEE_ADDRESS': 'first',
        'WEIGHT': 'sum',
        'VOLUMETRIC_WEIGHT': 'sum',
        'PIECES': 'first'
    }).reset_index()

    # Route summary calculations
    route_summary = hwb_aggregated.groupby('MATCHED_ROUTE').agg(
        total_shipments=('HWB', 'count'),
        unique_consignees=('CONSIGNEE_NAME_NORM', 'nunique'),
        total_weight=('WEIGHT', 'sum'),
        total_pieces=('PIECES', 'sum')
    ).reset_index().rename(columns={'MATCHED_ROUTE': 'ROUTE'})

    # Merge with targets
    targets_df = load_targets('input/targets.xlsx')
    if not targets_df.empty:
        route_summary = pd.merge(route_summary, targets_df, on='ROUTE', how='left')
    else:
        route_summary['Average PU stops'] = 0
        route_summary['Target stops'] = 0

    # Formatting and calculations
    route_summary['total_weight'] = route_summary['total_weight'].round(1)
    route_summary['Predicted Stops'] = route_summary['unique_consignees'] + route_summary['Average PU stops']
    route_summary['Predicted - Target'] = route_summary['Predicted Stops'] - route_summary['Target stops']
    for col in ['Average PU stops', 'Predicted Stops', 'Predicted - Target']:
        route_summary[col] = route_summary[col].round(0).astype('Int64')
    
    # Report columns setup
    route_summary.insert(5, '', '')
    route_summary = route_summary.reindex(columns=[
        'ROUTE', 'total_shipments', 'unique_consignees', 'total_weight', 'total_pieces', '',
        'Average PU stops', 'Predicted Stops', 'Target stops', 'Predicted - Target'
    ]).rename(columns={'ROUTE': 'MATCHED_ROUTE'})

    # Special cases processing
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
        'PIECES': 'first'
    }).reset_index()
    special_cases['TRIGGER_REASON'] = special_cases.apply(lambda row: ', '.join([
        f'Weight >{weight_thr}kg' if row['WEIGHT'] > weight_thr else '',
        f'Volumetric >{vol_weight_thr}kg' if row['VOLUMETRIC_WEIGHT'] > vol_weight_thr else '',
        f'Pieces >{pieces_thr}' if row['PIECES'] > pieces_thr else ''
    ]).strip(', '), axis=1)
    special_cases['WEIGHT_PER_PIECE'] = special_cases.apply(
        lambda x: round(x['WEIGHT']/x['PIECES'], 2) if x['PIECES'] > pieces_thr else None, axis=1)

    # Vehicle suggestions
    special_cases[''] = ''
    def get_vehicle_suggestion(row):
        conditions = [
            row['WEIGHT'] > vehicle_weight_thr,
            row['VOLUMETRIC_WEIGHT'] > vehicle_vol_thr,
            row['PIECES'] > vehicle_pieces_thr,
            (row['WEIGHT_PER_PIECE'] > vehicle_kg_per_piece_thr) and 
            (row['PIECES'] > vehicle_van_max_pieces)
        ]
        return "Truck" if any(conditions) else "Van"
    special_cases['Capacity Suggestion'] = special_cases.apply(get_vehicle_suggestion, axis=1)
    special_cases = special_cases.sort_values(by="CONSIGNEE_ZIP", ascending=True)

    # Matching details
    matching_details = manifest_df[[
        'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ZIP', 'CONSIGNEE_ADDRESS',
        'MATCHED_ROUTE', 'MATCH_METHOD'
    ]].copy()
    matching_details['MATCHED_ROUTE'] = matching_details['MATCHED_ROUTE'].fillna('UNMATCHED')
    matching_details['MATCH_METHOD'] = matching_details['MATCH_METHOD'].fillna('UNMATCHED')
    matching_details.rename(columns={'MATCH_METHOD': 'MATCHING_METHOD'}, inplace=True)

    # File paths
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    os.makedirs(output_path, exist_ok=True)
    summary_path = f"{output_path}/route_summary_{timestamp}.xlsx"
    special_cases_path = f"{output_path}/special_cases_{timestamp}.xlsx"
    matching_details_path = f"{output_path}/matching_details_{timestamp}.xlsx"
    wth_mpcs_path = f"{output_path}/WTH_MPCS_Report_{timestamp}.xlsx"

    # Save route summary with formatting
    with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
        route_summary.to_excel(writer, sheet_name='Summary', index=False)
        workbook = writer.book
        sheet = workbook['Summary']
        sheet.append([])
        sheet.append(["PCC Statistics:"])
        sheet.append(["Product", "Shipments", "Pieces", "Pieces/Shipment"])
        pcc_categories = [
            ('WPX', 'WPX'), ('TDY', 'TDY'), ('ESI', 'ESI'),
            ('ECX', 'ECX'), ('ESU', 'ESU'), ('ALL', 'All volume')
        ]
        for pcc_code, label in pcc_categories:
            if 'PCC' in manifest_df.columns:
                filtered_df = manifest_df[manifest_df['PCC'] == pcc_code] if pcc_code != 'ALL' else manifest_df[manifest_df['PCC'].notna()]
                shipments = filtered_df['HWB'].nunique()
                pieces = filtered_df['PIECES'].sum()
                ratio = round(pieces / shipments, 2) if shipments > 0 else 0.0
            else:
                shipments, pieces, ratio = 0, 0, 0.0
            sheet.append([label, shipments, pieces, ratio])
        for row in sheet.iter_rows(min_row=2, max_row=len(route_summary)+1, min_col=4, max_col=4):
            for cell in row: cell.number_format = '0.0'
        add_target_conditional_formatting(sheet, 'J', 2, len(route_summary)+1)

    # Save special cases with formatting
    with pd.ExcelWriter(special_cases_path, engine='openpyxl') as writer:
        special_cases.to_excel(writer, index=False)
        workbook = writer.book
        sheet = workbook.active
        truck_font = Font(color='FF0000', bold=True)
        van_font = Font(color='000000', bold=True)
        suggestion_col = special_cases.columns.get_loc('Capacity Suggestion') + 1
        for row_idx in range(2, len(special_cases)+2):
            cell = sheet.cell(row=row_idx, column=suggestion_col)
            cell.font = truck_font if cell.value == "Truck" else van_font

    # Save other reports
    matching_details.to_excel(matching_details_path, index=False)
    special_cases.sort_values('PIECES', ascending=False).to_excel(wth_mpcs_path, index=False)

    # MBX Report
    mbx_routes = {'MB1A', 'MB1B', 'MB1C', 'MB1X', 'MB2A', 'MB2B', 'MB2C', 'MB2X'}
    mbx_report = manifest_df[manifest_df['MATCHED_ROUTE'].isin(mbx_routes)][[
        'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ADDRESS', 'CONSIGNEE_ZIP', 'MATCHED_ROUTE'
    ]].sort_values(by=['MATCHED_ROUTE', 'CONSIGNEE_ZIP'], ascending=[True, True])
    mbx_path = None
    if not mbx_report.empty:
        mbx_path = f"{output_path}/MBX_details_{timestamp}.xlsx"
        mbx_report.to_excel(mbx_path, index=False)

    # Priority shipments report
    pcc_col = next((col for col in manifest_df.columns if col.strip().upper() == 'PCC'), None)
    if pcc_col:
        manifest_df[pcc_col] = manifest_df[pcc_col].astype(str).str.strip().str.upper()
        priority_codes = ['CMX', 'WMX', 'TDT', 'TDY']
        priority_pccs = manifest_df[manifest_df[pcc_col].isin(priority_codes)]
        if not priority_pccs.empty:
            wb = Workbook()
            ws = wb.active
            ws.title = "Priority Shipments"
            cols = ['MATCHED_ROUTE', 'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ZIP', pcc_col, 'WEIGHT', 'VOLUMETRIC_WEIGHT', 'PIECES']
            header_font = Font(bold=True)
            ws['A1'] = "CMX/WMX Priority Shipments"
            ws['A1'].font = header_font
            for col_idx, col in enumerate(cols, 1):
                ws.cell(row=3, column=col_idx, value=col).font = header_font
            for row_idx, row in enumerate(dataframe_to_rows(
                priority_pccs[priority_pccs[pcc_col].isin(['CMX', 'WMX'])].sort_values(
                    by=['CONSIGNEE_ZIP', 'MATCHED_ROUTE']), index=False, header=False), 4):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            last_row = ws.max_row + 3
            ws.cell(row=last_row, column=1, value="TDT/TDY Priority Shipments").font = header_font
            for col_idx, col in enumerate(cols, 1):
                ws.cell(row=last_row + 2, column=col_idx, value=col).font = header_font
            for row_idx, row in enumerate(dataframe_to_rows(
                priority_pccs[priority_pccs[pcc_col].isin(['TDT', 'TDY'])].sort_values(
                    by=['CONSIGNEE_ZIP', 'MATCHED_ROUTE']), index=False, header=False), last_row + 3):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            wb.save(f"{output_path}/Priority_Shipments_{timestamp}.xlsx")

    return timestamp, route_summary, mbx_path

# Streamlit UI
def main():
    st.title("🚚 Delivery Route Analyzer")
    st.sidebar.header("Settings")
    weight_thr = st.sidebar.number_input("Weight Threshold (kg)", value=70)
    vol_weight_thr = st.sidebar.number_input("Volumetric Weight Threshold (kg)", value=150)
    pieces_thr = st.sidebar.number_input("Pieces Threshold", value=6)
    st.sidebar.subheader("Vehicle Suggestions")
    vehicle_weight_thr = st.sidebar.number_input("Truck weight threshold (kg)", value=70)
    vehicle_vol_thr = st.sidebar.number_input("Truck volumetric threshold (kg)", value=150)
    vehicle_pieces_thr = st.sidebar.number_input("Truck pieces threshold", value=12)
    vehicle_kg_per_piece_thr = st.sidebar.number_input("Max kg/piece for Van", value=10)
    vehicle_van_max_pieces = st.sidebar.number_input("Max pieces for Van", value=20)
    
    uploaded_file = st.file_uploader("Upload Manifest File", type=["xlsx", "xls", "csv"])
    if uploaded_file:
        st.info("Processing manifest...")
        street_city_routes = load_street_city_routes('input/route_street_city.xlsx')
        fallback_routes = load_fallback_routes('input/routes_database.xlsx')
        manifest = process_manifest(uploaded_file)
        matched_manifest = match_address_to_route(manifest, street_city_routes, fallback_routes)
        output_path = "output"
        os.makedirs(output_path, exist_ok=True)
        timestamp, route_summary, mbx_path = generate_reports(
            matched_manifest, output_path,
            weight_thr, vol_weight_thr, pieces_thr,
            vehicle_weight_thr, vehicle_vol_thr,
            vehicle_pieces_thr, vehicle_kg_per_piece_thr, vehicle_van_max_pieces
        )

        try:
            if not route_summary.empty:
                predicted_spr = route_summary['Predicted Stops'].mean()
                st.metric("Predicted SPR (Average Predicted Stops)", f"{predicted_spr:.1f}")
            else:
                st.warning("No routes matched - cannot calculate SPR")
        except Exception as e:
            st.error(f"SPR calculation error: {str(e)}")
        
        st.success("Processing complete! 🎉")
        
        # Reports download section
        st.subheader("Standard Reports")
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

        st.subheader("Additional Reports")
        col4, col5 = st.columns(2)
        with col4:
            with open(f"{output_path}/WTH_MPCS_Report_{timestamp}.xlsx", "rb") as f:
                st.download_button("WTH MPCS Report", f, f"WTH_MPCS_Report_{timestamp}.xlsx")
        with col5:
            try:
                with open(f"{output_path}/Priority_Shipments_{timestamp}.xlsx", "rb") as f:
                    st.download_button("Priority Shipments", f, f"Priority_Shipments_{timestamp}.xlsx")
            except FileNotFoundError:
                st.write("Priority report unavailable (missing PCC data)")

        st.subheader("Specialized Reports")
        col6, col7 = st.columns(2)
        with col7:
            if mbx_path and os.path.exists(mbx_path):
                with open(mbx_path, "rb") as f:
                    st.download_button("MBX Details", f, f"MBX_details_{timestamp}.xlsx",
                                      help="Shipments for MB1/MB2 routes sorted by route and ZIP")
            else:
                st.write("No MBX shipments found")

        st.subheader("Preview of Processed Data")
        st.dataframe(matched_manifest.head(10))

if __name__ == "__main__":
    main()
