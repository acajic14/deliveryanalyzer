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
    address = re.sub(r'[^\w\s]', ' ', address)
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
        'doo', 'd o o', 'd.o.o', 'd.o.o.', 'd d', 'd.d.', 'd.d', 'dno', 'd.n.o.', 'd.n.o', 'ddoo', 
        'ltd', 'limited', 'llc', 'gmbh', 'inc'
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
        st.error(f"Street-city routes error: {str(e)}")
        return pd.DataFrame()

def load_fallback_routes(path):
    try:
        df = pd.read_excel(path)
        df.columns = ['ROUTE', 'ZIP']
        df['ZIP'] = df['ZIP'].astype(str).str.zfill(4)
        return df
    except Exception as e:
        st.error(f"Fallback routes error: {str(e)}")
        return pd.DataFrame()

def load_targets(path):
    try:
        df = pd.read_excel(path, usecols="A:C", header=0, names=['ROUTE', 'Average PU stops', 'Target stops'])
        return df
    except Exception as e:
        st.warning(f"Couldn't load targets.xlsx: {str(e)}")
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
        'Cnee Str #': 'HOUSE_NUMBER',
        'PCC': 'PCC'
    }

    new_df = pd.DataFrame()
    for orig_col, new_col in column_mapping.items():
        if orig_col in df.columns:
            new_df[new_col] = df[orig_col]
        else:
            if new_col in ['HWB', 'CONSIGNEE_ZIP', 'WEIGHT', 'PIECES', 'PCC']:
                new_df[new_col] = None

    if 'PIECES' in new_df.columns:
        new_df['PIECES'] = new_df['PIECES'].apply(parse_pieces)

    if 'PCC' in new_df.columns:
        new_df['PCC'] = new_df['PCC'].astype(str).str.strip().str.upper()

    street1 = new_df['CONSIGNEE_STREET1'] if 'CONSIGNEE_STREET1' in new_df.columns else pd.Series('', index=new_df.index)
    street2 = new_df['CONSIGNEE_STREET2'] if 'CONSIGNEE_STREET2' in new_df.columns else pd.Series('', index=new_df.index)
    new_df['CONSIGNEE_STREET'] = street1.astype(str).fillna('') + ' ' + street2.astype(str).fillna('')
    new_df['CONSIGNEE_STREET'] = new_df['CONSIGNEE_STREET'].str.strip().replace('nan', '')

    new_df['HOUSE_NUMBER'] = new_df['CONSIGNEE_STREET'].apply(extract_house_number)
    new_df['HOUSE_NUMBER_FLOAT'] = new_df['HOUSE_NUMBER'].apply(house_number_to_float)
    new_df['STREET_NAME'] = new_df['CONSIGNEE_STREET'].apply(clean_street_name)

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

        if zip_code in special_zips:
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
                    if matched:
                        continue

            if zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[
                    fallback_routes['ZIP'] == zip_code, 'ROUTE'
                ].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
                continue

        else:
            if zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[
                    fallback_routes['ZIP'] == zip_code, 'ROUTE'
                ].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
                continue

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
                    if matched:
                        continue

    return manifest_df

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

def generate_reports(
    manifest_df, output_path, weight_thr=70, vol_weight_thr=150, pieces_thr=6,
    vehicle_weight_thr=70, vehicle_vol_thr=150, vehicle_pieces_thr=12,
    vehicle_kg_per_piece_thr=10, vehicle_van_max_pieces=20
):
    hwb_aggregated = manifest_df.groupby(['HWB', 'MATCHED_ROUTE']).agg({
        'CONSIGNEE_NAME_NORM': 'first',
        'CONSIGNEE_NAME': 'first',
        'CONSIGNEE_ZIP': 'first',
        'CONSIGNEE_ADDRESS': 'first',
        'WEIGHT': 'sum',
        'VOLUMETRIC_WEIGHT': 'sum',
        'PIECES': 'first'
    }).reset_index()

    route_summary = hwb_aggregated.groupby('MATCHED_ROUTE').agg(
        total_shipments=('HWB', 'count'),
        unique_consignees=('CONSIGNEE_NAME_NORM', 'nunique'),
        total_weight=('WEIGHT', 'sum'),
        total_pieces=('PIECES', 'sum')
    ).reset_index().rename(columns={'MATCHED_ROUTE': 'ROUTE'})

    targets_df = load_targets('input/targets.xlsx')
    if not targets_df.empty:
        route_summary = pd.merge(route_summary, targets_df, on='ROUTE', how='left')
    else:
        route_summary['Average PU stops'] = 0
        route_summary['Target stops'] = 0

    route_summary['total_weight'] = route_summary['total_weight'].round(1)
    route_summary['Predicted Stops'] = route_summary['unique_consignees'] + route_summary['Average PU stops']
    route_summary['Predicted - Target'] = route_summary['Predicted Stops'] - route_summary['Target stops']
    for col in ['Average PU stops', 'Predicted Stops', 'Predicted - Target']:
        route_summary[col] = route_summary[col].round(0).astype('Int64')
    route_summary.insert(5, '', '')
    route_summary = route_summary[['ROUTE', 'total_shipments', 'unique_consignees', 'total_weight', 'total_pieces', '',
                                  'Average PU stops', 'Predicted Stops', 'Target stops', 'Predicted - Target']]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    os.makedirs(output_path, exist_ok=True)
    summary_path = f"{output_path}/route_summary_{timestamp}.xlsx"

    with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
        route_summary.to_excel(writer, sheet_name='Summary', index=False)
        sheet = writer.sheets['Summary']
        current_row = sheet.max_row + 2
        sheet.cell(row=current_row, column=1, value="PCC Statistics:")
        current_row += 1
        sheet.cell(row=current_row, column=1, value="Product")
        sheet.cell(row=current_row, column=2, value="Shipments")
        sheet.cell(row=current_row, column=3, value="Pieces")
        sheet.cell(row=current_row, column=4, value="Pieces/Shipment")
        current_row += 1
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
            sheet.cell(row=current_row, column=1, value=label)
            sheet.cell(row=current_row, column=2, value=shipments)
            sheet.cell(row=current_row, column=3, value=pieces)
            sheet.cell(row=current_row, column=4, value=ratio)
            current_row += 1

        # ZIP Code Statistics section
        zip_stats = manifest_df.groupby('CONSIGNEE_ZIP').agg(
            total_shipments=('HWB', 'nunique'),
            unique_consignees=('CONSIGNEE_NAME_NORM', 'nunique')
        ).reset_index().sort_values('CONSIGNEE_ZIP', ascending=True)
        current_row += 2
        sheet.cell(row=current_row, column=1, value="ZIP Code Statistics:")
        current_row += 1
        sheet.cell(row=current_row, column=1, value="ZIP Code")
        sheet.cell(row=current_row, column=2, value="Shipments")
        sheet.cell(row=current_row, column=3, value="Unique Consignees")
        current_row += 1
        for _, row in zip_stats.iterrows():
            sheet.cell(row=current_row, column=1, value=row['CONSIGNEE_ZIP'])
            sheet.cell(row=current_row, column=2, value=row['total_shipments'])
            sheet.cell(row=current_row, column=3, value=row['unique_consignees'])
            current_row += 1

        add_target_conditional_formatting(sheet, 'J', 2, len(route_summary)+1)

    # ... (rest of your report generation code for special cases, MBX, etc.) ...

    return timestamp, route_summary, None

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
        timestamp, route_summary, _ = generate_reports(
            matched_manifest, output_path,
            weight_thr, vol_weight_thr, pieces_thr,
            vehicle_weight_thr, vehicle_vol_thr,
            vehicle_pieces_thr, vehicle_kg_per_piece_thr, vehicle_van_max_pieces
        )
        try:
            if not route_summary.empty and 'Predicted Stops' in route_summary:
                predicted_spr = route_summary['Predicted Stops'].mean()
                st.metric("Predicted SPR (Average Predicted Stops)", f"{predicted_spr:.1f}")
            else:
                st.warning("No routes matched - cannot calculate SPR")
        except Exception as e:
            st.error(f"SPR calculation error: {str(e)}")
        st.success("Processing complete! 🎉")
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
        st.subheader("Preview of Processed Data")
        st.dataframe(matched_manifest.head(10))

if __name__ == "__main__":
    main()
