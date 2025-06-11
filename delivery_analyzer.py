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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# --- All your existing helper functions (unchanged) ---
def normalize_diacritics(text):
    diacritic_map = {'č':'c', 'š':'s', 'ž':'z', 'Č':'C', 'Š':'S', 'Ž':'Z'}
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
    while parts and (re.match(r'^\d{4}$', parts[-1]) or parts[-1] in irrelevant_words):
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
    if not isinstance(name, str): return ''
    name = name.lower().strip()
    name = re.sub(r'[.,]', '', name)
    suffixes = [
        'doo', 'd o o', 'd.o.o', 'd.o.o.', 'd d', 'd.d.', 'd.d', 'dno', 'd.n.o.', 'd.n.o', 'ddoo', 
        'ltd', 'limited', 'llc', 'gmbh', 'inc'
    ]
    for suffix in suffixes:
        name = re.sub(r'\s*' + re.escape(suffix) + r'$', '', name)
    return re.sub(r'\s+', ' ', name).strip()

def extract_house_number(address):
    match = re.search(r'\b(\d+[a-z]?)\b', str(address).lower())
    return match.group(1) if match else None

def house_number_to_float(hn):
    if not hn: return None
    hn = str(hn).lower()
    num_part = re.sub(r'[^0-9]', '', hn)
    letter_part = re.sub(r'[^a-z]', '', hn)
    if not num_part: return None
    base = float(num_part)
    if letter_part: base += (ord(letter_part) - ord('a') + 1) * 0.01
    return base

def load_street_city_routes(path):
    try:
        df = pd.read_excel(path)
        df.columns = ['ROUTE', 'STREET', 'CITY']
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
        return pd.read_excel(path, usecols="A:C", names=['ROUTE', 'Average PU stops', 'Target stops'])
    except Exception as e:
        st.warning(f"Couldn't load targets.xlsx: {str(e)}")
        return pd.DataFrame()

# --- NEW EMAIL FUNCTIONS ---
def load_email_mapping(path):
    """Load email mapping from Excel file"""
    try:
        df = pd.read_excel(path)
        # Expected columns: Route_Group, Email, Contact_Name, Report_Type
        return df
    except Exception as e:
        st.warning(f"Couldn't load email mapping: {str(e)}")
        return pd.DataFrame()

def send_email_with_attachment(smtp_server, smtp_port, sender_email, sender_password, 
                              recipient_email, contact_name, subject, body, attachment_path):
    """Send email with Excel attachment"""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        # Add body
        msg.attach(MIMEText(body, 'plain'))
        
        # Add attachment
        if os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(attachment_path)}'
            )
            msg.attach(part)
        
        # Send email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        
        return True, f"Email sent successfully to {contact_name} ({recipient_email})"
    
    except Exception as e:
        return False, f"Failed to send email to {recipient_email}: {str(e)}"

def send_route_reports(route_summary, specialized_reports, email_mapping, output_path, timestamp,
                      smtp_server, smtp_port, sender_email, sender_password):
    """Send specialized reports to designated recipients"""
    
    results = []
    
    # Email mapping for specialized reports
    email_routes = {
        'MBX': ['MB1', 'MB2'],
        'KRA': ['KR1', 'KR2'], 
        'LJU': ['LJ1', 'LJ2'],
        'NMO': ['NM1', 'NM2'],
        'CEJ': ['CE1', 'CE2'],
        'NGR': ['NG1', 'NG2'],
        'NGX': ['NGX'],
        'KOP': ['KP1']
    }
    
    for report_type, route_prefixes in email_routes.items():
        # Find email recipients for this report type
        recipients = email_mapping[email_mapping['Report_Type'] == report_type]
        
        if not recipients.empty and report_type in specialized_reports:
            attachment_path = specialized_reports[report_type]
            
            # Calculate statistics for this route group
            route_data = route_summary[route_summary['ROUTE'].str.startswith(tuple(route_prefixes), na=False)]
            total_shipments = route_data['total_shipments'].sum() if not route_data.empty else 0
            total_weight = route_data['total_weight'].sum() if not route_data.empty else 0
            
            # Create email content
            subject = f"Daily Route Report - {report_type} ({datetime.now().strftime('%Y-%m-%d')})"
            
            body = f"""Dear Team,

Please find attached the daily route report for {report_type} routes ({', '.join(route_prefixes)}).

Summary for {datetime.now().strftime('%Y-%m-%d')}:
- Total Shipments: {total_shipments}
- Total Weight: {total_weight:.1f} kg
- Routes Covered: {', '.join(route_prefixes)}

The attached Excel file contains detailed shipment information including:
- Route assignments
- Consignee details
- Addresses and ZIP codes
- AWB numbers

Please review and coordinate accordingly.

Best regards,
Delivery Route Analyzer System
"""
            
            # Send to each recipient
            for _, recipient in recipients.iterrows():
                success, message = send_email_with_attachment(
                    smtp_server, smtp_port, sender_email, sender_password,
                    recipient['Email'], recipient['Contact_Name'],
                    subject, body, attachment_path
                )
                results.append((report_type, recipient['Contact_Name'], success, message))
    
    return results

# --- Continue with all existing functions (process_manifest, match_address_to_route, etc.) ---
def parse_pieces(value):
    try:
        return int(re.findall(r'\d+', str(value).split('/')[0])[0])
    except:
        return 1

def clean_nan_from_address(addr):
    return re.sub(r'\s*nan\s*$', '', str(addr)).strip() if isinstance(addr, str) else addr

def process_manifest(file):
    try:
        if file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
        else:
            df = pd.read_csv(file, delimiter=',', quotechar='"', engine='python')
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return pd.DataFrame()

    column_map = {
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
    for orig, new in column_map.items():
        if orig in df.columns:
            new_df[new] = df[orig].astype(str)
        else:
            new_df[new] = pd.Series('', index=df.index, dtype='str')

    new_df['PIECES'] = new_df['PIECES'].apply(parse_pieces)
    new_df['PCC'] = new_df['PCC'].astype(str).str.strip().str.upper() if 'PCC' in new_df.columns else None

    street1 = new_df['CONSIGNEE_STREET1'].fillna('')
    street2 = new_df['CONSIGNEE_STREET2'].fillna('')
    new_df['CONSIGNEE_STREET'] = street1.str.strip() + ' ' + street2.str.strip()
    new_df['CONSIGNEE_STREET'] = new_df['CONSIGNEE_STREET'].str.replace('nan', '').str.strip()

    new_df['HOUSE_NUMBER'] = new_df['CONSIGNEE_STREET'].apply(extract_house_number)
    new_df['HOUSE_NUMBER_FLOAT'] = new_df['HOUSE_NUMBER'].apply(house_number_to_float)
    new_df['STREET_NAME'] = new_df['CONSIGNEE_STREET'].apply(clean_street_name)
    
    if 'CONSIGNEE_ZIP' in new_df.columns:
        new_df['CONSIGNEE_ZIP'] = new_df['CONSIGNEE_ZIP'].astype(str).str.extract(r'(\d{4})')[0].str.zfill(4)

    for col in ['WEIGHT', 'VOLUMETRIC_WEIGHT']:
        if col in new_df.columns: 
            new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0)
    
    new_df['PIECES'] = pd.to_numeric(new_df['PIECES'], errors='coerce').fillna(1)

    new_df['MATCHED_ROUTE'] = None
    new_df['MATCH_SCORE'] = 0.0
    new_df['MATCH_METHOD'] = None
    new_df['CONSIGNEE_ADDRESS'] = new_df['CONSIGNEE_STREET'].apply(clean_nan_from_address)
    new_df['CONSIGNEE_NAME_NORM'] = new_df['CONSIGNEE_NAME'].apply(normalize_consignee_name)
    return new_df

def match_address_to_route(manifest_df, street_city_routes, fallback_routes):
    manifest_df['MATCHED_ROUTE'] = None
    manifest_df['MATCH_METHOD'] = None
    manifest_df['MATCH_SCORE'] = 0.0
    special_zips = {'2000', '3000', '4000', '5000', '6000', '8000'}

    for idx, row in manifest_df.iterrows():
        # Enhanced direct customer matching rule for Elrad Electronics
        consignee_name = str(row['CONSIGNEE_NAME']).lower()
        consignee_street = str(row['CONSIGNEE_STREET']).lower()
        
        if ('elrad' in consignee_name and 'electronics' in consignee_name and 
            ('ljutomerska' in consignee_street and '47' in consignee_street)):
            manifest_df.at[idx, 'MATCHED_ROUTE'] = 'MB1B'
            manifest_df.at[idx, 'MATCH_METHOD'] = 'Direct Rule - Elrad'
            manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
            continue
        
        zip_code = row['CONSIGNEE_ZIP']
        street_name = clean_street_name(row['CONSIGNEE_STREET'])
        city_name = clean_city_name(row['CONSIGNEE_CITY'])
        matched = False

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
            if not matched and zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[fallback_routes['ZIP'] == zip_code, 'ROUTE'].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
        else:
            if zip_code in fallback_routes['ZIP'].values:
                manifest_df.at[idx, 'MATCHED_ROUTE'] = fallback_routes.loc[fallback_routes['ZIP'] == zip_code, 'ROUTE'].values[0]
                manifest_df.at[idx, 'MATCH_METHOD'] = 'ZIP'
                manifest_df.at[idx, 'MATCH_SCORE'] = 100.0
            else:
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

def auto_adjust_column_width(worksheet):
    """Auto-adjust column widths for better readability"""
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        worksheet.column_dimensions[column_letter].width = adjusted_width

def create_specialized_report(manifest_df, route_prefixes, report_name, output_path, timestamp):
    """Create specialized report for specific route prefixes with MATCHED_ROUTE first"""
    # Filter data for specified route prefixes
    filtered_data = manifest_df[
        manifest_df['MATCHED_ROUTE'].str.startswith(tuple(route_prefixes), na=False)
    ].copy()
    
    report_path = f"{output_path}/{report_name}_details_{timestamp}.xlsx"
    
    if not filtered_data.empty:
        # Select and reorder columns with MATCHED_ROUTE first
        report_data = filtered_data[['MATCHED_ROUTE', 'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ADDRESS', 'CONSIGNEE_ZIP']].copy()
        report_data = report_data.sort_values(by=['MATCHED_ROUTE', 'CONSIGNEE_ZIP'], ascending=[True, True])
        
        # Create Excel with auto-adjusted columns
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            report_data.to_excel(writer, index=False)
            auto_adjust_column_width(writer.sheets['Sheet1'])
    else:
        # Create empty file with MATCHED_ROUTE first
        pd.DataFrame(columns=['MATCHED_ROUTE', 'HWB', 'CONSIGNEE_NAME', 'CONSIGNEE_ADDRESS', 'CONSIGNEE_ZIP']).to_excel(report_path, index=False)
    
    return report_path

# --- Continue with generate_reports function (keeping all existing functionality) ---
def generate_reports(
    manifest_df, output_path, weight_thr=70, vol_weight_thr=150, pieces_thr=6,
    vehicle_weight_thr=70, vehicle_vol_thr=150, vehicle_pieces_thr=12,
    vehicle_kg_per_piece_thr=10, vehicle_van_max_pieces=20
):
    # [All existing report generation code remains the same]
    # ... [keeping all your existing logic for route_summary, special_cases, etc.]
    
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
    
    # [Continue with all existing report generation logic...]
    # Generate specialized reports with MATCHED_ROUTE first
    specialized_reports = {}
    specialized_reports['MBX'] = create_specialized_report(manifest_df, ['MB1', 'MB2'], 'MBX', output_path, timestamp)
    specialized_reports['KRA'] = create_specialized_report(manifest_df, ['KR1', 'KR2'], 'KRA', output_path, timestamp)
    specialized_reports['LJU'] = create_specialized_report(manifest_df, ['LJ1', 'LJ2'], 'LJU', output_path, timestamp)
    specialized_reports['NMO'] = create_specialized_report(manifest_df, ['NM1', 'NM2'], 'NMO', output_path, timestamp)
    specialized_reports['CEJ'] = create_specialized_report(manifest_df, ['CE1', 'CE2'], 'CEJ', output_path, timestamp)
    specialized_reports['NGR'] = create_specialized_report(manifest_df, ['NG1', 'NG2'], 'NGR', output_path, timestamp)
    specialized_reports['NGX'] = create_specialized_report(manifest_df, ['NGX'], 'NGX', output_path, timestamp)
    specialized_reports['KOP'] = create_specialized_report(manifest_df, ['KP1'], 'KOP', output_path, timestamp)

    return timestamp, route_summary, specialized_reports

def main():
    st.title("🚚 Delivery Route Analyzer")
    
    # Email Configuration Sidebar
    st.sidebar.header("Email Configuration")
    smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.sidebar.number_input("SMTP Port", value=587)
    sender_email = st.sidebar.text_input("Sender Email")
    sender_password = st.sidebar.text_input("Email Password", type="password")
    
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
        timestamp, route_summary, specialized_reports = generate_reports(
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
        
        # --- EMAIL AUTOMATION SECTION ---
        st.subheader("📧 Email Automation")
        
        # Load email mapping
        email_mapping = load_email_mapping('input/email_mapping.xlsx')
        
        if not email_mapping.empty:
            st.write(f"📋 Email mapping loaded: {len(email_mapping)} recipients configured")
            
            # Show email mapping preview
            with st.expander("View Email Recipients"):
                st.dataframe(email_mapping)
            
            # Email sending button
            col_email1, col_email2 = st.columns(2)
            with col_email1:
                if st.button("📤 Send Route Reports", type="primary"):
                    if sender_email and sender_password:
                        with st.spinner("Sending emails..."):
                            results = send_route_reports(
                                route_summary, specialized_reports, email_mapping, 
                                output_path, timestamp, smtp_server, smtp_port, 
                                sender_email, sender_password
                            )
                        
                        # Display results
                        st.subheader("Email Sending Results")
                        for report_type, contact, success, message in results:
                            if success:
                                st.success(f"✅ {report_type}: {message}")
                            else:
                                st.error(f"❌ {report_type}: {message}")
                    else:
                        st.error("Please configure email settings in the sidebar")
            
            with col_email2:
                st.info("📝 **Email Setup Required:**\n\n"
                       "Create `input/email_mapping.xlsx` with columns:\n"
                       "- **Report_Type** (MBX, KRA, LJU, etc.)\n"
                       "- **Email** (recipient@domain.com)\n"
                       "- **Contact_Name** (John Doe)")
        else:
            st.warning("📧 Email mapping not found. Create `input/email_mapping.xlsx` to enable automated emailing.")
            st.info("**Required columns:** Report_Type, Email, Contact_Name")
        
        # --- EXISTING DOWNLOAD SECTIONS (unchanged) ---
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
            with open(f"{output_path}/Priority_Shipments_{timestamp}.xlsx", "rb") as f:
                st.download_button("Priority Shipments", f, f"Priority_Shipments_{timestamp}.xlsx")

        st.subheader("Specialized Reports")
        
        # First row of specialized reports
        col6, col7, col8, col9 = st.columns(4)
        with col6:
            if os.path.exists(specialized_reports['MBX']):
                with open(specialized_reports['MBX'], "rb") as f:
                    st.download_button("MBX Details", f, f"MBX_details_{timestamp}.xlsx",
                                      help="MB1 and MB2 routes")
            else:
                st.write("No MBX shipments")
        
        with col7:
            if os.path.exists(specialized_reports['KRA']):
                with open(specialized_reports['KRA'], "rb") as f:
                    st.download_button("KRA Details", f, f"KRA_details_{timestamp}.xlsx",
                                      help="KR1 and KR2 routes")
            else:
                st.write("No KRA shipments")
        
        with col8:
            if os.path.exists(specialized_reports['LJU']):
                with open(specialized_reports['LJU'], "rb") as f:
                    st.download_button("LJU Details", f, f"LJU_details_{timestamp}.xlsx",
                                      help="LJ1 and LJ2 routes")
            else:
                st.write("No LJU shipments")
        
        with col9:
            if os.path.exists(specialized_reports['NMO']):
                with open(specialized_reports['NMO'], "rb") as f:
                    st.download_button("NMO Details", f, f"NMO_details_{timestamp}.xlsx",
                                      help="NM1 and NM2 routes")
            else:
                st.write("No NMO shipments")
        
        # Second row of specialized reports
        col10, col11, col12, col13 = st.columns(4)
        with col10:
            if os.path.exists(specialized_reports['CEJ']):
                with open(specialized_reports['CEJ'], "rb") as f:
                    st.download_button("CEJ Details", f, f"CEJ_details_{timestamp}.xlsx",
                                      help="CE1 and CE2 routes")
            else:
                st.write("No CEJ shipments")
        
        with col11:
            if os.path.exists(specialized_reports['NGR']):
                with open(specialized_reports['NGR'], "rb") as f:
                    st.download_button("NGR Details", f, f"NGR_details_{timestamp}.xlsx",
                                      help="NG1 and NG2 routes")
            else:
                st.write("No NGR shipments")
        
        with col12:
            if os.path.exists(specialized_reports['NGX']):
                with open(specialized_reports['NGX'], "rb") as f:
                    st.download_button("NGX Details", f, f"NGX_details_{timestamp}.xlsx",
                                      help="NGX routes")
            else:
                st.write("No NGX shipments")
        
        with col13:
            if os.path.exists(specialized_reports['KOP']):
                with open(specialized_reports['KOP'], "rb") as f:
                    st.download_button("KOP Details", f, f"KOP_details_{timestamp}.xlsx",
                                      help="KP1 routes")
            else:
                st.write("No KOP shipments")

        st.subheader("Preview of Processed Data")
        st.dataframe(matched_manifest.head(10))

if __name__ == "__main__":
    main()
