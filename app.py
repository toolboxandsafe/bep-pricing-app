"""
BEP Pricing Calculator - Streamlit Web App
Tool Box & Safe Moving - Vending Machine Move Pricing
Replicates Email-to-Trello Workflow with Google Maps Integration
"""

import streamlit as st
import json
import math
import re
import subprocess
import tempfile
import os
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
from fpdf import FPDF
import requests
from PyPDF2 import PdfReader, PdfWriter

# Page config
st.set_page_config(
    page_title="BEP Pricing Calculator",
    page_icon="🚚",
    layout="wide"
)

# Constants
HQ_ADDRESS = "Gilbert, AZ 85295"
HOURLY_RATE = 170
BUFFER_THRESHOLD_MILES = 35
BUFFER_MINUTES = 20
JOB_TIME_PER_MACHINE = 30  # minutes
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")

# =============================================================================
# EXCEL PARSING - BEP REQUEST TAB (IMPROVED)
# =============================================================================

def normalize_address(address):
    """Normalize address for deduplication"""
    if not address:
        return ""
    
    addr = str(address).upper().strip()
    
    # Remove facility prefixes
    prefixes = ['BEP ', 'MAXIMUS ', 'DES ', 'ADES ', 'DCS ']
    for prefix in prefixes:
        if addr.startswith(prefix):
            addr = addr[len(prefix):]
    
    # Extract street address pattern
    street_match = re.search(r'(\d+\s+[A-Z0-9\s\.]+(?:STREET|ST|AVENUE|AVE|BLVD|ROAD|RD|DRIVE|DR|LANE|LN|WAY|PKWY|HWY|CAMELBACK|WASHINGTON|JEFFERSON|VAN BUREN|INDIAN SCHOOL)[\.]*)', addr)
    if street_match:
        addr = street_match.group(1)
    
    # Standardize
    addr = addr.replace(' STREET', ' ST').replace(' AVENUE', ' AVE')
    addr = addr.replace(' BOULEVARD', ' BLVD').replace(' ROAD', ' RD')
    addr = addr.replace('.', '').replace(',', '').replace('#', ' ')
    addr = re.sub(r'\s+', ' ', addr).strip()
    
    return addr

def parse_bep_excel_v2(uploaded_file):
    """
    Parse BEP Move Request Excel - Extract machines with pickup/delivery pairs
    IMPORTANT: Stop parsing machines when we hit "Other Notes" section
    """
    try:
        # Read REQUEST tab
        try:
            df = pd.read_excel(uploaded_file, sheet_name='REQUEST', header=None)
        except:
            df = pd.read_excel(uploaded_file, header=None)
        
        result = {
            "success": True,
            "requester": None,
            "mr_number": None,
            "machines": [],
            "other_notes": None,
            "move_date": None,
            "contacts": [],
            "raw_data": []
        }
        
        # Convert to list of rows
        rows = []
        for idx, row in df.iterrows():
            row_data = [str(c).strip() if pd.notna(c) else "" for c in row]
            rows.append(row_data)
            result["raw_data"].append(" | ".join([c for c in row_data if c]))
        
        # FIRST PASS: Find key section boundaries
        other_notes_row = None
        machine_section_start = None
        
        for row_idx, row in enumerate(rows):
            row_text = " ".join(row).upper()
            
            # Find where "Comment" section starts - this is END of machine data
            if "COMMENT" in row_text:
                other_notes_row = row_idx
                # Capture ALL content below Comments as other_notes
                notes_lines = []
                for check_row in rows[row_idx:]:
                    for cell in check_row:
                        if cell and len(cell) > 3 and 'NAN' not in cell.upper():
                            # Skip the "Comments:" label itself
                            if cell.upper().strip() not in ['COMMENT', 'COMMENTS', 'COMMENTS:']:
                                notes_lines.append(cell.strip())
                result["other_notes"] = " | ".join(notes_lines) if notes_lines else None
                break
            
            # Also check for "Other Notes" as alternate boundary
            if "OTHER NOTE" in row_text or "SPECIAL INSTRUCTION" in row_text:
                if other_notes_row is None:
                    other_notes_row = row_idx
            
            # Find where machine data starts (first numbered item or "PICK UP" header)
            if machine_section_start is None:
                if "ITEMS TO" in row_text or "ITEM TO" in row_text:
                    machine_section_start = row_idx
        
        # Find MR number
        for row in rows:
            for cell in row:
                mr_match = re.search(r'(1\w{2}-\d{2}/\d{2})', cell)
                if mr_match:
                    result["mr_number"] = mr_match.group(1)
                    break
        
        # Find requester - try multiple strategies
        # Strategy 1: Look for "Requester Name" label and grab adjacent cell
        for row_idx, row in enumerate(rows):
            row_text = " ".join(row).upper()
            if "REQUESTER" in row_text and "NAME" in row_text:
                # Check row above for the actual name (sometimes name is above label)
                if row_idx > 0:
                    for cell in rows[row_idx - 1]:
                        if cell and len(cell) > 2 and 'NAN' not in cell.upper():
                            if not any(x in cell.upper() for x in ['REQUESTER', 'DATE', 'SIGNATURE', 'FACILITY']):
                                result["requester"] = cell
                                break
                # Also check same row for name in adjacent cell
                if not result["requester"]:
                    for col_idx, cell in enumerate(row):
                        if "REQUESTER" in cell.upper():
                            # Get next non-empty cell
                            for next_cell in row[col_idx+1:]:
                                if next_cell and len(next_cell) > 2 and 'NAN' not in next_cell.upper():
                                    result["requester"] = next_cell
                                    break
                            break
                break
        
        # Strategy 2: Fallback to row 47 (index 46)
        if not result["requester"] and len(rows) > 46:
            for cell in rows[46][:4]:
                if cell and len(cell) > 2 and not any(x in cell.upper() for x in ['REQUESTER', 'NAME', 'DATE', 'SIGNATURE', 'NAN']):
                    result["requester"] = cell
                    break
        
        # SECOND PASS: Extract machines ONLY up to "Other Notes" section
        end_row = other_notes_row if other_notes_row else len(rows)
        start_row = machine_section_start if machine_section_start else 0
        
        current_machine = None
        current_section = None
        
        for row_idx in range(start_row, end_row):
            row = rows[row_idx]
            row_text = " ".join(row).upper()
            
            # Skip if we've hit other notes
            if "OTHER NOTE" in row_text:
                break
            
            # Check for machine number (must be in specific column position, not random numbers)
            # Machine numbers are typically in column A or B (index 0 or 1)
            for col_idx in range(min(2, len(row))):
                cell = row[col_idx].strip()
                # Must be standalone digit 1-10, not part of address or date
                if cell in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']:
                    # Verify this row also has pickup/delivery context nearby
                    if any(kw in row_text for kw in ['PICK', 'DELIVER', 'SITE', 'ITEM', 'MACHINE', 'VEND', 'COMBO', 'KIOSK']):
                        if current_machine and (current_machine.get("pickup") or current_machine.get("delivery")):
                            result["machines"].append(current_machine)
                        current_machine = {
                            "number": int(cell),
                            "pickup": None,
                            "pickup_address": None,
                            "delivery": None,
                            "delivery_address": None,
                            "type": None
                        }
                        break
            
            # Extract data based on EXACT row labels
            # "Pick up site:" row contains pickup address
            # "Delivery site:" row contains delivery address  
            # "Items to be moved" row contains machine type (NOT an address)
            
            if current_machine:
                for col_idx, cell in enumerate(row):
                    cell_upper = cell.upper() if cell else ""
                    
                    # PICKUP ADDRESS: Look for "Pick up site" label, grab address from same row
                    if "PICK UP SITE" in cell_upper or "PICKUP SITE" in cell_upper:
                        # Address should be in next cell(s) on same row
                        for addr_cell in row[col_idx+1:]:
                            if addr_cell and len(addr_cell) > 5 and 'NAN' not in addr_cell.upper():
                                if not current_machine["pickup"]:
                                    current_machine["pickup"] = addr_cell
                                    current_machine["pickup_address"] = normalize_address(addr_cell)
                                break
                    
                    # DELIVERY ADDRESS: Look for "Delivery site" label, grab address from same row
                    elif "DELIVERY SITE" in cell_upper or "DELIVER SITE" in cell_upper or "DELIVER TO" in cell_upper:
                        # Address should be in next cell(s) on same row
                        for addr_cell in row[col_idx+1:]:
                            if addr_cell and len(addr_cell) > 5 and 'NAN' not in addr_cell.upper():
                                if not current_machine["delivery"]:
                                    current_machine["delivery"] = addr_cell
                                    current_machine["delivery_address"] = normalize_address(addr_cell)
                                break
                    
                    # MACHINE TYPE: Look for "Items to be moved" label, grab type from same row
                    elif "ITEM" in cell_upper and "MOVE" in cell_upper:
                        # Machine type should be in next cell(s) on same row
                        for type_cell in row[col_idx+1:]:
                            if type_cell and len(type_cell) > 2 and 'NAN' not in type_cell.upper():
                                if not current_machine["type"]:
                                    current_machine["type"] = type_cell
                                break
        
        # Add last machine if valid
        if current_machine and (current_machine.get("pickup") or current_machine.get("delivery")):
            result["machines"].append(current_machine)
        
        # Fallback: If no numbered machines found, try alternate method
        if not result["machines"]:
            result["machines"] = extract_machines_alternate(rows[:end_row])
        
        # Deduplicate
        result["unique_pickups"] = list(set([m["pickup_address"] for m in result["machines"] if m.get("pickup_address")]))
        result["unique_deliveries"] = list(set([m["delivery_address"] for m in result["machines"] if m.get("delivery_address")]))
        
        return result
        
    except Exception as e:
        return {"success": False, "error": str(e)}

def extract_machines_alternate(rows):
    """Alternate extraction when numbered sections aren't found"""
    machines = []
    pickups = []
    deliveries = []
    current_section = None
    
    for row in rows:
        row_text = " ".join(row).upper()
        
        if "PICK UP" in row_text or "PICKUP" in row_text:
            current_section = "pickup"
            continue
        elif "DELIVER" in row_text or "DELIVERY" in row_text:
            current_section = "delivery"
            continue
        
        for cell in row:
            if not cell or cell.upper() == 'NAN':
                continue
            cell_upper = cell.upper()
            
            is_address = any(ind in cell_upper for ind in [
                'AVE', 'STREET', 'ST ', 'BLVD', 'RD ', 'DRIVE', 'DR ', 'PHOENIX', 
                'PHX', 'TUCSON', 'MAXIMUS', 'DES ', 'BEP '
            ]) or re.search(r'\d{4,5}\s+[A-Z]', cell_upper)
            
            if is_address:
                if current_section == "pickup":
                    pickups.append(cell)
                elif current_section == "delivery":
                    deliveries.append(cell)
    
    # Pair pickups with deliveries
    num_machines = max(len(pickups), len(deliveries), 1)
    for i in range(num_machines):
        machines.append({
            "number": i + 1,
            "pickup": pickups[i] if i < len(pickups) else None,
            "pickup_address": normalize_address(pickups[i]) if i < len(pickups) else None,
            "delivery": deliveries[i] if i < len(deliveries) else None,
            "delivery_address": normalize_address(deliveries[i]) if i < len(deliveries) else None,
            "type": None
        })
    
    return machines

# =============================================================================
# GOOGLE MAPS DISTANCE MATRIX API
# =============================================================================

def get_distance_matrix(origins, destinations):
    """Call Google Maps Distance Matrix API"""
    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    
    # Simple params without departure_time (avoids billing/permission issues)
    params = {
        "origins": "|".join(origins),
        "destinations": "|".join(destinations),
        "key": GOOGLE_MAPS_API_KEY,
        "units": "imperial"
    }
    
    try:
        response = requests.get(url, params=params)
        data = response.json()
        
        if data["status"] == "OK":
            return data
        else:
            st.error(f"Google Maps API error: {data.get('status')}")
            return None
    except Exception as e:
        st.error(f"API request failed: {e}")
        return None

def calculate_route(pickups, deliveries):
    """
    Calculate sequential route: HQ → Pickups → Deliveries → HQ
    Returns list of legs with distance and duration
    """
    # Build route
    route = [HQ_ADDRESS]
    route.extend(pickups)
    route.extend(deliveries)
    route.append(HQ_ADDRESS)
    
    legs = []
    total_duration_minutes = 0
    max_distance_miles = 0
    
    # Calculate each consecutive leg
    for i in range(len(route) - 1):
        origin = route[i]
        destination = route[i + 1]
        
        data = get_distance_matrix([origin], [destination])
        
        if data and data["rows"][0]["elements"][0]["status"] == "OK":
            element = data["rows"][0]["elements"][0]
            
            # Distance in miles
            distance_meters = element["distance"]["value"]
            distance_miles = distance_meters / 1609.34
            
            # Duration in minutes
            duration_seconds = element["duration"]["value"]
            duration_minutes = duration_seconds / 60
            
            legs.append({
                "from": origin,
                "to": destination,
                "distance_miles": round(distance_miles, 1),
                "duration_minutes": round(duration_minutes, 1)
            })
            
            total_duration_minutes += duration_minutes
            max_distance_miles = max(max_distance_miles, distance_miles)
        else:
            st.warning(f"Could not calculate: {origin} → {destination}")
            # Estimate fallback
            legs.append({
                "from": origin,
                "to": destination,
                "distance_miles": 20,
                "duration_minutes": 30,
                "estimated": True
            })
            total_duration_minutes += 30
    
    return {
        "legs": legs,
        "total_duration_minutes": round(total_duration_minutes, 1),
        "max_distance_miles": round(max_distance_miles, 1),
        "route": route
    }

# =============================================================================
# QUOTE CALCULATION
# =============================================================================

def calculate_quote(route_data, num_machines, pickups, deliveries):
    """Calculate BEP quote following workflow rules"""
    
    drive_time = route_data["total_duration_minutes"]
    max_distance = route_data["max_distance_miles"]
    
    # Job time: 30 minutes per machine
    job_time = JOB_TIME_PER_MACHINE * num_machines
    
    # Buffer: 20 minutes if max distance > 35 miles
    if max_distance > BUFFER_THRESHOLD_MILES:
        buffer_time = BUFFER_MINUTES
        no_buffer_discount = 0
    else:
        buffer_time = 0
        no_buffer_discount = 60  # -$60 for short trips
    
    # Total time
    total_minutes = drive_time + job_time + buffer_time
    total_hours = total_minutes / 60
    
    # Base price
    base_price = total_hours * HOURLY_RATE
    
    # Apply no-buffer discount
    base_price -= no_buffer_discount
    
    # Check locations for minimums
    all_locations = " ".join(pickups + deliveries).upper()
    
    is_tucson = "TUCSON" in all_locations
    is_prison = any(kw in all_locations for kw in ['ASPC', 'PRISON', 'CORRECTIONAL', 'CIMARRON', 'FLORENCE'])
    
    # Determine minimum
    if is_tucson:
        min_price = 850
    elif is_prison:
        min_price = 900
    else:
        min_price = 220
    
    # Apply minimum
    final_price = max(base_price, min_price)
    
    # Round UP to nearest $25
    final_price = math.ceil(final_price / 25) * 25
    
    return {
        "drive_time": round(drive_time, 1),
        "job_time": job_time,
        "buffer_time": buffer_time,
        "no_buffer_discount": no_buffer_discount,
        "total_minutes": round(total_minutes, 1),
        "total_hours": round(total_hours, 2),
        "max_distance_miles": round(max_distance, 1),
        "base_price": round(base_price, 2),
        "min_price": min_price,
        "final_price": int(final_price),
        "is_tucson": is_tucson,
        "is_prison": is_prison,
        "formula": f"({round(drive_time)} + {job_time} + {buffer_time}) ÷ 60 × ${HOURLY_RATE} - ${no_buffer_discount} = ${base_price:.0f}"
    }

# =============================================================================
# EXCEL TO PDF CONVERSION
# =============================================================================

def convert_excel_to_pdf(excel_bytes, filename="request.xlsx"):
    """Convert Excel file to PDF using LibreOffice"""
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, filename)
        with open(excel_path, 'wb') as f:
            f.write(excel_bytes)
        
        try:
            result = subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', tmpdir, excel_path
            ], capture_output=True, text=True, timeout=60)
            
            pdf_name = os.path.splitext(filename)[0] + '.pdf'
            pdf_path = os.path.join(tmpdir, pdf_name)
            
            if os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    return f.read()
            return None
        except Exception as e:
            st.error(f"LibreOffice conversion failed: {e}")
            return None

def remove_pdf_pages(pdf_bytes, pages_to_remove=[1]):
    """Remove specific pages from PDF (0-indexed)"""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        writer = PdfWriter()
        
        for i, page in enumerate(reader.pages):
            if i not in pages_to_remove:
                writer.add_page(page)
        
        output = BytesIO()
        writer.write(output)
        return output.getvalue()
    except Exception as e:
        st.error(f"Error removing pages: {e}")
        return pdf_bytes

# =============================================================================
# TRELLO INTEGRATION
# =============================================================================

def create_trello_card(data, api_key, api_token, list_id):
    """Create Trello card with quote data"""
    
    # Build driving stops
    stops = ["HQ"]
    for p in data.get("unique_pickups", []):
        short_name = p.split()[0][:6].upper() if p else "?"
        stops.append(short_name)
    for d in data.get("unique_deliveries", []):
        short_name = d.split()[0][:6].upper() if d else "?"
        stops.append(short_name)
    stops.append("HQ")
    driving_stops = " - ".join(stops)
    
    title = f"{data.get('requester', 'Unknown')} - {data.get('mr_number', 'BEP')} - ${data.get('final_price', 0)}"
    
    desc = f"""## Move Request Quote

**MR Number:** {data.get('mr_number', 'N/A')}
**Requester:** {data.get('requester', 'N/A')}
**Move Date:** {data.get('move_date', 'TBD')}
**Machines:** {data.get('num_machines', 1)}

---

### 📍 MACHINES & LOCATIONS
"""
    
    for m in data.get("machines", []):
        desc += f"\n**Machine {m.get('number', '?')}:** {m.get('type', 'Vending')}\n"
        desc += f"  - Pickup: {m.get('pickup', 'N/A')}\n"
        desc += f"  - Delivery: {m.get('delivery', 'N/A')}\n"
    
    desc += f"""
---

### 🚗 DRIVING STOPS
{driving_stops}

({len(data.get('unique_pickups', [])) + len(data.get('unique_deliveries', []))} unique stops)

---

### 💰 QUOTE: ${data.get('final_price', 0):,}

**Breakdown:**
- Drive Time: {data.get('drive_time', 0)} min
- Job Time: {data.get('job_time', 0)} min ({data.get('num_machines', 1)} machines × 30 min)
- Buffer: {data.get('buffer_time', 0)} min
- Max Distance Leg: {data.get('max_distance_miles', 0)} miles
- Total Hours: {data.get('total_hours', 0)}
- Rate: ${HOURLY_RATE}/hour

**Formula:** {data.get('formula', '')}

---

### 📝 OTHER NOTES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{data.get('other_notes', 'None')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

---
@luissaravia2
"""
    
    url = "https://api.trello.com/1/cards"
    params = {
        'key': api_key,
        'token': api_token,
        'idList': list_id,
        'name': title,
        'desc': desc,
        'pos': 'top'
    }
    
    response = requests.post(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        st.error(f"Trello API error {response.status_code}: {response.text}")
        return None

def attach_file_to_card(card_id, file_bytes, filename, mime_type, api_key, api_token):
    """Attach any file to a Trello card"""
    url = f"https://api.trello.com/1/cards/{card_id}/attachments"
    
    params = {
        'key': api_key,
        'token': api_token,
        'name': filename
    }
    
    files = {
        'file': (filename, file_bytes, mime_type)
    }
    
    response = requests.post(url, params=params, files=files)
    if response.status_code == 200:
        return True
    else:
        st.error(f"Attachment error: {response.status_code}: {response.text}")
        return False

def attach_pdf_to_card(card_id, pdf_bytes, filename, api_key, api_token):
    """Attach PDF file to a Trello card"""
    return attach_file_to_card(card_id, pdf_bytes, filename, 'application/pdf', api_key, api_token)

def attach_excel_to_card(card_id, excel_bytes, filename, api_key, api_token):
    """Attach Excel file to a Trello card"""
    return attach_file_to_card(card_id, excel_bytes, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', api_key, api_token)

def get_card_attachments(card_id, api_key, api_token):
    """Get attachments from a Trello card"""
    url = f"https://api.trello.com/1/cards/{card_id}/attachments"
    params = {'key': api_key, 'token': api_token}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    return []

def get_card_info(card_id, api_key, api_token):
    """Get card info including name"""
    url = f"https://api.trello.com/1/cards/{card_id}"
    params = {'key': api_key, 'token': api_token, 'fields': 'name,desc,shortUrl'}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    return None

def download_attachment(url):
    """Download file from URL"""
    response = requests.get(url)
    if response.status_code == 200:
        return response.content
    return None

def extract_quote_from_title(title):
    """Extract quote amount from card title like 'Name - MR# - $325'"""
    match = re.search(r'\$(\d+)', title)
    if match:
        return int(match.group(1))
    return None

def fill_worksheet_and_generate_pdf(excel_bytes, quote_amount, signature="Ryan Kearl"):
    """Fill Worksheet tab with hours and generate PDF"""
    import openpyxl
    
    # Calculate hours
    hours = round(quote_amount / HOURLY_RATE, 2)
    today_date = datetime.now().strftime("%m/%d/%Y")
    
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save Excel
        excel_path = os.path.join(tmpdir, "workbook.xlsx")
        with open(excel_path, 'wb') as f:
            f.write(excel_bytes)
        
        # Open and modify
        try:
            wb = openpyxl.load_workbook(excel_path)
            
            # Find Worksheet tab (might be named differently)
            ws_names = [s for s in wb.sheetnames if 'WORKSHEET' in s.upper() or 'WORK' in s.upper()]
            if ws_names:
                ws = wb[ws_names[0]]
            else:
                # Fallback to second sheet
                ws = wb.worksheets[1] if len(wb.worksheets) > 1 else wb.worksheets[0]
            
            # Fill hours in row 6, columns D and I
            ws['D6'] = hours
            ws['I6'] = hours
            
            # Fill signature and date
            ws['K19'] = signature
            ws['K21'] = today_date
            
            # Save modified workbook
            modified_path = os.path.join(tmpdir, "modified.xlsx")
            wb.save(modified_path)
            wb.close()
            
            # Convert to PDF using LibreOffice
            result = subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', tmpdir, modified_path
            ], capture_output=True, text=True, timeout=60)
            
            pdf_path = os.path.join(tmpdir, "modified.pdf")
            
            if os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    pdf_bytes = f.read()
                
                # Extract just the Worksheet page (usually page 2, index 1)
                try:
                    reader = PdfReader(BytesIO(pdf_bytes))
                    writer = PdfWriter()
                    
                    # Add only page 2 (Worksheet) if exists, else page 1
                    if len(reader.pages) > 1:
                        writer.add_page(reader.pages[1])
                    else:
                        writer.add_page(reader.pages[0])
                    
                    output = BytesIO()
                    writer.write(output)
                    return output.getvalue(), hours
                except:
                    return pdf_bytes, hours
            
            return None, hours
            
        except Exception as e:
            st.error(f"Error filling worksheet: {e}")
            return None, hours

# =============================================================================
# STREAMLIT UI
# =============================================================================

# Get Trello credentials once
trello_key = os.environ.get("TRELLO_API_KEY", "")
trello_token = os.environ.get("TRELLO_TOKEN", "")
trello_list = os.environ.get("TRELLO_LIST_ID", "699c9f9d6117bdcbb2d0e0aa")

# Sidebar - Page Navigation
with st.sidebar:
    st.title("🚚 BEP Tools")
    page = st.radio("Select Page:", ["📤 New Request", "📝 Generate Quote"], label_visibility="collapsed")
    
    st.divider()
    
    st.header("📋 Pricing Rules")
    st.markdown(f"""
    **Rate:** ${HOURLY_RATE}/hour
    **Job Time:** 30 min/machine
    **Buffer:** +20 min if >35 miles
    **No-buffer discount:** -$60 if ≤35 miles
    
    **Minimums:**
    - General: $220
    - Tucson: $850
    - Prison: $900+
    
    **Rounding:** Up to $25
    """)
    
    st.divider()
    
    with st.expander("🔧 Trello Settings", expanded=False):
        if trello_key and trello_token:
            st.success("✅ Trello credentials loaded")
        else:
            st.warning("⚠️ Set TRELLO_API_KEY and TRELLO_TOKEN in Railway variables")

# =============================================================================
# PAGE 2: GENERATE QUOTE PDF
# =============================================================================
if page == "📝 Generate Quote":
    st.title("📝 Generate Quote PDF")
    st.markdown("**After Ryan approves → Generate QUOTE PDF with filled worksheet**")
    
    st.divider()
    
    # Input: Card URL or ID
    card_input = st.text_input(
        "Trello Card URL or ID",
        placeholder="https://trello.com/c/ABC123 or card ID",
        help="Paste the Trello card URL or just the card ID"
    )
    
    # Extract card ID from URL
    card_id = None
    if card_input:
        # Handle full URL
        match = re.search(r'/c/([a-zA-Z0-9]+)', card_input)
        if match:
            card_id = match.group(1)
        else:
            # Assume it's just the ID
            card_id = card_input.strip()
    
    if card_id and trello_key and trello_token:
        # Fetch card info
        card_info = get_card_info(card_id, trello_key, trello_token)
        
        if card_info:
            st.success(f"✅ Found card: **{card_info.get('name')}**")
            
            # Extract quote from title
            auto_quote = extract_quote_from_title(card_info.get('name', ''))
            
            col1, col2 = st.columns(2)
            with col1:
                quote_amount = st.number_input(
                    "Quote Amount ($)",
                    min_value=100,
                    max_value=10000,
                    value=auto_quote or 300,
                    step=25
                )
            with col2:
                signature = st.text_input("Signature", value="Ryan Kearl")
            
            # Show calculated hours
            hours = round(quote_amount / HOURLY_RATE, 2)
            st.info(f"**Calculated Hours:** {hours} hrs (${quote_amount} ÷ ${HOURLY_RATE})")
            
            # Fetch attachments
            attachments = get_card_attachments(card_id, trello_key, trello_token)
            excel_attachments = [a for a in attachments if a.get('name', '').endswith(('.xlsx', '.xls'))]
            
            if excel_attachments:
                st.success(f"✅ Found Excel file: **{excel_attachments[0].get('name')}**")
                
                if st.button("🧮 Generate QUOTE PDF", type="primary", use_container_width=True):
                    with st.spinner("Downloading Excel..."):
                        excel_url = excel_attachments[0].get('url')
                        # Add auth to URL for Trello attachments
                        if '?' in excel_url:
                            excel_url += f"&key={trello_key}&token={trello_token}"
                        else:
                            excel_url += f"?key={trello_key}&token={trello_token}"
                        
                        excel_bytes = download_attachment(excel_url)
                    
                    if excel_bytes:
                        with st.spinner("Filling worksheet & generating PDF..."):
                            pdf_bytes, calc_hours = fill_worksheet_and_generate_pdf(
                                excel_bytes, quote_amount, signature
                            )
                        
                        if pdf_bytes:
                            st.success(f"✅ QUOTE PDF generated! ({calc_hours} hours)")
                            
                            # Show download button
                            st.download_button(
                                "⬇️ Download QUOTE PDF",
                                data=pdf_bytes,
                                file_name=f"QUOTE_{card_id}_{datetime.now().strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                            
                            # Attach to card
                            if st.button("📎 Attach QUOTE PDF to Card", use_container_width=True):
                                with st.spinner("Attaching..."):
                                    filename = f"QUOTE_{datetime.now().strftime('%Y%m%d')}.pdf"
                                    if attach_pdf_to_card(card_id, pdf_bytes, filename, trello_key, trello_token):
                                        st.success(f"✅ QUOTE PDF attached to card!")
                                        st.markdown(f"[Open Card]({card_info.get('shortUrl')})")
                                    else:
                                        st.error("Failed to attach PDF")
                        else:
                            st.error("Failed to generate PDF")
                    else:
                        st.error("Failed to download Excel file")
            else:
                st.warning("⚠️ No Excel file found on this card. Upload one first on Page 1.")
        else:
            st.error("❌ Card not found. Check the URL/ID.")
    elif card_id and not (trello_key and trello_token):
        st.error("⚠️ Trello credentials not configured")

# =============================================================================
# PAGE 1: NEW REQUEST (Original functionality)
# =============================================================================
else:
    st.title("📤 New Request")
    st.markdown("**Upload Excel → Calculate quote → Create Trello card**")
    
    # Main content
    st.header("Upload BEP Move Request")
    
    uploaded_file = st.file_uploader(
        "Drop your BEP Excel file here",
        type=["xlsx", "xls"],
        help="Upload the Move Request Excel file"
    )
    
    if uploaded_file:
        # Store Excel bytes for later attachment
        excel_bytes = uploaded_file.getvalue()
        excel_name = uploaded_file.name
        
        with st.spinner("Parsing Excel file..."):
            data = parse_bep_excel_v2(uploaded_file)
        
        if data["success"]:
            st.success(f"✅ Extracted {len(data['machines'])} machine(s)")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📋 Extracted Data")
                
                requester = st.text_input("Requester", value=data.get("requester") or "")
                mr_number = st.text_input("MR Number", value=data.get("mr_number") or "")
                move_date = st.text_input("Move Date", value=data.get("move_date") or "")
                
                st.markdown("### 🚛 Machines")
                
                # Editable machine list
                machines = data.get("machines", [])
                edited_machines = []
                
                for i, m in enumerate(machines):
                    with st.expander(f"Machine {m.get('number', i+1)}: {m.get('type', 'Vending')}", expanded=True):
                        pickup = st.text_input(f"Pickup {i+1}", value=m.get("pickup") or "", key=f"pickup_{i}")
                        delivery = st.text_input(f"Delivery {i+1}", value=m.get("delivery") or "", key=f"delivery_{i}")
                        mtype = st.text_input(f"Type {i+1}", value=m.get("type") or "Vending", key=f"type_{i}")
                        
                        edited_machines.append({
                            "number": i + 1,
                            "pickup": pickup,
                            "pickup_address": normalize_address(pickup),
                            "delivery": delivery,
                            "delivery_address": normalize_address(delivery),
                            "type": mtype
                        })
                
                # Add machine button
                if st.button("➕ Add Machine"):
                    edited_machines.append({
                        "number": len(edited_machines) + 1,
                        "pickup": "",
                        "pickup_address": "",
                        "delivery": "",
                        "delivery_address": "",
                        "type": "Vending"
                    })
                    st.rerun()
                
                other_notes = st.text_area("Other Notes", value=data.get("other_notes") or "", height=100)
            
            with col2:
                st.subheader("🗺️ Route & Quote")
                
                # Get unique addresses for routing
                unique_pickups = list(set([m["pickup"] for m in edited_machines if m.get("pickup")]))
                unique_deliveries = list(set([m["delivery"] for m in edited_machines if m.get("delivery")]))
                
                st.info(f"""
                **Route:** Gilbert, AZ 85295 → {len(unique_pickups)} pickup(s) → {len(unique_deliveries)} delivery(s) → HQ
                
                **Machines:** {len(edited_machines)}
                """)
                
                if st.button("🧮 CALCULATE ROUTE & QUOTE", type="primary", use_container_width=True):
                    if unique_pickups or unique_deliveries:
                        with st.spinner("Calculating route via Google Maps..."):
                            route_data = calculate_route(unique_pickups, unique_deliveries)
                            quote = calculate_quote(route_data, len(edited_machines), unique_pickups, unique_deliveries)
                            
                            st.session_state['route_data'] = route_data
                            st.session_state['quote'] = quote
                            st.session_state['excel_bytes'] = excel_bytes
                            st.session_state['excel_name'] = excel_name
                            st.session_state['full_data'] = {
                                "requester": requester,
                                "mr_number": mr_number,
                                "move_date": move_date,
                                "machines": edited_machines,
                                "unique_pickups": unique_pickups,
                                "unique_deliveries": unique_deliveries,
                                "num_machines": len(edited_machines),
                                "other_notes": other_notes,
                                **quote
                            }
                    else:
                        st.error("Need at least one pickup or delivery address")
                
                # Show results
                if 'quote' in st.session_state:
                    quote = st.session_state['quote']
                    route = st.session_state['route_data']
                    
                    st.success(f"### 💵 Quote: ${quote['final_price']:,}")
                    
                    # Route details
                    st.markdown("**Route Legs:**")
                    for leg in route['legs']:
                        est = " ⚠️" if leg.get('estimated') else ""
                        st.caption(f"• {leg['from'][:30]}... → {leg['to'][:30]}...: {leg['distance_miles']} mi, {leg['duration_minutes']} min{est}")
                    
                    # Indicators
                    col_a, col_b = st.columns(2)
                    with col_a:
                        if quote['is_tucson']:
                            st.warning("🌵 Tucson - $850 min")
                        if quote['is_prison']:
                            st.warning("🏛️ Prison - $900 min")
                    with col_b:
                        if quote['buffer_time'] > 0:
                            st.info(f"📍 Buffer: +{quote['buffer_time']} min (>{BUFFER_THRESHOLD_MILES} mi)")
                        else:
                            st.info(f"💰 No-buffer discount: -$60")
                    
                    # Breakdown
                    st.markdown(f"""
                    | Component | Value |
                    |-----------|-------|
                    | Drive Time | {quote['drive_time']} min |
                    | Job Time | {quote['job_time']} min |
                    | Buffer | {quote['buffer_time']} min |
                    | Max Leg Distance | {quote['max_distance_miles']} mi |
                    | **Total** | **{quote['total_hours']} hrs** |
                    """)
                    
                    st.caption(f"Formula: {quote['formula']}")
                    
                    st.divider()
                    
                    # Downloads
                    full_data = st.session_state.get('full_data', {})
                    
                    col_dl1, col_dl2 = st.columns(2)
                    
                    with col_dl1:
                        if st.button("📋 Convert Excel to PDF", use_container_width=True):
                            with st.spinner("Converting..."):
                                pdf = convert_excel_to_pdf(st.session_state.get('excel_bytes'), st.session_state.get('excel_name', 'request.xlsx'))
                                if pdf:
                                    pdf = remove_pdf_pages(pdf, [1])
                                    st.session_state['request_pdf'] = pdf
                                    st.success("✅ PDF ready!")
                    
                    with col_dl2:
                        if 'request_pdf' in st.session_state:
                            st.download_button(
                                "⬇️ Download CAPTURE PDF",
                                data=st.session_state['request_pdf'],
                                file_name=f"CAPTURE_{mr_number or 'BEP'}_{datetime.now().strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                    
                    # Trello
                    if trello_key and trello_token:
                        if st.button("📋 Create Trello Card + Attach Files", use_container_width=True):
                            with st.spinner("Creating card..."):
                                card = create_trello_card(full_data, trello_key, trello_token, trello_list)
                                if card:
                                    card_id = card.get('id')
                                    st.success(f"✅ Card created: {card.get('shortUrl')}")
                                    
                                    # Attach Excel file (needed for Quote generation later)
                                    with st.spinner("Attaching Excel..."):
                                        if attach_excel_to_card(card_id, st.session_state.get('excel_bytes'), st.session_state.get('excel_name', 'request.xlsx'), trello_key, trello_token):
                                            st.success("✅ Excel attached!")
                                        else:
                                            st.warning("⚠️ Excel attachment failed")
                                    
                                    # Attach PDF if available
                                    if 'request_pdf' in st.session_state:
                                        with st.spinner("Attaching CAPTURE PDF..."):
                                            pdf_filename = f"CAPTURE_{mr_number or 'BEP'}_{datetime.now().strftime('%Y%m%d')}.pdf"
                                            if attach_pdf_to_card(card_id, st.session_state['request_pdf'], pdf_filename, trello_key, trello_token):
                                                st.success("✅ CAPTURE PDF attached!")
                                            else:
                                                st.warning("⚠️ PDF attachment failed")
                                    else:
                                        st.info("💡 Convert Excel to PDF first for CAPTURE attachment")
                                    
                                    st.markdown(f"**Next:** After Ryan approves, go to **📝 Generate Quote** page to create QUOTE PDF")
                                else:
                                    st.error("Failed to create Trello card")
            
            # Raw data viewer
            with st.expander("🔍 Raw Excel Data"):
                st.text("\n".join(data.get("raw_data", [])[:60]))
        
        else:
            st.error(f"❌ Error: {data.get('error')}")
    
    else:
        st.info("👆 Upload a BEP Move Request Excel file to get started")

# Footer
st.divider()
st.caption("BEP Pricing Calculator v4.0 | Two-Step Workflow | Tool Box & Safe Moving")
