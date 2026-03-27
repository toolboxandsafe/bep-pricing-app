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
GOOGLE_MAPS_API_KEY = "AIzaSyDnCiWB4EiXP8nSYVaveMJv367PsmxCFDw"

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
            
            # Find where "Other Notes" section starts - this is END of machine data
            if "OTHER NOTE" in row_text or "SPECIAL INSTRUCTION" in row_text or "ADDITIONAL NOTE" in row_text:
                other_notes_row = row_idx
                # Capture the notes content (might be on same row or next rows)
                for check_row in rows[row_idx:row_idx+5]:
                    for cell in check_row:
                        if cell and len(cell) > 20 and 'NAN' not in cell.upper() and 'OTHER NOTE' not in cell.upper():
                            result["other_notes"] = cell
                            break
                break
            
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
        
        # Find requester (row 47, index 46)
        if len(rows) > 46:
            for cell in rows[46][:4]:
                if cell and not any(x in cell.upper() for x in ['REQUESTER', 'NAME', 'DATE', 'SIGNATURE', 'NAN']):
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
            
            # Section detection
            if "PICK UP" in row_text or "PICKUP" in row_text:
                current_section = "pickup"
            elif "DELIVER" in row_text or "DELIVERY" in row_text:
                current_section = "delivery"
            
            # Extract addresses (only if we have a current machine and valid section)
            if current_machine and current_section in ['pickup', 'delivery']:
                for cell in row:
                    if not cell or len(cell) < 5:
                        continue
                    cell_upper = cell.upper()
                    
                    # Skip labels
                    if cell_upper in ['NAN', 'PICK UP', 'PICKUP', 'DELIVERY', 'DELIVER TO', 'SITE', 'LOCATION']:
                        continue
                    
                    # Check for address indicators
                    is_address = any(ind in cell_upper for ind in [
                        'AVE', 'STREET', 'ST ', 'BLVD', 'ROAD', 'RD ', 'DRIVE', 'DR ',
                        'LANE', 'WAY', 'PHOENIX', 'PHX', 'TUCSON', 'MESA', 'TEMPE',
                        'GILBERT', 'SCOTTSDALE', 'CHANDLER', 'GLENDALE', 'PEORIA',
                        'MAXIMUS', 'DES ', 'ADES', 'BEP ', 'DCS', 'CIVIC'
                    ]) or re.search(r'\d{3,5}\s+[A-Z]', cell_upper)
                    
                    if is_address:
                        if current_section == "pickup" and not current_machine["pickup"]:
                            current_machine["pickup"] = cell
                            current_machine["pickup_address"] = normalize_address(cell)
                        elif current_section == "delivery" and not current_machine["delivery"]:
                            current_machine["delivery"] = cell
                            current_machine["delivery_address"] = normalize_address(cell)
            
            # Extract machine type
            if current_machine:
                for cell in row:
                    cell_upper = cell.upper()
                    if any(kw in cell_upper for kw in ['VENDING', 'COMBO', 'SNACK', 'SODA', 'CHANGER', 'KIOSK', 'FROZEN', 'ATM']):
                        if not current_machine["type"]:
                            current_machine["type"] = cell
        
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
    
    # Use 11am tomorrow for consistent traffic
    tomorrow_11am = datetime.now().replace(hour=11, minute=0, second=0) + timedelta(days=1)
    departure_timestamp = int(tomorrow_11am.timestamp())
    
    params = {
        "origins": "|".join(origins),
        "destinations": "|".join(destinations),
        "key": GOOGLE_MAPS_API_KEY,
        "departure_time": departure_timestamp,
        "traffic_model": "optimistic",
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
            
            # Duration in minutes (use traffic duration if available)
            if "duration_in_traffic" in element:
                duration_seconds = element["duration_in_traffic"]["value"]
            else:
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
    return response.json() if response.status_code == 200 else None

# =============================================================================
# STREAMLIT UI
# =============================================================================

st.title("🚚 BEP Pricing Calculator v3")
st.markdown("**Auto-extract machines → Google Maps routing → Quote calculation**")

# Sidebar
with st.sidebar:
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
    
    with st.expander("🔧 Trello Settings"):
        trello_key = st.text_input("API Key", type="password")
        trello_token = st.text_input("API Token", type="password")
        trello_list = st.text_input("List ID", value="699c9f9d6117bdcbb2d0e0aa")

# Main content
st.header("📤 Upload BEP Move Request")

uploaded_file = st.file_uploader(
    "Drop your BEP Excel file here",
    type=["xlsx", "xls"],
    help="Upload the Move Request Excel file"
)

if uploaded_file:
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
                            pdf = convert_excel_to_pdf(uploaded_file.getvalue(), uploaded_file.name)
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
                    if st.button("📋 Create Trello Card", use_container_width=True):
                        card = create_trello_card(full_data, trello_key, trello_token, trello_list)
                        if card:
                            st.success(f"✅ Card created: {card.get('shortUrl')}")
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
st.caption("BEP Pricing Calculator v3.0 | Google Maps Routing | Tool Box & Safe Moving")
