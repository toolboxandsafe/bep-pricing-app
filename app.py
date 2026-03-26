"""
BEP Pricing Calculator - Streamlit Web App
Tool Box & Safe Moving - Vending Machine Move Pricing
Replicates Email-to-Trello Workflow
"""

import streamlit as st
import json
import math
import re
from datetime import datetime
import pandas as pd
from io import BytesIO
from fpdf import FPDF
import requests

# Page config
st.set_page_config(
    page_title="BEP Pricing Calculator",
    page_icon="🚚",
    layout="wide"
)

# Load pricing rules
@st.cache_data
def load_pricing_rules():
    with open("data/pricing_rules.json", "r") as f:
        return json.load(f)

# Load known locations
@st.cache_data  
def load_locations():
    try:
        with open("data/known_locations.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

rules = load_pricing_rules()
locations = load_locations()

# Constants
HQ_ADDRESS = "Gilbert, AZ 85295"
HOURLY_RATE = 170
BUFFER_THRESHOLD_MILES = 35
BUFFER_MINUTES = 20
JOB_TIME_PER_MACHINE = 30  # minutes

# =============================================================================
# ADDRESS DEDUPLICATION
# =============================================================================

def normalize_address(address):
    """Normalize address for deduplication - extract just the street address"""
    if not address:
        return ""
    
    addr = str(address).upper().strip()
    
    # Remove common prefixes like facility names
    prefixes_to_remove = ['BEP ', 'MAXIMUS ', 'DES ', 'ADES ']
    for prefix in prefixes_to_remove:
        if addr.startswith(prefix):
            addr = addr[len(prefix):]
    
    # Remove extra whitespace
    addr = ' '.join(addr.split())
    
    # Extract just street address portion (number + street name)
    # Look for pattern: number + street name
    import re
    street_match = re.search(r'(\d+\s+[A-Z0-9\s\.]+(?:STREET|ST|AVENUE|AVE|BLVD|ROAD|RD|DRIVE|DR|LANE|LN|WAY|PKWY|HWY)[\.]*)', addr)
    if street_match:
        addr = street_match.group(1)
    
    # Standardize abbreviations
    addr = addr.replace(' STREET', ' ST').replace(' AVENUE', ' AVE')
    addr = addr.replace(' BOULEVARD', ' BLVD').replace(' ROAD', ' RD')
    addr = addr.replace(' DRIVE', ' DR').replace(' LANE', ' LN')
    addr = addr.replace('.', '').replace(',', '')
    
    # Remove suite/unit numbers for deduplication
    addr = re.sub(r'\s*#\s*\d+', '', addr)
    addr = re.sub(r'\s*SUITE\s*\d+', '', addr)
    addr = re.sub(r'\s*UNIT\s*\d+', '', addr)
    
    return addr.strip()

def deduplicate_locations(locations):
    """Deduplicate locations based on normalized address"""
    seen_addresses = {}
    unique_locations = []
    
    for loc in locations:
        if not loc:
            continue
        
        normalized = normalize_address(loc)
        
        if normalized and normalized not in seen_addresses:
            seen_addresses[normalized] = loc
            unique_locations.append(loc)
        elif not normalized:
            # Keep locations without clear addresses (might be facility names)
            unique_locations.append(loc)
    
    return unique_locations

# =============================================================================
# EXCEL PARSING - BEP REQUEST TAB
# =============================================================================

def parse_bep_excel(uploaded_file):
    """Parse BEP Move Request Excel file - extracts from REQUEST tab"""
    try:
        # Try to read REQUEST tab specifically
        try:
            df = pd.read_excel(uploaded_file, sheet_name='REQUEST', header=None)
        except:
            # Fallback to first sheet
            df = pd.read_excel(uploaded_file, header=None)
        
        extracted = {
            "success": True,
            "requester": None,
            "mr_number": None,
            "pickups": [],
            "deliveries": [],
            "num_machines": 0,
            "machine_types": [],
            "move_date": None,
            "other_notes": None,
            "contacts": [],
            "raw_text": ""
        }
        
        # Convert entire sheet to searchable text
        all_text = []
        current_section = None
        
        for row_idx, row in df.iterrows():
            row_text = " | ".join([str(c) for c in row if pd.notna(c)])
            all_text.append(row_text)
            
            for col_idx, cell in enumerate(row):
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    cell_upper = cell_str.upper()
                    
                    # Get adjacent cell values
                    next_col = str(row.iloc[col_idx + 1]).strip() if col_idx + 1 < len(row) and pd.notna(row.iloc[col_idx + 1]) else None
                    prev_row = str(df.iloc[row_idx - 1, col_idx]).strip() if row_idx > 0 and col_idx < len(df.iloc[row_idx - 1]) and pd.notna(df.iloc[row_idx - 1, col_idx]) else None
                    
                    # MR Number (format: 1XX-XX/XX)
                    if re.match(r'1\w{2}-\d{2}/\d{2}', cell_str):
                        extracted["mr_number"] = cell_str
                    
                    # Requester Name - check row 47 (index 46)
                    if row_idx == 46 and col_idx < 3:
                        if cell_str and not any(x in cell_upper for x in ['REQUESTER', 'NAME', 'DATE', 'SIGNATURE']):
                            extracted["requester"] = cell_str
                    
                    # Section detection
                    if "PICK UP" in cell_upper or "PICKUP" in cell_upper:
                        current_section = "pickup"
                    elif "DELIVER" in cell_upper or "DELIVERY" in cell_upper:
                        current_section = "delivery"
                    elif "OTHER NOTE" in cell_upper or "SPECIAL" in cell_upper:
                        current_section = "notes"
                    
                    # Address extraction (look for address patterns)
                    address_indicators = ['AVE', 'STREET', 'ST.', 'ST ', 'BLVD', 'ROAD', 'RD', 'DRIVE', 'DR.', 'LANE', 'WAY', 'PHOENIX', 'PHX', 'TUCSON', 'MESA', 'TEMPE', 'GILBERT', 'SCOTTSDALE', 'CHANDLER']
                    if any(ind in cell_upper for ind in address_indicators):
                        # This looks like an address
                        if current_section == "pickup" and cell_str not in extracted["pickups"]:
                            extracted["pickups"].append(cell_str)
                        elif current_section == "delivery" and cell_str not in extracted["deliveries"]:
                            extracted["deliveries"].append(cell_str)
                    
                    # Site/Location names (often have facility info)
                    if "SITE" in cell_upper or "LOCATION" in cell_upper or "FACILITY" in cell_upper:
                        if next_col:
                            if current_section == "pickup" and next_col not in extracted["pickups"]:
                                extracted["pickups"].append(next_col)
                            elif current_section == "delivery" and next_col not in extracted["deliveries"]:
                                extracted["deliveries"].append(next_col)
                    
                    # Machine type detection
                    machine_keywords = ['VENDING', 'MACHINE', 'KIOSK', 'ATM', 'SNACK', 'SODA', 'COMBO', 'CHANGER', 'FROZEN']
                    if any(kw in cell_upper for kw in machine_keywords):
                        if cell_str not in extracted["machine_types"]:
                            extracted["machine_types"].append(cell_str)
                    
                    # Other Notes
                    if current_section == "notes" and cell_str and len(cell_str) > 10:
                        extracted["other_notes"] = cell_str
                    
                    # Contact info
                    if "CONTACT" in cell_upper or "PHONE" in cell_upper:
                        if next_col:
                            extracted["contacts"].append(next_col)
                    
                    # Date
                    if "DATE" in cell_upper and next_col:
                        try:
                            extracted["move_date"] = next_col
                        except:
                            pass
        
        # DEDUPLICATE locations based on normalized addresses
        extracted["pickups_raw"] = extracted["pickups"].copy()
        extracted["deliveries_raw"] = extracted["deliveries"].copy()
        
        extracted["pickups"] = deduplicate_locations(extracted["pickups"])
        extracted["deliveries"] = deduplicate_locations(extracted["deliveries"])
        
        # Count machines based on pickup/delivery pairs (use raw count before dedup)
        extracted["num_machines"] = max(
            len(extracted["pickups_raw"]), 
            len(extracted["deliveries_raw"]), 
            1
        )
        
        # Count unique stops (deduplicated)
        extracted["unique_stops"] = len(set(
            [normalize_address(p) for p in extracted["pickups"]] + 
            [normalize_address(d) for d in extracted["deliveries"]]
        ))
        
        # Store raw text for display
        extracted["raw_text"] = "\n".join(all_text)
        
        # Try alternate requester extraction if not found
        if not extracted["requester"]:
            # Look for name patterns in first few rows
            for row_idx in range(min(5, len(df))):
                for col_idx in range(min(5, len(df.columns))):
                    if pd.notna(df.iloc[row_idx, col_idx]):
                        val = str(df.iloc[row_idx, col_idx]).strip()
                        # Look for name-like patterns (First Last)
                        if re.match(r'^[A-Z][a-z]+ [A-Z][a-z]+$', val):
                            extracted["requester"] = val
                            break
        
        return extracted
        
    except Exception as e:
        return {"success": False, "error": str(e)}

# =============================================================================
# PRICING FUNCTIONS
# =============================================================================

def calculate_quote(pickups, deliveries, num_machines, drive_time_minutes, max_distance_miles=None):
    """Calculate BEP quote following the workflow rules"""
    
    # Drive time (round trip calculation already included if user provides one-way)
    # The workflow calculates: HQ → Pickups → Deliveries → HQ
    # User provides total drive time for the route
    
    total_drive_time = drive_time_minutes
    
    # Job time: 30 minutes per machine
    job_time = JOB_TIME_PER_MACHINE * num_machines
    
    # Buffer: 20 minutes if max distance > 35 miles
    buffer_time = 0
    if max_distance_miles and max_distance_miles > BUFFER_THRESHOLD_MILES:
        buffer_time = BUFFER_MINUTES
    
    # Total time
    total_minutes = total_drive_time + job_time + buffer_time
    total_hours = total_minutes / 60
    
    # Base price
    base_price = total_hours * HOURLY_RATE
    
    # Apply minimums
    min_price = 220  # General minimum
    
    # Check for Tucson locations
    all_locations = pickups + deliveries
    is_tucson = any('TUCSON' in loc.upper() for loc in all_locations if loc)
    if is_tucson:
        min_price = 850
    
    # Check for prison
    is_prison = any(kw in str(all_locations).upper() for kw in ['ASPC', 'PRISON', 'CORRECTIONAL', 'CIMARRON'])
    if is_prison:
        min_price = max(min_price, 900)
    
    # Apply minimum
    final_price = max(base_price, min_price)
    
    # Round UP to nearest $25
    final_price = math.ceil(final_price / 25) * 25
    
    return {
        "drive_time": total_drive_time,
        "job_time": job_time,
        "buffer_time": buffer_time,
        "total_minutes": total_minutes,
        "total_hours": round(total_hours, 2),
        "base_price": round(base_price, 2),
        "min_price": min_price,
        "final_price": int(final_price),
        "is_tucson": is_tucson,
        "is_prison": is_prison,
        "formula": f"({total_drive_time} + {job_time} + {buffer_time}) ÷ 60 × ${HOURLY_RATE} = ${base_price:.2f}"
    }

# =============================================================================
# PDF GENERATION
# =============================================================================

def sanitize_for_pdf(text):
    """Remove or replace characters that cause PDF encoding issues"""
    if not text:
        return ""
    
    text = str(text)
    
    # Replace common problematic characters
    replacements = {
        '–': '-',  # en dash
        '—': '-',  # em dash
        ''': "'",  # smart quote
        ''': "'",  # smart quote
        '"': '"',  # smart quote
        '"': '"',  # smart quote
        '…': '...',  # ellipsis
        '•': '*',  # bullet
        '\u2022': '*',  # bullet
        '\u2019': "'",  # right single quote
        '\u201c': '"',  # left double quote
        '\u201d': '"',  # right double quote
        '\xa0': ' ',  # non-breaking space
    }
    
    for old, new in replacements.items():
        text = text.replace(old, new)
    
    # Remove any remaining non-ASCII characters
    text = text.encode('ascii', 'ignore').decode('ascii')
    
    return text

def generate_quote_pdf(data):
    """Generate professional quote PDF"""
    pdf = FPDF()
    pdf.add_page()
    
    # Header
    pdf.set_font('Helvetica', 'B', 24)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(0, 15, 'Tool Box & Safe Moving', ln=True, align='C')
    
    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, 'BEP Vending Machine Move Quote', ln=True, align='C')
    pdf.cell(0, 6, 'Phone: (602) 935-4209', ln=True, align='C')
    
    pdf.ln(10)
    
    # Quote amount
    pdf.set_fill_color(34, 139, 34)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('Helvetica', 'B', 20)
    pdf.cell(0, 15, f"  QUOTE: ${data.get('final_price', 0):,}", ln=True, fill=True)
    
    pdf.ln(5)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 6, f"Date: {datetime.now().strftime('%B %d, %Y')}", ln=True, align='R')
    
    pdf.ln(5)
    
    # Move Details
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 8, '  MOVE DETAILS', ln=True, fill=True)
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Helvetica', '', 10)
    
    details = [
        ('MR Number', sanitize_for_pdf(data.get('mr_number', 'N/A'))),
        ('Requester', sanitize_for_pdf(data.get('requester', 'N/A'))),
        ('Move Date', sanitize_for_pdf(data.get('move_date', 'TBD'))),
        ('Machines', str(data.get('num_machines', 1))),
    ]
    
    for label, value in details:
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(50, 7, f"{label}:")
        pdf.set_font('Helvetica', '', 10)
        pdf.cell(0, 7, sanitize_for_pdf(str(value)) if value else 'N/A', ln=True)
    
    # Pickups
    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(0, 7, 'PICKUP LOCATIONS:', ln=True)
    pdf.set_font('Helvetica', '', 9)
    for pickup in data.get('pickups', []):
        pdf.cell(0, 6, f"  - {sanitize_for_pdf(pickup)}", ln=True)
    
    # Deliveries
    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(0, 7, 'DELIVERY LOCATIONS:', ln=True)
    pdf.set_font('Helvetica', '', 9)
    for delivery in data.get('deliveries', []):
        pdf.cell(0, 6, f"  - {sanitize_for_pdf(delivery)}", ln=True)
    
    pdf.ln(5)
    
    # Pricing breakdown
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 8, '  PRICING BREAKDOWN', ln=True, fill=True)
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Helvetica', '', 10)
    
    pricing = [
        ('Drive Time', f"{data.get('drive_time', 0)} min"),
        ('Job Time', f"{data.get('job_time', 0)} min ({data.get('num_machines', 1)} machines × 30 min)"),
        ('Buffer', f"{data.get('buffer_time', 0)} min"),
        ('Total Time', f"{data.get('total_hours', 0)} hours"),
        ('Rate', f"${HOURLY_RATE}/hour"),
        ('Minimum Applied', f"${data.get('min_price', 220)}"),
    ]
    
    for label, value in pricing:
        pdf.set_font('Helvetica', '', 10)
        pdf.cell(50, 6, f"{label}:")
        pdf.cell(0, 6, str(value), ln=True)
    
    pdf.ln(3)
    pdf.set_font('Helvetica', 'I', 9)
    pdf.cell(0, 5, f"Formula: {sanitize_for_pdf(data.get('formula', ''))}", ln=True)
    
    # Other Notes
    if data.get('other_notes'):
        pdf.ln(5)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(0, 7, 'NOTES:', ln=True)
        pdf.set_font('Helvetica', '', 9)
        pdf.multi_cell(0, 5, sanitize_for_pdf(data.get('other_notes', '')))
    
    # Footer
    pdf.ln(10)
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(128, 128, 128)
    pdf.cell(0, 5, 'Quote valid for 30 days. Additional charges may apply for stairs.', ln=True, align='C')
    
    # Return as bytes for Streamlit download
    pdf_output = pdf.output()
    if isinstance(pdf_output, bytearray):
        return bytes(pdf_output)
    elif isinstance(pdf_output, bytes):
        return pdf_output
    else:
        # Fallback: write to BytesIO
        buffer = BytesIO()
        pdf.output(buffer)
        return buffer.getvalue()

# =============================================================================
# TRELLO INTEGRATION
# =============================================================================

def create_trello_card(data, api_key, api_token, list_id):
    """Create Trello card with quote data"""
    
    # Build title: Requester - MR# - $Quote
    title = f"{data.get('requester', 'Unknown')} - {data.get('mr_number', 'BEP')} - ${data.get('final_price', 0)}"
    
    # Build description
    desc = f"""## Move Request Quote

**MR Number:** {data.get('mr_number', 'N/A')}
**Requester:** {data.get('requester', 'N/A')}
**Move Date:** {data.get('move_date', 'TBD')}
**Machines:** {data.get('num_machines', 1)}

---

### 📍 PICKUP LOCATIONS
{chr(10).join(['• ' + p for p in data.get('pickups', [])])}

### 📍 DELIVERY LOCATIONS
{chr(10).join(['• ' + d for d in data.get('deliveries', [])])}

---

### 💰 QUOTE: ${data.get('final_price', 0):,}

**Breakdown:**
- Drive Time: {data.get('drive_time', 0)} min
- Job Time: {data.get('job_time', 0)} min
- Buffer: {data.get('buffer_time', 0)} min
- Total Hours: {data.get('total_hours', 0)}
- Rate: ${HOURLY_RATE}/hour

**Formula:** {data.get('formula', '')}

---

### 📝 OTHER NOTES
{data.get('other_notes', 'None')}

---
@luissaravia2
"""
    
    url = f"https://api.trello.com/1/cards"
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

st.title("🚚 BEP Pricing Calculator")
st.markdown("**Tool Box & Safe Moving** - Upload Excel → Auto-Calculate → Generate Quote")

# Sidebar
with st.sidebar:
    st.header("📋 Pricing Rules")
    st.markdown(f"""
    **Rate:** ${HOURLY_RATE}/hour
    **Job Time:** 30 min/machine
    **Buffer:** +20 min if >35 miles
    
    **Minimums:**
    - General: $220
    - Tucson: $850
    - Prison: $900+
    
    **Rounding:** Up to $25
    """)
    
    st.divider()
    
    # Trello settings (collapsible)
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
        data = parse_bep_excel(uploaded_file)
    
    if data["success"]:
        st.success("✅ File parsed! Review and edit the extracted data below.")
        
        # Two columns
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📋 Move Details")
            
            requester = st.text_input("Requester Name", value=data.get("requester") or "")
            mr_number = st.text_input("MR Number", value=data.get("mr_number") or "")
            move_date = st.text_input("Move Date", value=data.get("move_date") or "")
            
            st.markdown("**Pickup Locations:**")
            pickups_text = st.text_area(
                "Pickups (one per line)",
                value="\n".join(data.get("pickups", [])),
                height=100
            )
            pickups = [p.strip() for p in pickups_text.split("\n") if p.strip()]
            
            st.markdown("**Delivery Locations:**")
            deliveries_text = st.text_area(
                "Deliveries (one per line)",
                value="\n".join(data.get("deliveries", [])),
                height=100
            )
            deliveries = [d.strip() for d in deliveries_text.split("\n") if d.strip()]
            
            num_machines = st.number_input(
                "Number of Machines",
                min_value=1,
                max_value=20,
                value=max(data.get("num_machines", 1), len(pickups), len(deliveries))
            )
            
            other_notes = st.text_area(
                "Other Notes",
                value=data.get("other_notes") or "",
                height=80
            )
            
            # Show deduplication info
            if data.get("pickups_raw") or data.get("deliveries_raw"):
                raw_count = len(data.get("pickups_raw", [])) + len(data.get("deliveries_raw", []))
                dedup_count = len(pickups) + len(deliveries)
                if raw_count > dedup_count:
                    st.info(f"📍 **Deduplicated:** {raw_count} locations → {dedup_count} unique stops")
        
        with col2:
            st.subheader("🕐 Drive Time & Quote")
            
            st.info("""
            **Route:** Gilbert, AZ → Pickups → Deliveries → Gilbert, AZ
            
            Enter the TOTAL drive time for this route (use Google Maps).
            """)
            
            drive_time = st.number_input(
                "Total Route Drive Time (minutes)",
                min_value=10,
                max_value=600,
                value=120,
                help="Total driving time for the entire route"
            )
            
            max_distance = st.number_input(
                "Max Single Leg Distance (miles)",
                min_value=0,
                max_value=300,
                value=40,
                help="Longest single leg of the trip (for buffer calculation)"
            )
            
            st.divider()
            
            # Calculate quote
            if st.button("🧮 CALCULATE QUOTE", type="primary", use_container_width=True):
                result = calculate_quote(pickups, deliveries, num_machines, drive_time, max_distance)
                
                # Store in session state
                st.session_state['quote_result'] = result
                st.session_state['quote_data'] = {
                    "requester": requester,
                    "mr_number": mr_number,
                    "move_date": move_date,
                    "pickups": pickups,
                    "deliveries": deliveries,
                    "num_machines": num_machines,
                    "other_notes": other_notes,
                    **result
                }
            
            # Show results if calculated
            if 'quote_result' in st.session_state:
                result = st.session_state['quote_result']
                
                st.success(f"### 💵 Quote: ${result['final_price']:,}")
                
                # Indicators
                col_a, col_b = st.columns(2)
                with col_a:
                    if result['is_tucson']:
                        st.warning("🌵 Tucson job - $850 min")
                    if result['is_prison']:
                        st.warning("🏛️ Prison job - $900 min")
                with col_b:
                    if result['buffer_time'] > 0:
                        st.info(f"📍 Buffer: +{result['buffer_time']} min")
                
                # Breakdown
                st.markdown(f"""
                | Component | Value |
                |-----------|-------|
                | Drive Time | {result['drive_time']} min |
                | Job Time | {result['job_time']} min |
                | Buffer | {result['buffer_time']} min |
                | **Total** | **{result['total_hours']} hrs** |
                | Rate | ${HOURLY_RATE}/hr |
                | Minimum | ${result['min_price']} |
                | **QUOTE** | **${result['final_price']}** |
                """)
                
                st.caption(f"Formula: {result['formula']}")
                
                st.divider()
                
                # Action buttons
                quote_data = st.session_state.get('quote_data', {})
                
                # PDF Download
                pdf_bytes = generate_quote_pdf(quote_data)
                st.download_button(
                    "📄 Download Quote PDF",
                    data=pdf_bytes,
                    file_name=f"QUOTE_{mr_number or 'BEP'}_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # Trello button
                if trello_key and trello_token:
                    if st.button("📋 Create Trello Card", use_container_width=True):
                        card = create_trello_card(quote_data, trello_key, trello_token, trello_list)
                        if card:
                            st.success(f"✅ Trello card created: {card.get('shortUrl')}")
                        else:
                            st.error("Failed to create Trello card")
                else:
                    st.info("💡 Add Trello credentials in sidebar to create cards")
        
        # Raw data viewer
        with st.expander("🔍 View Raw Excel Data"):
            st.dataframe(pd.read_excel(uploaded_file, header=None).head(60))
    
    else:
        st.error(f"❌ Error: {data.get('error')}")

else:
    st.info("👆 Upload a BEP Move Request Excel file to get started")
    
    # Show sample workflow
    with st.expander("📖 How it works"):
        st.markdown("""
        1. **Upload** the BEP Excel file (Move Request worksheet)
        2. **Review** the auto-extracted data (requester, addresses, machines)
        3. **Enter** the drive time from Google Maps
        4. **Calculate** the quote automatically
        5. **Download** the PDF quote
        6. **Create** Trello card (optional)
        
        **Pricing Formula:**
        ```
        (Drive Time + Job Time + Buffer) ÷ 60 × $170/hour
        ```
        
        **Job Time:** 30 minutes per machine
        **Buffer:** 20 minutes if any leg > 35 miles
        **Minimums:** $220 general, $850 Tucson, $900 prison
        """)

# Footer
st.divider()
st.caption("BEP Pricing Calculator v2.0 | Tool Box & Safe Moving | Built by Grant")
