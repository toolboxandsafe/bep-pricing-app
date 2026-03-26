"""
BEP Pricing Calculator - Streamlit Web App
Tool Box & Safe Moving - Vending Machine Move Pricing
"""

import streamlit as st
import json
import math
from datetime import datetime
import pandas as pd
from io import BytesIO

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

# =============================================================================
# EXCEL PARSING FUNCTIONS
# =============================================================================

def parse_bep_excel(uploaded_file):
    """Parse BEP Move Request Excel file and extract all relevant data"""
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, header=None)
        
        extracted_data = {
            "success": True,
            "requester": None,
            "mr_number": None,
            "pickup_address": None,
            "delivery_address": None,
            "num_machines": 1,
            "machine_type": None,
            "move_date": None,
            "contact_name": None,
            "contact_phone": None,
            "notes": None,
            "raw_data": {}
        }
        
        # Search for key fields in the Excel
        for row_idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                if pd.notna(cell):
                    cell_str = str(cell).strip().upper()
                    cell_value = str(cell).strip()
                    
                    # Get the value to the right or below
                    next_col_value = None
                    next_row_value = None
                    
                    if col_idx + 1 < len(row) and pd.notna(row.iloc[col_idx + 1]):
                        next_col_value = str(row.iloc[col_idx + 1]).strip()
                    if row_idx + 1 < len(df) and col_idx < len(df.iloc[row_idx + 1]) and pd.notna(df.iloc[row_idx + 1, col_idx]):
                        next_row_value = str(df.iloc[row_idx + 1, col_idx]).strip()
                    
                    # Requester Name - typically Row 47 or labeled
                    if "REQUESTER" in cell_str and "NAME" in cell_str:
                        # Value is above this cell (Row 47)
                        if row_idx > 0 and pd.notna(df.iloc[row_idx - 1, col_idx]):
                            extracted_data["requester"] = str(df.iloc[row_idx - 1, col_idx]).strip()
                    
                    # MR Number
                    if "MR" in cell_str and any(c.isdigit() for c in cell_value):
                        extracted_data["mr_number"] = cell_value
                    if cell_str.startswith("1") and "-" in cell_value and "/" in cell_value:
                        extracted_data["mr_number"] = cell_value
                    
                    # Pickup/From Address
                    if any(keyword in cell_str for keyword in ["PICKUP", "FROM", "ORIGIN", "CURRENT LOCATION"]):
                        if next_col_value:
                            extracted_data["pickup_address"] = next_col_value
                        elif next_row_value:
                            extracted_data["pickup_address"] = next_row_value
                    
                    # Delivery/To Address
                    if any(keyword in cell_str for keyword in ["DELIVER", "TO ADDRESS", "DESTINATION", "NEW LOCATION"]):
                        if next_col_value:
                            extracted_data["delivery_address"] = next_col_value
                        elif next_row_value:
                            extracted_data["delivery_address"] = next_row_value
                    
                    # Number of machines
                    if any(keyword in cell_str for keyword in ["QTY", "QUANTITY", "# OF", "NUMBER OF", "MACHINES"]):
                        if next_col_value and next_col_value.isdigit():
                            extracted_data["num_machines"] = int(next_col_value)
                        elif next_row_value and next_row_value.isdigit():
                            extracted_data["num_machines"] = int(next_row_value)
                    
                    # Machine type
                    if any(keyword in cell_str for keyword in ["MACHINE TYPE", "EQUIPMENT", "TYPE"]):
                        if next_col_value:
                            extracted_data["machine_type"] = next_col_value
                    
                    # Move date
                    if any(keyword in cell_str for keyword in ["DATE", "MOVE DATE", "SCHEDULED"]):
                        if next_col_value:
                            extracted_data["move_date"] = next_col_value
                    
                    # Contact info
                    if "CONTACT" in cell_str and "NAME" in cell_str:
                        if next_col_value:
                            extracted_data["contact_name"] = next_col_value
                    if "PHONE" in cell_str or "TEL" in cell_str:
                        if next_col_value:
                            extracted_data["contact_phone"] = next_col_value
                    
                    # Notes/Comments
                    if any(keyword in cell_str for keyword in ["NOTE", "COMMENT", "SPECIAL", "INSTRUCTION"]):
                        if next_col_value:
                            extracted_data["notes"] = next_col_value
                    
                    # Look for addresses in cells (contain street indicators)
                    if any(indicator in cell_str for indicator in ["AVE", "STREET", "ST.", "BLVD", "ROAD", "RD.", "DRIVE", "DR.", "LANE", "LN."]):
                        # This might be an address
                        if not extracted_data["pickup_address"]:
                            extracted_data["raw_data"]["possible_address_1"] = cell_value
                        elif not extracted_data["delivery_address"]:
                            extracted_data["raw_data"]["possible_address_2"] = cell_value
        
        # Try to extract requester from Row 47 specifically (0-indexed = row 46)
        if not extracted_data["requester"] and len(df) > 46:
            for col_idx in range(min(5, len(df.columns))):
                if pd.notna(df.iloc[46, col_idx]):
                    val = str(df.iloc[46, col_idx]).strip()
                    if val and not val.upper().startswith(("REQUESTER", "NAME", "DATE")):
                        extracted_data["requester"] = val
                        break
        
        # Store raw preview
        extracted_data["raw_data"]["preview"] = df.head(20).to_dict()
        
        return extracted_data
        
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

# =============================================================================
# PRICING FUNCTIONS
# =============================================================================

def round_up_to_nearest(value, nearest):
    """Round up to nearest value (e.g., nearest $25)"""
    return math.ceil(value / nearest) * nearest

def detect_location_type(address):
    """Detect location type from address keywords"""
    if not address:
        return "standard", None
    
    address_upper = address.upper()
    
    for loc_type, config in rules["location_types"].items():
        keywords = config.get("keywords", [])
        for keyword in keywords:
            if keyword.upper() in address_upper:
                return loc_type, config
    
    return "standard", None

def calculate_price(pickup_address, delivery_address, num_machines, drive_time_minutes, is_internal_move=False):
    """Calculate BEP move price based on rules"""
    
    hourly_rate = rules["base_rate"]["hourly_rate"]
    
    # Detect location types
    pickup_type, pickup_config = detect_location_type(pickup_address)
    delivery_type, delivery_config = detect_location_type(delivery_address)
    
    # Determine primary location type
    primary_type = delivery_type if delivery_type != "standard" else pickup_type
    primary_config = delivery_config if delivery_config else pickup_config
    
    # Base calculation: drive time * 2 (round trip) * hourly rate
    base_hours = (drive_time_minutes * 2) / 60
    base_price = base_hours * hourly_rate
    
    # Apply buffer if distance > threshold
    estimated_miles = drive_time_minutes / 1.5
    
    if estimated_miles > rules["distance_rules"]["buffer_threshold_miles"]:
        buffer_minutes = rules["distance_rules"]["buffer_minutes"]
        buffer_price = (buffer_minutes / 60) * hourly_rate
        base_price += buffer_price
    
    # Apply location-specific rules
    min_price = rules["minimums"]["general"]
    
    if primary_type == "tucson_internal" and is_internal_move:
        min_price = rules["minimums"]["tucson_internal"]
        if primary_config:
            base_price = max(base_price, primary_config.get("typical", 400))
            
    elif primary_type == "tucson_delivery" or (delivery_address and "TUCSON" in delivery_address.upper() and not is_internal_move):
        min_price = rules["minimums"]["tucson_delivery"]
        
    elif primary_type == "prison" and primary_config:
        base_price = max(base_price, primary_config.get("typical", 1000))
        min_price = primary_config.get("min_price", 900)
        
    elif primary_type == "border" and primary_config:
        base_price = max(base_price, primary_config.get("typical", 2650))
        min_price = primary_config.get("min_price", 2500)
        
    elif primary_type == "far_scottsdale" and primary_config:
        bump = (primary_config.get("bump_min", 40) + primary_config.get("bump_max", 80)) / 2
        base_price += bump
    
    # Apply multi-machine adjustment
    if num_machines > 1:
        multi_config = rules["multi_machine_rules"]["same_trip_adjustment"]
        additional_machines = num_machines - 1
        adjustment_per = (multi_config["min"] + multi_config["max"]) / 2
        base_price += additional_machines * adjustment_per
    
    # Apply minimum
    final_price = max(base_price, min_price)
    
    # Round up to nearest $25
    final_price = round_up_to_nearest(final_price, rules["rounding"]["round_to"])
    
    return {
        "base_price": round(base_price, 2),
        "min_price": min_price,
        "final_price": final_price,
        "location_type": primary_type,
        "estimated_miles": round(estimated_miles, 1),
        "num_machines": num_machines,
        "confidence": calculate_confidence(primary_type, estimated_miles, is_internal_move)
    }

def calculate_confidence(location_type, estimated_miles, is_internal_move):
    """Calculate confidence score for the quote"""
    if location_type in ["prison", "border"]:
        return "HIGH" if estimated_miles < 150 else "MEDIUM"
    elif location_type in ["tucson_internal", "tucson_delivery"]:
        return "HIGH" if is_internal_move else "MEDIUM"
    elif location_type == "standard" and estimated_miles < 50:
        return "HIGH"
    elif estimated_miles > 100:
        return "LOW"
    else:
        return "MEDIUM"

# =============================================================================
# STREAMLIT UI
# =============================================================================

st.title("🚚 BEP Pricing Calculator")
st.markdown("**Tool Box & Safe Moving** - Vending Machine Move Pricing System")

# Sidebar with rules summary
with st.sidebar:
    st.header("📋 Pricing Rules")
    st.markdown(f"""
    **Base Rate:** ${rules['base_rate']['hourly_rate']}/hour
    
    **Minimums:**
    - General: ${rules['minimums']['general']}
    - Tucson Delivery: ${rules['minimums']['tucson_delivery']}
    - Tucson Internal: ${rules['minimums']['tucson_internal']}
    
    **Distance Buffer:**
    - Threshold: {rules['distance_rules']['buffer_threshold_miles']} miles
    - Buffer: +{rules['distance_rules']['buffer_minutes']} min
    
    **Rounding:** Up to ${rules['rounding']['round_to']}
    """)
    
    st.divider()
    
    st.markdown("""
    **Location Premiums:**
    - 🏛️ Prison: $900-$1,200
    - 🌵 Border: $2,500-$2,800
    - 🏥 VA Hospital: ~$900
    """)

# Main content - tabs
tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload Sheet", "💰 Manual Calculator", "📍 Locations", "⚙️ Rules"])

with tab1:
    st.header("📤 Upload BEP Move Request Sheet")
    st.markdown("Upload the Excel file and we'll extract all the data automatically!")
    
    uploaded_file = st.file_uploader(
        "Drop your BEP Excel file here",
        type=["xlsx", "xls"],
        help="Upload the Move Request Excel file from BEP"
    )
    
    if uploaded_file is not None:
        with st.spinner("Parsing Excel file..."):
            data = parse_bep_excel(uploaded_file)
        
        if data["success"]:
            st.success("✅ File parsed successfully!")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📋 Extracted Data")
                
                # Editable fields
                requester = st.text_input("Requester Name", value=data.get("requester") or "")
                mr_number = st.text_input("MR Number", value=data.get("mr_number") or "")
                pickup = st.text_input("Pickup Address", value=data.get("pickup_address") or "")
                delivery = st.text_input("Delivery Address", value=data.get("delivery_address") or "")
                machines = st.number_input("Number of Machines", min_value=1, max_value=10, value=data.get("num_machines") or 1)
                
                if data.get("machine_type"):
                    st.text_input("Machine Type", value=data.get("machine_type"), disabled=True)
                if data.get("move_date"):
                    st.text_input("Move Date", value=data.get("move_date"), disabled=True)
                if data.get("contact_name"):
                    st.text_input("Contact Name", value=data.get("contact_name"), disabled=True)
                if data.get("notes"):
                    st.text_area("Notes", value=data.get("notes"), disabled=True)
            
            with col2:
                st.subheader("🕐 Drive Time & Calculate")
                
                drive_time = st.number_input(
                    "One-Way Drive Time (minutes)",
                    min_value=5,
                    max_value=300,
                    value=60,
                    help="Enter the one-way drive time from Maximus to delivery"
                )
                
                is_internal = st.checkbox(
                    "Internal Move (same facility)",
                    help="Check if moving within the same facility"
                )
                
                # Quick presets
                st.markdown("**Quick Presets:**")
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("Phoenix (30 min)", key="p1"):
                        drive_time = 30
                    if st.button("Tucson (90 min)", key="p2"):
                        drive_time = 90
                with col_b:
                    if st.button("Scottsdale (45 min)", key="p3"):
                        drive_time = 45
                    if st.button("Border (180 min)", key="p4"):
                        drive_time = 180
                
                st.divider()
                
                if st.button("🧮 CALCULATE QUOTE", type="primary", use_container_width=True):
                    if pickup and delivery:
                        result = calculate_price(pickup, delivery, machines, drive_time, is_internal)
                        
                        st.success(f"### 💵 Quote: ${result['final_price']:,.0f}")
                        
                        confidence_emoji = {"HIGH": "🟢", "MEDIUM": "🟡", "LOW": "🔴"}
                        st.markdown(f"**Confidence:** {confidence_emoji.get(result['confidence'], '⚪')} {result['confidence']}")
                        st.markdown(f"**Location Type:** {result['location_type'].replace('_', ' ').title()}")
                        
                        with st.expander("📋 Full Breakdown"):
                            st.markdown(f"""
                            | Field | Value |
                            |-------|-------|
                            | Requester | {requester} |
                            | MR Number | {mr_number} |
                            | Pickup | {pickup} |
                            | Delivery | {delivery} |
                            | Machines | {machines} |
                            | Drive Time | {drive_time} min |
                            | Est. Miles | {result['estimated_miles']} |
                            | Location Type | {result['location_type']} |
                            | Minimum | ${result['min_price']} |
                            | **QUOTE** | **${result['final_price']:,.0f}** |
                            """)
                    else:
                        st.error("Please enter pickup and delivery addresses")
            
            # Show raw data for debugging
            with st.expander("🔍 View Raw Excel Data"):
                st.dataframe(pd.read_excel(uploaded_file, header=None).head(50))
        else:
            st.error(f"❌ Error parsing file: {data.get('error')}")

with tab2:
    st.header("💰 Manual Quote Calculator")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📍 Move Details")
        
        pickup_address = st.text_input(
            "Pickup Address",
            placeholder="e.g., 3838 N Central Ave, Phoenix, AZ (Maximus)",
            key="manual_pickup"
        )
        
        delivery_address = st.text_input(
            "Delivery Address", 
            placeholder="e.g., 1000 S Wilmont Rd, Tucson, AZ",
            key="manual_delivery"
        )
        
        num_machines = st.number_input(
            "Number of Machines",
            min_value=1,
            max_value=10,
            value=1,
            key="manual_machines"
        )
        
        is_internal = st.checkbox(
            "Internal Move (same facility/building)",
            help="Check if moving within the same facility - uses lower Tucson internal pricing",
            key="manual_internal"
        )
    
    with col2:
        st.subheader("🕐 Drive Time")
        
        drive_time = st.number_input(
            "One-Way Drive Time (minutes)",
            min_value=5,
            max_value=300,
            value=60,
            help="Enter the one-way drive time from Maximus to delivery location",
            key="manual_drive"
        )
        
        st.info("💡 **Tip:** Use Google Maps to get accurate drive time")
    
    st.divider()
    
    if st.button("🧮 Calculate Quote", type="primary", use_container_width=True, key="manual_calc"):
        if pickup_address and delivery_address:
            result = calculate_price(
                pickup_address,
                delivery_address,
                num_machines,
                drive_time,
                is_internal
            )
            
            st.success("### Quote Generated!")
            
            result_col1, result_col2, result_col3 = st.columns(3)
            
            with result_col1:
                st.metric(label="💵 Quoted Price", value=f"${result['final_price']:,.0f}")
            
            with result_col2:
                confidence_color = {"HIGH": "🟢", "MEDIUM": "🟡", "LOW": "🔴"}
                st.metric(label="📊 Confidence", value=f"{confidence_color.get(result['confidence'], '⚪')} {result['confidence']}")
            
            with result_col3:
                st.metric(label="📍 Location Type", value=result['location_type'].replace("_", " ").title())
            
            with st.expander("📋 Price Breakdown", expanded=True):
                st.markdown(f"""
                | Component | Value |
                |-----------|-------|
                | Drive Time (one-way) | {drive_time} minutes |
                | Estimated Distance | {result['estimated_miles']} miles |
                | Number of Machines | {result['num_machines']} |
                | Location Type | {result['location_type']} |
                | Minimum Price | ${result['min_price']} |
                | Base Calculation | ${result['base_price']:,.2f} |
                | **Final Quote** | **${result['final_price']:,.0f}** |
                """)
        else:
            st.error("Please enter both pickup and delivery addresses")

with tab3:
    st.header("📍 Known Locations Database")
    
    if locations:
        for loc_id, loc_data in locations.items():
            with st.expander(f"📍 {loc_data.get('name', loc_id)}"):
                st.markdown(f"""
                - **Address:** {loc_data.get('address', 'N/A')}
                - **Type:** {loc_data.get('type', 'standard')}
                - **Typical Price:** ${loc_data.get('typical_price', 'N/A')}
                - **Notes:** {loc_data.get('notes', 'None')}
                """)
    else:
        st.info("No locations in database yet.")

with tab4:
    st.header("⚙️ Pricing Rules Configuration")
    
    st.json(rules)
    
    st.divider()
    st.markdown("""
    **Rule Version:** 1.0 (Based on 413-job analysis)
    
    To modify rules, edit `data/pricing_rules.json`
    """)

# Footer
st.divider()
st.markdown("""
<div style="text-align: center; color: gray; font-size: 12px;">
    BEP Pricing Calculator v1.1 | Tool Box & Safe Moving | Built with ❤️ by Grant
</div>
""", unsafe_allow_html=True)
