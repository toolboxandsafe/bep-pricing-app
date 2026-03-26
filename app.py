"""
BEP Pricing Calculator - Streamlit Web App
Tool Box & Safe Moving - Vending Machine Move Pricing
"""

import streamlit as st
import json
import math
from datetime import datetime
import pandas as pd

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

# Helper functions
def round_up_to_nearest(value, nearest):
    """Round up to nearest value (e.g., nearest $25)"""
    return math.ceil(value / nearest) * nearest

def detect_location_type(address):
    """Detect location type from address keywords"""
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
    # Note: We'd need actual distance, using drive time as proxy for now
    # Assuming 1 mile ≈ 1.5 minutes in Phoenix metro
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
            
    elif primary_type == "tucson_delivery" or ("TUCSON" in delivery_address.upper() and not is_internal_move):
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
        # Add $50-100 per additional machine, not 2x
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
tab1, tab2, tab3, tab4 = st.tabs(["💰 Quote Calculator", "📊 Price History", "📍 Locations", "⚙️ Rules Config"])

with tab1:
    st.header("Calculate BEP Move Quote")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📍 Move Details")
        
        pickup_address = st.text_input(
            "Pickup Address",
            placeholder="e.g., 3838 N Central Ave, Phoenix, AZ (Maximus)"
        )
        
        delivery_address = st.text_input(
            "Delivery Address", 
            placeholder="e.g., 1000 S Wilmont Rd, Tucson, AZ"
        )
        
        num_machines = st.number_input(
            "Number of Machines",
            min_value=1,
            max_value=10,
            value=1
        )
        
        is_internal = st.checkbox(
            "Internal Move (same facility/building)",
            help="Check if moving within the same facility - uses lower Tucson internal pricing"
        )
    
    with col2:
        st.subheader("🕐 Drive Time")
        
        drive_time = st.number_input(
            "One-Way Drive Time (minutes)",
            min_value=5,
            max_value=300,
            value=60,
            help="Enter the one-way drive time from Maximus to delivery location"
        )
        
        st.info("💡 **Tip:** Use Google Maps to get accurate drive time")
        
        # Quick presets
        st.markdown("**Quick Presets:**")
        preset_col1, preset_col2 = st.columns(2)
        with preset_col1:
            if st.button("Phoenix Metro (30 min)"):
                drive_time = 30
            if st.button("Tucson (90 min)"):
                drive_time = 90
        with preset_col2:
            if st.button("Far Scottsdale (45 min)"):
                drive_time = 45
            if st.button("Border (180 min)"):
                drive_time = 180
    
    st.divider()
    
    # Calculate button
    if st.button("🧮 Calculate Quote", type="primary", use_container_width=True):
        if pickup_address and delivery_address:
            result = calculate_price(
                pickup_address,
                delivery_address,
                num_machines,
                drive_time,
                is_internal
            )
            
            # Display results
            st.success("### Quote Generated!")
            
            result_col1, result_col2, result_col3 = st.columns(3)
            
            with result_col1:
                st.metric(
                    label="💵 Quoted Price",
                    value=f"${result['final_price']:,.0f}"
                )
            
            with result_col2:
                confidence_color = {
                    "HIGH": "🟢",
                    "MEDIUM": "🟡", 
                    "LOW": "🔴"
                }
                st.metric(
                    label="📊 Confidence",
                    value=f"{confidence_color.get(result['confidence'], '⚪')} {result['confidence']}"
                )
            
            with result_col3:
                st.metric(
                    label="📍 Location Type",
                    value=result['location_type'].replace("_", " ").title()
                )
            
            # Breakdown
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
            
            # Warnings/notes
            if result['confidence'] == "LOW":
                st.warning("⚠️ **Low Confidence Quote** - Consider manual review before sending")
            
            if result['location_type'] == "prison":
                st.info("🏛️ **Prison Location Detected** - Typical range: $900-$1,200")
            
            if result['location_type'] == "tucson_internal":
                st.info("🏢 **Tucson Internal Move** - Using reduced pricing ($300-$500, NOT $850)")
                
        else:
            st.error("Please enter both pickup and delivery addresses")

with tab2:
    st.header("📊 Quote History")
    st.info("Quote history will be stored here once we integrate with the email workflow.")
    
    # Placeholder for history
    st.markdown("""
    **Coming Soon:**
    - View all generated quotes
    - Track Ryan's adjustments
    - Analyze pricing patterns
    - Export to CSV/Excel
    """)

with tab3:
    st.header("📍 Known Locations Database")
    
    if locations:
        df = pd.DataFrame(locations).T
        st.dataframe(df, use_container_width=True)
    else:
        st.info("No locations in database yet. They will be added as quotes are generated.")
    
    # Add new location form
    with st.expander("➕ Add New Location"):
        new_name = st.text_input("Location Name")
        new_address = st.text_input("Full Address")
        new_type = st.selectbox("Location Type", 
            ["standard", "prison", "va_hospital", "tucson_internal", "tucson_delivery", "border", "far_scottsdale"])
        new_typical_price = st.number_input("Typical Price", min_value=0, value=400)
        
        if st.button("Add Location"):
            st.success(f"Location '{new_name}' added! (Feature coming soon)")

with tab4:
    st.header("⚙️ Pricing Rules Configuration")
    
    st.json(rules)
    
    st.divider()
    
    st.markdown("""
    **To modify rules:**
    1. Edit `data/pricing_rules.json`
    2. Restart the app to apply changes
    
    **Rule Version:** 1.0 (Based on 413-job analysis)
    """)

# Footer
st.divider()
st.markdown("""
<div style="text-align: center; color: gray; font-size: 12px;">
    BEP Pricing Calculator v1.0 | Tool Box & Safe Moving | Built with ❤️ by Grant
</div>
""", unsafe_allow_html=True)
