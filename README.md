# 🚚 BEP Pricing Calculator

**Tool Box & Safe Moving** - Automated Vending Machine Move Pricing System

## Overview

A Streamlit web app that calculates accurate BEP (Business Enterprise Program) vending machine move quotes based on established pricing rules derived from analysis of 413+ historical jobs.

## Features

- 💰 **Quote Calculator** - Enter move details, get instant pricing
- 📊 **Confidence Scoring** - High/Medium/Low confidence indicators
- 📍 **Location Detection** - Auto-detects prisons, VA, Tucson, border locations
- 🧮 **Smart Pricing Rules** - Handles minimums, buffers, multi-machine adjustments
- 📋 **Price Breakdown** - Transparent calculation details

## Pricing Rules

Based on 413-job analysis:

| Rule | Value |
|------|-------|
| Base Rate | $170/hour |
| General Minimum | $220 |
| Tucson Delivery | $850 minimum |
| Tucson Internal | $300-500 |
| Prison Jobs | $900-1,200 |
| Border Jobs | $2,500-2,800 |
| Distance Buffer | +20min over 35 miles |
| Rounding | Up to nearest $25 |

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

## Deployment

Deployed on Streamlit Cloud:
- Push to GitHub
- Connect repo to streamlit.io
- App auto-deploys on push

## Files

```
bep-pricing-app/
├── app.py                 # Main Streamlit app
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── data/
    ├── pricing_rules.json    # Pricing configuration
    └── known_locations.json  # Location database
```

## Future Enhancements

- [ ] Google Maps API integration for auto-distance
- [ ] Gmail integration for email parsing
- [ ] Trello integration for card creation
- [ ] Quote history tracking
- [ ] Ryan adjustment learning
- [ ] PDF quote generation

---

Built with ❤️ by Grant for Tool Box & Safe Moving
# Redeployed: Mon Mar 30 10:03:47 MST 2026
