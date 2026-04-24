# Project: Parking Analytics Dashboard (Google Drive Link)

This project provides a Streamlit dashboard for analyzing parking usage data synced from Google Drive.

## Build and Deployment
- **Python Version**: 3.12 (Pinned to avoid dependency build failures on Streamlit Cloud)
- **Deployment Platform**: Streamlit Cloud
- **Secrets Management**: Requires `gcp_service_account` (JSON string), `google_drive_folder_id`, and `app_password`.

## Commands
- **Run Locally**: `streamlit run unified_dashboard_google_drive.py`
- **Check Syntax**: `python -m py_compile unified_dashboard_google_drive.py`

## Project Specific Rules
### 1. Data Processing Logic (Domain Knowledge)
- **Payment Type Determination**:
    - **RB**: If `Discount` codes include `[11, 12, 13, 14, 15, 43, 44]`.
    - **Ticket**: If `Discount` codes include `[30, 31, 32, 33, 34, 35]` (and NOT RB).
    - **Cash**: If no discount is applied.
- **Encoding**: Always try `UTF-8` first, then fallback to `CP932`.
- **Date Handling**: Convert `OnTime` to datetime, then create `is_holiday` flag based on DayOfWeek (5, 6).

### 2. Parking Codes Map
- 440: 南１ (South 1)
- 441: 南２ (South 2)
- 442: 南３ (South 3)
- 443: 南４ (South 4)
- 444: 北１ (North 1)
- 445: 北２ (North 2)
- 446: 北３ (North 3)

### 3. Design (UI/UX)
- Use high-contrast CSS for metrics.
- Font: Inter (via Google Fonts).
- Primary metrics should have a `text-shadow` for readability on dark backgrounds.

### Data
- Primary file: `updated_integrated_data_FY2025.csv.gz` (Google Drive)
- Analytics Excel: `*_with_avg.xlsx` (Google Drive)
- Parking Codes: 440-443 (South 1-4), 444-446 (North 1-3).

### Error Troubleshooting
- If "Oh no" on deploy, check Python version in settings (must be 3.12).
- If 403, enable Google Drive API in GCP.
- If 404, check folder ID spelling in Secrets.
