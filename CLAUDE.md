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
### Design (UI/UX)
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
