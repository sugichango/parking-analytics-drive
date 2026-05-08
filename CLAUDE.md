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
- If Segmentation Fault / "This app has gone over its resource limits": memory exhaustion on Streamlit Cloud free tier (~1GB). See Memory Management section below.

---

## Deployment Workflow (IMPORTANT)

Streamlit Cloud runs the **root-level** `unified_dashboard_google_drive.py`, NOT `parking_v2_public/unified_dashboard_google_drive.py`.

**Always follow this sequence before pushing:**
```
python -m py_compile parking_v2_public/unified_dashboard_google_drive.py
cp parking_v2_public/unified_dashboard_google_drive.py unified_dashboard_google_drive.py
git add unified_dashboard_google_drive.py parking_v2_public/unified_dashboard_google_drive.py
git commit -m "..."
git push
```

---

## Implemented Features (as of 2026-05-08)

### Dashboard ① 一般利用台数推移分析
- **単年分析グラフ**: Monthly/daily bar+line chart with payment type breakdown (現金/RB/回数券).
- **年度間比較分析** (FY2023〜FY2025):
  - Hidden behind a checkbox `📊 年度間比較分析を表示する` (value=False by default).
  - Clicking the checkbox triggers `load_comparison_monthly()` which loads CSVs on demand.
  - Shows: KPI panel (total 台数・収入 with YoY %), monthly line chart (April–March x-axis), parking-by-year table.
  - YoY display: positive = `+X.X%` (green), negative = `▼X.X%` (red).
  - Table has a vertical divider line between 台数 columns and 収入 columns.

### Dashboard ② 24時間稼働状況分析
- 5 tabs: waveform comparison, weekday/holiday, regular/general, year-over-year, comprehensive view.
- All KPI metrics have `help=` tooltips explaining calculation formulas.
- `common_layout` dict is defined at the top of the `else:` block (not shared with ①).

### Sidebar
- Manual (`📖 使い方マニュアルを表示`) is embedded directly in the app code as `_MANUAL_TEXT` — no Google Drive dependency.
- Cache clear button.
- Logout button.

---

## Memory Management (Streamlit Cloud Free Tier ~1GB)

Segmentation faults occur when memory is exhausted. Mitigations applied:

1. **Lazy loading**: Comparison section only loads data when checkbox is checked.
2. **Lightweight comparison loader** (`load_comparison_monthly`):
   - Does NOT call `load_data_dashboard1_drive` (which caches full DataFrame).
   - Downloads each year's CSV directly, reads only 3 columns (ParkingArea, OnTime, Cash).
   - Aggregates to monthly level immediately, then `del df; gc.collect()`.
   - Caches only the tiny aggregated result (~100 rows).
   - Takes `_file_keys` as `tuple of (filename, file_id)` — file_id used for direct download.
3. **Discount column cleanup**: After computing PaymentType, Discount1-7 columns are dropped from the DataFrame before caching (`df.drop(columns=active_disc_cols, inplace=True)`). Reduces cached size by ~35%.

---

## Communication Rule with User
- Messages wrapped in 【 】 are real instructions. Respond normally.
- Messages WITHOUT 【 】 brackets should receive only "続きをお待ちしています" as response.
