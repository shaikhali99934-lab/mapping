# GeoSync: Field Data Entry with GPS Coordinates

A web application designed for field data collection that captures **Center_Code**, **Reg No**, and **Coordinates** (with high-accuracy device GPS) and appends rows directly to Google Sheets.

---

## 📋 Google Sheet Column Setup
Ensure Row 1 of your Google Sheet has these headers:
| Column A | Column B | Column C | Column D (Optional) |
| :--- | :--- | :--- | :--- |
| `Center_Code` | `Reg No` | `Coordinates` | `Timestamp` |

---

## 🚀 Quick Setup (2 Minutes)

### Step 1: Update Google Apps Script
1. In your Google Sheet, open **Extensions** → **Apps Script**.
2. Replace all existing code with the updated content from [Code.gs](file:///c:/Users/shah/Desktop/sql/google%20sheet%20database/Code.gs).
3. Click the **Save** icon.

### Step 2: Deploy New Version
1. Click **Deploy** → **Manage deployments** (or **New deployment**).
2. Edit deployment or create a new **Web app** deployment:
   - **Execute as**: `Me (your Google email)`
   - **Who has access**: `Anyone`
3. Click **Deploy** and copy your **Web App URL** (ends in `/exec`).
*(If you see "Google hasn't verified this app", click **Advanced** → **Go to Untitled project (unsafe)** → **Allow**).*

### Step 3: Link & Collect Data
1. Open the form in your browser at `http://localhost:3000/index.html` (or open [index.html](file:///c:/Users/shah/Desktop/sql/google%20sheet%20database/index.html)).
2. Click **Setup Sheet URL** at the top right, paste your URL, and click **Save & Apply**.
3. Fill in:
   - **Center Code** (e.g. `CC-104`)
   - **Registration Number** (e.g. `REG-2024-8841`)
   - Click **📡 Capture GPS** to fetch high-precision coordinates with accuracy (e.g., `±3.5m`), or enter manually.
4. Click **Submit to Google Sheet** (or press `Ctrl + Enter`).

---

## 🌟 Key Features
- **Field GPS Location Capture**: 1-click device GPS coordinate acquisition (`lat, lng`) with precision up to ~10cm and live accuracy indicator (`±X meters`).
- **Map Previews**: Instant links to view any coordinate directly on Google Maps.
- **Offline & History Log**: View previous entries in a table with direct Google Map links and CSV export.
- **Instant Cloud Sync**: Directly writes to Google Sheet without intermediate databases or server configuration.
