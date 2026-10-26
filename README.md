# GeoSync: Offline-First Field Data Entry & GPS Portal

An offline-capable Progressive Web Application (PWA) designed for field teams to record **Center_Code**, **Reg No**, and **Coordinates** in remote locations without internet, automatically syncing with Google Sheets when connectivity returns.

---

## 🔒 App Access Password
The application is protected with an authorization gate:
- **Password**: `malcpak`
- Users can choose to stay logged in on their device.
- Users can lock the app or log out anytime by tapping the **Lock** icon in the top header.

---

## 📡 How Offline Mode Works in the Field

1. **Works Completely Without Internet (PWA)**:
   - All files (`index.html`, `style.css`, `app.js`) are cached on the device via a Service Worker ([sw.js](file:///c:/Users/shah/Desktop/Master%20Dashboards/Other%20Tasks/google%20sheet%20database/sw.js)).
   - You can open the form and use it in airplanes, deserts, mountains, or basements with zero bars of signal.

2. **Satellite GPS Needs No Internet**:
   - Device GPS communicates directly with global positioning satellites in space, requiring no mobile cellular data.
   - Tap **"Capture Current Location"** to record high-precision latitude and longitude coordinates anywhere.

3. **Offline Queue & Automatic Cloud Sync**:
   - Submissions made without internet are safely stored on the device with the status **`⏳ Offline (Pending)`**.
   - As soon as the device reconnects to Wi-Fi or mobile data (3G/4G), GeoSync automatically uploads every pending record into your Google Sheet!
   - You can also tap the **Sync Now** button in the header at any time.

4. **Install as a Mobile App (Home Screen)**:
   - On Chrome (Android): Tap the 3 dots menu &rarr; **Add to Home screen** / **Install app**.
   - On Safari (iPhone): Tap the Share button &rarr; **Add to Home Screen**.
   - The app will open as a full-screen, native-feeling app without browser address bars!

---

## 📋 Google Sheet Setup & Update Guide

### Column Headers (Row 1):
Ensure your Google Sheet tab has the following 8 columns in Row 1:

| Column A | Column B | Column C | Column D | Column E | Column F | Column G | Column H |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `Center_Code` | `Reg No` | `Coordinates` | `Date of Mapping` | `Province` | `District` | `Name of Staff` | `Timestamp` |

*(Note: The updated script in [Code.gs](file:///c:/Users/shah/Desktop/Master%20Dashboards/Other%20Tasks/google%20sheet%20database/Code.gs) also features automatic header detection — if any columns are missing, it will automatically append them!)*

---

### ⚠️ IMPORTANT: How to Update Google Apps Script (Why New Fields Don't Record Yet)
In Google Apps Script, simply pasting new code **does NOT update the live Web App** until you create a **New Version**. If you only saved the code, Google continues running your previous 4-column deployment.

Follow these exact steps:

1. Open your Google Sheet &rarr; Click **Extensions** &rarr; **Apps Script**.
2. Select all code in `Code.gs` and replace it with the code from [Code.gs](file:///c:/Users/shah/Desktop/Master%20Dashboards/Other%20Tasks/google%20sheet%20database/Code.gs) or [linkage code.txt](file:///c:/Users/shah/Desktop/Master%20Dashboards/Other%20Tasks/google%20sheet%20database/linkage%20code.txt).
3. Click the **Save** (💾 disk) icon.
4. Click the blue **Deploy** button (top right) &rarr; Select **Manage deployments**.
5. In the modal, click the **Edit (pencil icon)** next to your active deployment.
6. Under **Version**, click the dropdown and choose **New version**.
7. Keep **Execute as: Me** and **Who has access: Anyone**.
8. Click **Deploy**.
9. Copy your **Web App URL** and paste it into **Settings** (⚙️) inside GeoSync. Test submitting a record!

