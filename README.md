# GeoSync: Offline-First Field Data Entry & GPS Portal

An offline-capable Progressive Web Application (PWA) designed for field teams to record **Center_Code**, **Reg No**, and **Coordinates** in remote locations without internet, automatically syncing with Google Sheets when connectivity returns.

---

## 📡 How Offline Mode Works in the Field

1. **Works Completely Without Internet (PWA)**:
   - All files (`index.html`, `style.css`, `app.js`) are cached on the device via a Service Worker ([sw.js](file:///c:/Users/shah/Desktop/sql/google%20sheet%20database/sw.js)).
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

## 📋 Google Sheet Setup (One-Time)

### Column Headers (Row 1):
| Column A | Column B | Column C | Column D |
| :--- | :--- | :--- | :--- |
| `Center_Code` | `Reg No` | `Coordinates` | `Timestamp` |

### Apps Script Code:
1. In your sheet, go to **Extensions** &rarr; **Apps Script**.
2. Replace with the code from [Code.gs](file:///c:/Users/shah/Desktop/sql/google%20sheet%20database/Code.gs).
3. Click **Deploy** &rarr; **New deployment** (or **Manage deployments**) &rarr; Select **Web app** &rarr; **Who has access: Anyone** &rarr; **Deploy**.
4. Paste the URL into the **Settings** modal in the app.
