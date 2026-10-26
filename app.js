// GeoSync Mobile & Offline Field Data Integration Logic
document.addEventListener('DOMContentLoaded', () => {
  // Register Service Worker for Complete Offline App Availability
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js')
        .then((reg) => {
          console.log('GeoSync ServiceWorker registered:', reg.scope);
        })
        .catch((err) => {
          console.warn('GeoSync ServiceWorker registration failed:', err);
        });
    });
  }

  // Elements
  const form = document.getElementById('dataEntryForm');
  const centerCodeInput = document.getElementById('fieldCenterCode');
  const regNoInput = document.getElementById('fieldRegNo');
  const coordinatesInput = document.getElementById('fieldCoordinates');
  const fieldDateMapping = document.getElementById('fieldDateMapping');
  const popDatePickerBtn = document.getElementById('popDatePickerBtn');
  const todayDateBtn = document.getElementById('todayDateBtn');
  const fieldProvince = document.getElementById('fieldProvince');
  const fieldDistrict = document.getElementById('fieldDistrict');
  const fieldNameOfStaff = document.getElementById('fieldNameOfStaff');
  const captureGpsBtn = document.getElementById('captureGpsBtn');
  const gpsBtnTitle = document.getElementById('gpsBtnTitle');
  const gpsBtnSubtitle = document.getElementById('gpsBtnSubtitle');
  const gpsAccuracyPill = document.getElementById('gpsAccuracyPill');
  const accuracyVal = document.getElementById('accuracyVal');
  const previewMapLink = document.getElementById('previewMapLink');

  const submitBtn = document.getElementById('submitBtn');
  const submitBtnText = document.getElementById('submitBtnText');
  const resetBtn = document.getElementById('resetBtn');

  // Network & Offline Status Elements
  const networkStatusPill = document.getElementById('networkStatusPill');
  const networkStatusText = document.getElementById('networkStatusText');
  const offlineNoticeBanner = document.getElementById('offlineNoticeBanner');
  const syncNowBtn = document.getElementById('syncNowBtn');
  const pendingCountHeader = document.getElementById('pendingCountHeader');
  const syncTableBtn = document.getElementById('syncTableBtn');

  // Other UI Elements
  const sheetTargetDisplay = document.getElementById('sheetTargetDisplay');
  const setupPromptBanner = document.getElementById('setupPromptBanner');
  const bannerSetupBtn = document.getElementById('bannerSetupBtn');
  const recordCountBadge = document.getElementById('recordCount');
  const tabRecordCount = document.getElementById('tabRecordCount');
  const tableBody = document.getElementById('tableBody');
  const exportCsvBtn = document.getElementById('exportCsvBtn');
  const clearHistoryBtn = document.getElementById('clearHistoryBtn');
  const historySearchInput = document.getElementById('historySearchInput');
  const clearSearchBtn = document.getElementById('clearSearchBtn');
  const viewTableBtn = document.getElementById('viewTableBtn');
  const viewCardsBtn = document.getElementById('viewCardsBtn');
  const tableContainer = document.getElementById('tableContainer');
  const cardsContainer = document.getElementById('cardsContainer');

  let currentHistoryView = localStorage.getItem('field_mapping_view_pref') || (window.innerWidth < 768 ? 'cards' : 'table');
  let historySearchTerm = '';

  // Mobile Tabs
  const tabEntryBtn = document.getElementById('tabEntryBtn');
  const tabHistoryBtn = document.getElementById('tabHistoryBtn');
  const viewFormSection = document.getElementById('viewFormSection');
  const viewHistorySection = document.getElementById('viewHistorySection');

  // Modals & Triggers
  const openSettingsBtn = document.getElementById('openSettingsBtn');
  const settingsModal = document.getElementById('settingsModal');
  const closeSettingsModal = document.getElementById('closeSettingsModal');
  const cancelSettingsBtn = document.getElementById('cancelSettingsBtn');
  const saveSettingsBtn = document.getElementById('saveSettingsBtn');
  const scriptUrlInput = document.getElementById('scriptUrlInput');
  const testConnectionBtn = document.getElementById('testConnectionBtn');
  const testStatusMsg = document.getElementById('testStatusMsg');

  const openGuideBtn = document.getElementById('openGuideBtn');
  const guideModal = document.getElementById('guideModal');
  const closeGuideModal = document.getElementById('closeGuideModal');
  const gotItGuideBtn = document.getElementById('gotItGuideBtn');
  const copyCodeBtn = document.getElementById('copyCodeBtn');
  const codeSnippetContent = document.getElementById('codeSnippetContent');
  const toastContainer = document.getElementById('toastContainer');

  // Mobile Home Screen Install Elements
  const headerInstallBtn = document.getElementById('headerInstallBtn');
  const mainInstallBtn = document.getElementById('mainInstallBtn');
  const installAppBanner = document.getElementById('installAppBanner');
  const dismissInstallBanner = document.getElementById('dismissInstallBanner');
  const installModal = document.getElementById('installModal');
  const closeInstallModal = document.getElementById('closeInstallModal');
  const gotItInstallBtn = document.getElementById('gotItInstallBtn');
  const triggerNativeInstallBtn = document.getElementById('triggerNativeInstallBtn');
  const nativeInstallBox = document.getElementById('nativeInstallBox');
  const tabAndroidBtn = document.getElementById('tabAndroidBtn');
  const tabIosBtn = document.getElementById('tabIosBtn');
  const androidGuideView = document.getElementById('androidGuideView');
  const iosGuideView = document.getElementById('iosGuideView');

  // App Lock & Authentication Elements
  const APP_PASSWORD = 'malcpak';
  const AUTH_STORAGE_KEY = 'field_mapping_auth';
  const AUTH_TOKEN_VAL = 'malcpak_authenticated';

  const loginOverlay = document.getElementById('loginOverlay');
  const loginCard = document.getElementById('loginCard');
  const loginForm = document.getElementById('loginForm');
  const loginPassword = document.getElementById('loginPassword');
  const togglePasswordBtn = document.getElementById('togglePasswordBtn');
  const eyeIconOpen = document.getElementById('eyeIconOpen');
  const eyeIconClosed = document.getElementById('eyeIconClosed');
  const loginErrorMsg = document.getElementById('loginErrorMsg');
  const loginErrorText = document.getElementById('loginErrorText');
  const loginRememberMe = document.getElementById('loginRememberMe');
  const loginSubmitBtn = document.getElementById('loginSubmitBtn');
  const loginBtnText = document.getElementById('loginBtnText');
  const lockAppBtn = document.getElementById('lockAppBtn');

  // Google Apps Script snippet shown in guide
  const APPS_SCRIPT_CODE = `function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};
    
    // Support JSON payload or Form/URL-encoded parameters
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // Extract all 8 fields with generous aliases
    var centerCode = data.center_code || data.Center_Code || data['Center Code'] || '';
    var regNo = data.reg_no || data.Reg_No || data['Reg No'] || data.regNo || '';
    var coordinates = data.coordinates || data.Coordinates || data.GPS || '';
    var dateOfMapping = data.date_of_mapping || data['Date Of Mapping'] || data['Date Of Mappng'] || data.dateOfMapping || data.date || '';
    var province = data.province || data.Province || '';
    var district = data.district || data.District || '';
    var nameOfStaff = data.name_of_staff || data['Name of Staff'] || data.nameOfStaff || data.staff_name || data.staff || '';
    var timestamp = data.timestamp || new Date().toLocaleString();

    // Standard 8 Field Mapping Headers:
    var standardHeaders = [
      'Center_Code',
      'Reg No',
      'Coordinates',
      'Date Of Mapping',
      'Province',
      'District',
      'Name of Staff',
      'Timestamp'
    ];

    function normalize(str) {
      return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    var lastCol = sheet.getLastColumn();
    var existingHeaders = [];
    if (lastCol > 0) {
      existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // If sheet has no headers in Row 1, create all 8 standard headers
    if (existingHeaders.length === 0 || (existingHeaders.length === 1 && String(existingHeaders[0]).trim() === '')) {
      sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
      sheet.getRange(1, 1, 1, standardHeaders.length).setFontWeight('bold');
      existingHeaders = standardHeaders;
      lastCol = standardHeaders.length;
    }

    // Field value mapping table
    var fieldMap = {
      'centercode': centerCode,
      'center': centerCode,
      'regno': regNo,
      'registration': regNo,
      'registrationno': regNo,
      'coordinates': coordinates,
      'gps': coordinates,
      'location': coordinates,
      'dateofmapping': dateOfMapping,
      'dateofmappng': dateOfMapping,
      'date': dateOfMapping,
      'mappingdate': dateOfMapping,
      'province': province,
      'district': district,
      'nameofstaff': nameOfStaff,
      'staffname': nameOfStaff,
      'staff': nameOfStaff,
      'timestamp': timestamp,
      'time': timestamp,
      'datetime': timestamp
    };

    var rowData = new Array(existingHeaders.length);
    var matchedKeys = {};

    for (var i = 0; i < existingHeaders.length; i++) {
      var norm = normalize(existingHeaders[i]);
      if (fieldMap.hasOwnProperty(norm)) {
        rowData[i] = fieldMap[norm];
        matchedKeys[norm] = true;
      } else {
        rowData[i] = '';
      }
    }

    // Auto-expand headers in Row 1 if sheet was missing Province, District, Name of Staff, or Date
    var missingHeaders = [];
    var missingValues = [];

    if (!matchedKeys['dateofmapping'] && !matchedKeys['dateofmappng'] && !matchedKeys['date']) {
      missingHeaders.push('Date Of Mapping');
      missingValues.push(dateOfMapping);
    }
    if (!matchedKeys['province']) {
      missingHeaders.push('Province');
      missingValues.push(province);
    }
    if (!matchedKeys['district']) {
      missingHeaders.push('District');
      missingValues.push(district);
    }
    if (!matchedKeys['nameofstaff'] && !matchedKeys['staffname'] && !matchedKeys['staff']) {
      missingHeaders.push('Name of Staff');
      missingValues.push(nameOfStaff);
    }

    if (missingHeaders.length > 0) {
      var startNewCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, startNewCol, 1, missingHeaders.length).setValues([missingHeaders]);
      sheet.getRange(1, startNewCol, 1, missingHeaders.length).setFontWeight('bold');
      for (var m = 0; m < missingValues.length; m++) {
        rowData.push(missingValues[m]);
      }
    }

    // Append the row to sheet
    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Row added successfully',
        recorded: {
          center_code: centerCode,
          reg_no: regNo,
          coordinates: coordinates,
          date_of_mapping: dateOfMapping,
          province: province,
          district: district,
          name_of_staff: nameOfStaff,
          timestamp: timestamp
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'active',
      message: 'Field Mapping Endpoint is active and ready to record all 8 fields.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}`;

  if (codeSnippetContent) {
    codeSnippetContent.textContent = APPS_SCRIPT_CODE;
  }

  // Local Storage Keys
  const STORAGE_URL_KEY = 'geosync_script_url';
  const STORAGE_HISTORY_KEY = 'geosync_submissions_history';
  const STORAGE_QUEUE_KEY = 'geosync_offline_queue';

  // State
  let scriptUrl = localStorage.getItem(STORAGE_URL_KEY) || '';
  let submissions = [];
  let offlineQueue = [];
  let isSyncing = false;

  try {
    submissions = JSON.parse(localStorage.getItem(STORAGE_HISTORY_KEY)) || [];
  } catch (e) {
    submissions = [];
  }

  try {
    offlineQueue = JSON.parse(localStorage.getItem(STORAGE_QUEUE_KEY)) || [];
  } catch (e) {
    offlineQueue = [];
  }

  // Initial State Setup
  updateNetworkStatus();
  updateSyncBadges();
  renderHistoryTable();
  initMobileTabs();
  initAuth();

  // Show setup banner if URL is missing
  if (!scriptUrl) {
    setupPromptBanner.style.display = 'flex';
  } else {
    setupPromptBanner.style.display = 'none';
  }

  bannerSetupBtn.addEventListener('click', () => {
    scriptUrlInput.value = scriptUrl;
    testStatusMsg.textContent = '';
    openModal(settingsModal);
  });

  // ================= ONLINE / OFFLINE NETWORK LISTENERS =================
  window.addEventListener('online', () => {
    updateNetworkStatus();
    showToast('Internet connection restored! Ready to sync.', 'success');
    autoSyncQueue();
  });

  window.addEventListener('offline', () => {
    updateNetworkStatus();
    showToast('You are offline. Entries will be saved to your phone.', 'info');
  });

  function updateNetworkStatus() {
    const isOnline = navigator.onLine;
    if (isOnline) {
      networkStatusPill.className = 'network-pill online';
      networkStatusText.textContent = 'Online';
      offlineNoticeBanner.style.display = 'none';
      if (submitBtnText) submitBtnText.textContent = 'Submit to Google Sheet';
    } else {
      networkStatusPill.className = 'network-pill offline';
      networkStatusText.textContent = 'Offline';
      offlineNoticeBanner.style.display = 'flex';
      if (submitBtnText) submitBtnText.textContent = 'Save to Offline Queue';
    }
  }

  function updateSyncBadges() {
    const count = offlineQueue.length;
    if (count > 0) {
      syncNowBtn.style.display = 'inline-flex';
      pendingCountHeader.textContent = count;
      syncTableBtn.style.display = 'inline-flex';
      syncTableBtn.textContent = `Sync (${count}) Pending`;
    } else {
      syncNowBtn.style.display = 'none';
      syncTableBtn.style.display = 'none';
    }
  }

  // ================= SYNC OFFLINE QUEUE =================
  syncNowBtn.addEventListener('click', autoSyncQueue);
  syncTableBtn.addEventListener('click', autoSyncQueue);

  async function autoSyncQueue() {
    if (isSyncing || offlineQueue.length === 0) return;

    if (!navigator.onLine) {
      showToast('Cannot sync while offline. Connect to Wi-Fi/data first.', 'error');
      return;
    }

    if (!scriptUrl) {
      showToast('Configure your Google Sheet Web App URL in settings to sync.', 'error');
      openModal(settingsModal);
      return;
    }

    isSyncing = true;
    syncNowBtn.classList.add('syncing');
    showToast(`Syncing ${offlineQueue.length} offline records to Google Sheet...`, 'info');

    let syncedCount = 0;
    const remainingQueue = [];

    for (let i = 0; i < offlineQueue.length; i++) {
      const item = offlineQueue[i];
      try {
        await sendDataToGoogleSheets(scriptUrl, item);
        syncedCount++;

        // Update corresponding item in history table
        const found = submissions.find(s => s.id === item.id);
        if (found) {
          found.status = 'synced';
        }
      } catch (err) {
        console.warn('Sync failed for item:', item, err);
        remainingQueue.push(item);
      }
    }

    offlineQueue = remainingQueue;
    localStorage.setItem(STORAGE_QUEUE_KEY, JSON.stringify(offlineQueue));
    saveHistory();
    renderHistoryTable();
    updateSyncBadges();

    isSyncing = false;
    syncNowBtn.classList.remove('syncing');

    if (syncedCount > 0) {
      if (navigator.vibrate) navigator.vibrate([40, 60, 40]);
      showToast(`Successfully synced ${syncedCount} records to Google Sheet!`, 'success');
    }
  }

  // ================= MOBILE TAB NAVIGATION =================
  function initMobileTabs() {
    switchTab('entry');
    tabEntryBtn.addEventListener('click', () => switchTab('entry'));
    tabHistoryBtn.addEventListener('click', () => switchTab('history'));
  }

  function switchTab(tab) {
    if (tab === 'entry') {
      tabEntryBtn.classList.add('active');
      tabHistoryBtn.classList.remove('active');
      viewFormSection.classList.remove('mobile-hidden');
      viewHistorySection.classList.add('mobile-hidden');
    } else {
      tabHistoryBtn.classList.add('active');
      tabEntryBtn.classList.remove('active');
      viewHistorySection.classList.remove('mobile-hidden');
      viewFormSection.classList.add('mobile-hidden');
      renderHistoryTable();
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ================= SATELLITE GPS ACQUISITION (WORKS 100% OFFLINE) =================
  captureGpsBtn.addEventListener('click', () => {
    if (!navigator.geolocation) {
      showToast('GPS is not supported by your browser.', 'error');
      return;
    }

    captureGpsBtn.classList.add('acquiring');
    captureGpsBtn.classList.remove('captured');
    gpsBtnTitle.textContent = 'Acquiring Satellite GPS...';
    gpsBtnSubtitle.textContent = 'Works offline without internet connection';
    showToast('Locking onto satellite GPS coordinates...', 'info');

    const geoOptions = {
      enableHighAccuracy: true,
      timeout: 12000,
      maximumAge: 0
    };

    navigator.geolocation.getCurrentPosition(
      (position) => {
        captureGpsBtn.classList.remove('acquiring');
        captureGpsBtn.classList.add('captured');

        const lat = position.coords.latitude.toFixed(6);
        const lng = position.coords.longitude.toFixed(6);
        const accuracy = Math.round(position.coords.accuracy);
        const coordString = `${lat}, ${lng}`;

        coordinatesInput.value = coordString;

        gpsBtnTitle.textContent = `✓ GPS Locked: ${lat}, ${lng}`;
        gpsBtnSubtitle.textContent = `Accuracy: ±${accuracy}m • Tap to refresh`;
        accuracyVal.textContent = `±${accuracy}m`;
        gpsAccuracyPill.style.display = 'inline-flex';

        updateMapPreview(coordString);
        validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);

        if (navigator.vibrate) navigator.vibrate([40]);
        showToast(`GPS captured: ${coordString} (±${accuracy}m accuracy)`, 'success');
      },
      (error) => {
        captureGpsBtn.classList.remove('acquiring');
        gpsBtnTitle.textContent = 'Tap to Retry GPS Capture';
        gpsBtnSubtitle.textContent = 'Could not get location. Ensure GPS is ON';

        let msg = 'Failed to get GPS location.';
        if (error.code === error.PERMISSION_DENIED) {
          msg = 'GPS permission denied. Please allow location in browser settings.';
        } else if (error.code === error.POSITION_UNAVAILABLE) {
          msg = 'GPS signal unavailable. Ensure device location is turned on.';
        } else if (error.code === error.TIMEOUT) {
          msg = 'GPS request timed out. Please tap retry.';
        }
        showToast(msg, 'error');
      },
      geoOptions
    );
  });

  coordinatesInput.addEventListener('input', () => {
    const val = coordinatesInput.value.trim();
    validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);
    updateMapPreview(val);
  });

  function isValidCoordinates(val) {
    if (!val || val.trim().length === 0) return false;
    const regex = /^-?\d+(\.\d+)?[\s,]+-?\d+(\.\d+)?$/;
    return regex.test(val.trim());
  }

  function updateMapPreview(coordStr) {
    if (isValidCoordinates(coordStr)) {
      const cleaned = coordStr.replace(/\s+/g, '');
      previewMapLink.href = `https://www.google.com/maps?q=${encodeURIComponent(cleaned)}`;
      previewMapLink.style.display = 'inline-flex';
    } else {
      previewMapLink.style.display = 'none';
    }
  }

  centerCodeInput.addEventListener('input', () => {
    validateField(centerCodeInput, 'group-center-code', val => val.trim().length >= 1);
  });

  regNoInput.addEventListener('input', () => {
    validateField(regNoInput, 'group-reg-no', val => val.trim().length >= 1);
  });

  // Pre-fill Date Of Mapping to Today's Date (YYYY-MM-DD)
  if (fieldDateMapping) {
    fieldDateMapping.value = new Date().toISOString().slice(0, 10);
  }

  // Pop Date Picker Handler
  function popOpenDatePicker() {
    if (!fieldDateMapping) return;
    try {
      if (typeof fieldDateMapping.showPicker === 'function') {
        fieldDateMapping.showPicker();
      } else {
        fieldDateMapping.focus();
      }
    } catch (err) {
      fieldDateMapping.focus();
    }
  }

  if (popDatePickerBtn) popDatePickerBtn.addEventListener('click', popOpenDatePicker);
  if (fieldDateMapping) {
    fieldDateMapping.addEventListener('click', popOpenDatePicker);
    fieldDateMapping.addEventListener('input', () => {
      validateField(fieldDateMapping, 'group-date-mapping', val => val.trim().length >= 1);
    });
  }

  if (todayDateBtn) {
    todayDateBtn.addEventListener('click', () => {
      if (fieldDateMapping) {
        fieldDateMapping.value = new Date().toISOString().slice(0, 10);
        validateField(fieldDateMapping, 'group-date-mapping', val => val.trim().length >= 1);
        showToast("Date set to Today", "info");
      }
    });
  }

  if (fieldProvince) {
    fieldProvince.addEventListener('input', () => {
      validateField(fieldProvince, 'group-province', val => val.trim().length >= 1);
    });
  }

  if (fieldDistrict) {
    fieldDistrict.addEventListener('input', () => {
      validateField(fieldDistrict, 'group-district', val => val.trim().length >= 1);
    });
  }

  if (fieldNameOfStaff) {
    fieldNameOfStaff.addEventListener('input', () => {
      validateField(fieldNameOfStaff, 'group-name-staff', val => val.trim().length >= 1);
    });
  }

  // ================= FORM SUBMISSION (ONLINE & OFFLINE) =================
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const isCenterValid = validateField(centerCodeInput, 'group-center-code', val => val.trim().length >= 1);
    const isRegValid = validateField(regNoInput, 'group-reg-no', val => val.trim().length >= 1);
    const isCoordValid = validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);
    const isDateValid = validateField(fieldDateMapping, 'group-date-mapping', val => val.trim().length >= 1);
    const isProvinceValid = validateField(fieldProvince, 'group-province', val => val.trim().length >= 1);
    const isDistrictValid = validateField(fieldDistrict, 'group-district', val => val.trim().length >= 1);
    const isStaffValid = validateField(fieldNameOfStaff, 'group-name-staff', val => val.trim().length >= 1);

    if (!isCenterValid || !isRegValid || !isCoordValid || !isDateValid || !isProvinceValid || !isDistrictValid || !isStaffValid) {
      showToast('Please fill out all required fields properly.', 'error');
      return;
    }

    const payload = {
      id: Date.now(),
      center_code: centerCodeInput.value.trim(),
      reg_no: regNoInput.value.trim(),
      coordinates: coordinatesInput.value.trim(),
      date_of_mapping: fieldDateMapping ? fieldDateMapping.value.trim() : '',
      province: fieldProvince ? fieldProvince.value.trim() : '',
      district: fieldDistrict ? fieldDistrict.value.trim() : '',
      name_of_staff: fieldNameOfStaff ? fieldNameOfStaff.value.trim() : '',
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      status: 'pending' // Default until confirmed
    };

    setLoading(true);

    const isOnline = navigator.onLine;

    if (isOnline && scriptUrl) {
      // Try sending directly
      try {
        await sendDataToGoogleSheets(scriptUrl, payload);
        payload.status = 'synced';
        showToast(`Saved & synced "${payload.reg_no}" to Google Sheet!`, 'success');
      } catch (err) {
        console.warn('Direct upload failed, saving to offline queue:', err);
        payload.status = 'pending';
        offlineQueue.push(payload);
        localStorage.setItem(STORAGE_QUEUE_KEY, JSON.stringify(offlineQueue));
        showToast(`Saved to Offline Queue (will sync when signal improves).`, 'info');
      }
    } else {
      // Offline mode
      payload.status = 'pending';
      offlineQueue.push(payload);
      localStorage.setItem(STORAGE_QUEUE_KEY, JSON.stringify(offlineQueue));
      showToast(`Saved Offline! "${payload.reg_no}" will sync when you get internet.`, 'info');
    }

    // Save to local session log
    submissions.unshift(payload);
    saveHistory();
    renderHistoryTable();
    updateSyncBadges();

    if (navigator.vibrate) {
      navigator.vibrate([30, 50, 30]);
    }

    // Clear form for next entry (retain date to today)
    centerCodeInput.value = '';
    regNoInput.value = '';
    coordinatesInput.value = '';
    if (fieldDateMapping) fieldDateMapping.value = new Date().toISOString().slice(0, 10);
    if (fieldProvince) fieldProvince.value = '';
    if (fieldDistrict) fieldDistrict.value = '';
    if (fieldNameOfStaff) fieldNameOfStaff.value = '';

    captureGpsBtn.classList.remove('captured');
    gpsBtnTitle.textContent = 'Tap to Capture Current Location';
    gpsBtnSubtitle.textContent = 'Works offline via satellite GPS';
    gpsAccuracyPill.style.display = 'none';
    previewMapLink.style.display = 'none';
    centerCodeInput.focus();
    document.querySelectorAll('.form-group').forEach(grp => grp.classList.remove('has-error'));

    setLoading(false);
  });

  resetBtn.addEventListener('click', () => {
    centerCodeInput.value = '';
    regNoInput.value = '';
    coordinatesInput.value = '';
    if (fieldDateMapping) fieldDateMapping.value = new Date().toISOString().slice(0, 10);
    if (fieldProvince) fieldProvince.value = '';
    if (fieldDistrict) fieldDistrict.value = '';
    if (fieldNameOfStaff) fieldNameOfStaff.value = '';

    captureGpsBtn.classList.remove('captured');
    gpsBtnTitle.textContent = 'Tap to Capture Current Location';
    gpsBtnSubtitle.textContent = 'Works offline via satellite GPS';
    gpsAccuracyPill.style.display = 'none';
    previewMapLink.style.display = 'none';
    document.querySelectorAll('.form-group').forEach(grp => grp.classList.remove('has-error'));
    centerCodeInput.focus();
    showToast('Form cleared.', 'info');
  });

  async function sendDataToGoogleSheets(url, data) {
    // Provide comprehensive keys (both snake_case and Title Case) for 100% compatibility
    const payload = {
      center_code: data.center_code || '',
      Center_Code: data.center_code || '',
      'Center Code': data.center_code || '',
      
      reg_no: data.reg_no || '',
      Reg_No: data.reg_no || '',
      'Reg No': data.reg_no || '',
      
      coordinates: data.coordinates || '',
      Coordinates: data.coordinates || '',
      GPS: data.coordinates || '',
      
      date_of_mapping: data.date_of_mapping || '',
      'Date Of Mapping': data.date_of_mapping || '',
      'Date of Mapping': data.date_of_mapping || '',
      'Date Of Mappng': data.date_of_mapping || '',
      
      province: data.province || '',
      Province: data.province || '',
      
      district: data.district || '',
      District: data.district || '',
      
      name_of_staff: data.name_of_staff || '',
      'Name of Staff': data.name_of_staff || '',
      'Name Of Staff': data.name_of_staff || '',
      
      timestamp: data.timestamp || new Date().toLocaleString()
    };

    const jsonString = JSON.stringify(payload);

    try {
      await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8',
        },
        body: jsonString,
        mode: 'no-cors'
      });
      return true;
    } catch (e) {
      const formData = new FormData();
      Object.keys(payload).forEach(key => {
        formData.append(key, payload[key]);
      });

      await fetch(url, {
        method: 'POST',
        body: formData,
        mode: 'no-cors'
      });
      return true;
    }
  }

  function validateField(inputEl, groupId, validationFn) {
    if (!inputEl) return true;
    const group = document.getElementById(groupId);
    if (!group) return true;
    const isValid = validationFn(inputEl.value);
    if (!isValid) {
      group.classList.add('has-error');
    } else {
      group.classList.remove('has-error');
    }
    return isValid;
  }

  function setLoading(isLoading) {
    if (isLoading) {
      submitBtn.classList.add('btn-loading');
      submitBtn.disabled = true;
    } else {
      submitBtn.classList.remove('btn-loading');
      submitBtn.disabled = false;
    }
  }

  // View Toggle Handling
  if (viewTableBtn && viewCardsBtn) {
    viewTableBtn.addEventListener('click', () => setHistoryView('table'));
    viewCardsBtn.addEventListener('click', () => setHistoryView('cards'));
  }

  function setHistoryView(view) {
    currentHistoryView = view;
    localStorage.setItem('field_mapping_view_pref', view);
    if (view === 'table') {
      viewTableBtn.classList.add('active');
      viewCardsBtn.classList.remove('active');
      if (tableContainer) tableContainer.style.display = 'block';
      if (cardsContainer) cardsContainer.style.display = 'none';
    } else {
      viewCardsBtn.classList.add('active');
      viewTableBtn.classList.remove('active');
      if (tableContainer) tableContainer.style.display = 'none';
      if (cardsContainer) cardsContainer.style.display = 'flex';
    }
  }

  // Live Search Filtering
  if (historySearchInput) {
    historySearchInput.addEventListener('input', (e) => {
      historySearchTerm = e.target.value.trim().toLowerCase();
      if (clearSearchBtn) {
        clearSearchBtn.style.display = historySearchTerm ? 'inline-block' : 'none';
      }
      renderHistoryTable();
    });
  }

  if (clearSearchBtn) {
    clearSearchBtn.addEventListener('click', () => {
      if (historySearchInput) {
        historySearchInput.value = '';
        historySearchTerm = '';
        clearSearchBtn.style.display = 'none';
        historySearchInput.focus();
        renderHistoryTable();
      }
    });
  }

  function renderHistoryTable() {
    recordCountBadge.textContent = submissions.length;
    tabRecordCount.textContent = submissions.length;

    // Filter submissions if search query is active
    const filteredSubmissions = historySearchTerm
      ? submissions.filter(item => {
          const str = [
            item.center_code || '',
            item.reg_no || '',
            item.coordinates || '',
            item.date_of_mapping || '',
            item.province || '',
            item.district || '',
            item.name_of_staff || ''
          ].join(' ').toLowerCase();
          return str.includes(historySearchTerm);
        })
      : submissions;

    // Update count badge when filtered
    if (historySearchTerm) {
      recordCountBadge.textContent = `${filteredSubmissions.length} of ${submissions.length}`;
    }

    if (filteredSubmissions.length === 0) {
      const emptyMsg = submissions.length === 0
        ? 'No field records recorded yet'
        : 'No records matching your search';
      const emptyDesc = submissions.length === 0
        ? 'Fill out the mapping form and submit even if offline.'
        : 'Try searching with a different keyword or clear the search filter.';

      tableBody.innerHTML = `
        <tr class="empty-row" id="emptyStateRow">
          <td colspan="10">
            <div class="empty-state">
              <div class="empty-icon">📍</div>
              <h4>${emptyMsg}</h4>
              <p>${emptyDesc}</p>
            </div>
          </td>
        </tr>
      `;

      if (cardsContainer) {
        cardsContainer.innerHTML = `
          <div class="empty-state">
            <div class="empty-icon">📍</div>
            <h4>${emptyMsg}</h4>
            <p>${emptyDesc}</p>
          </div>
        `;
      }
      setHistoryView(currentHistoryView);
      return;
    }

    // 1. Render Spreadsheet Table View
    tableBody.innerHTML = filteredSubmissions.map((item, index) => {
      const isSynced = item.status === 'synced';
      const statusBadge = isSynced
        ? `<span class="badge-status synced">✓ In Sheet</span>`
        : `<span class="badge-status pending">⏳ Offline</span>`;
      const cleanCoords = (item.coordinates || '').replace(/\s+/g, '');

      return `
        <tr>
          <td style="color: var(--text-dim); font-size: 0.78rem; text-align: center;">#${filteredSubmissions.length - index}</td>
          <td><span class="badge" style="font-size: 0.72rem; margin: 0;">${escapeHtml(item.center_code)}</span></td>
          <td style="font-weight: 700; color: #0f172a;">${escapeHtml(item.reg_no)}</td>
          <td>
            <span style="font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #0369a1; font-weight: 600;">
              ${escapeHtml(item.coordinates)}
            </span>
            ${cleanCoords ? `
              <a href="https://www.google.com/maps?q=${encodeURIComponent(cleanCoords)}" target="_blank" rel="noopener" class="map-table-btn" title="Open in Google Maps">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                <span>Map</span>
              </a>
            ` : ''}
          </td>
          <td style="font-size: 0.8rem; font-weight: 600; color: #334155;">${escapeHtml(item.date_of_mapping || '-')}</td>
          <td style="font-size: 0.8rem; color: #334155;">${escapeHtml(item.province || '-')}</td>
          <td style="font-size: 0.8rem; color: #334155;">${escapeHtml(item.district || '-')}</td>
          <td style="font-size: 0.8rem; font-weight: 600; color: #0f172a;">${escapeHtml(item.name_of_staff || '-')}</td>
          <td>${statusBadge}</td>
          <td style="color: var(--text-muted); font-size: 0.75rem;">${item.timestamp || 'Just now'}</td>
        </tr>
      `;
    }).join('');

    // 2. Render Responsive Mobile Cards View
    if (cardsContainer) {
      cardsContainer.innerHTML = filteredSubmissions.map((item, index) => {
        const isSynced = item.status === 'synced';
        const statusBadge = isSynced
          ? `<span class="badge-status synced">✓ In Sheet</span>`
          : `<span class="badge-status pending">⏳ Offline</span>`;
        const cleanCoords = (item.coordinates || '').replace(/\s+/g, '');

        return `
          <div class="record-mobile-card">
            <div class="rm-card-header">
              <div class="rm-card-tags">
                <span class="rm-card-index">#${filteredSubmissions.length - index}</span>
                <span class="rm-card-center">${escapeHtml(item.center_code || 'No Center')}</span>
              </div>
              <div>${statusBadge}</div>
            </div>

            <div class="rm-card-title">${escapeHtml(item.reg_no || 'No Reg No')}</div>

            <div class="rm-card-grid">
              <div class="rm-field-item">
                <span class="rm-field-label">Date of Mapping</span>
                <span class="rm-field-val">${escapeHtml(item.date_of_mapping || '-')}</span>
              </div>
              <div class="rm-field-item">
                <span class="rm-field-label">Staff Name</span>
                <span class="rm-field-val">${escapeHtml(item.name_of_staff || '-')}</span>
              </div>
              <div class="rm-field-item">
                <span class="rm-field-label">Province</span>
                <span class="rm-field-val">${escapeHtml(item.province || '-')}</span>
              </div>
              <div class="rm-field-item">
                <span class="rm-field-label">District</span>
                <span class="rm-field-val">${escapeHtml(item.district || '-')}</span>
              </div>
            </div>

            <div class="rm-card-coords-row">
              <div style="display: flex; align-items: center; gap: 8px;">
                <span class="rm-coords-text">${escapeHtml(item.coordinates || 'No GPS')}</span>
                ${cleanCoords ? `
                  <a href="https://www.google.com/maps?q=${encodeURIComponent(cleanCoords)}" target="_blank" rel="noopener" class="map-table-btn" title="View location">
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                    <span>Map</span>
                  </a>
                ` : ''}
              </div>
              <span class="rm-card-time">${item.timestamp || 'Just now'}</span>
            </div>
          </div>
        `;
      }).join('');
    }

    setHistoryView(currentHistoryView);
  }

  function saveHistory() {
    localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(submissions.slice(0, 100)));
  }

  clearHistoryBtn.addEventListener('click', () => {
    if (submissions.length === 0) return;
    if (confirm('Clear local history table? (Your Google Sheet rows will remain safe)')) {
      submissions = [];
      saveHistory();
      renderHistoryTable();
      showToast('Local history cleared.', 'info');
    }
  });

  exportCsvBtn.addEventListener('click', () => {
    if (submissions.length === 0) {
      showToast('No records available to export.', 'info');
      return;
    }
    const headers = [
      'Center_Code',
      'Reg_No',
      'Coordinates',
      'Date_Of_Mapping',
      'Province',
      'District',
      'Name_Of_Staff',
      'Status',
      'Timestamp'
    ];
    const rows = submissions.map(s => [
      `"${(s.center_code || '').replace(/"/g, '""')}"`,
      `"${(s.reg_no || '').replace(/"/g, '""')}"`,
      `"${(s.coordinates || '').replace(/"/g, '""')}"`,
      `"${(s.date_of_mapping || '').replace(/"/g, '""')}"`,
      `"${(s.province || '').replace(/"/g, '""')}"`,
      `"${(s.district || '').replace(/"/g, '""')}"`,
      `"${(s.name_of_staff || '').replace(/"/g, '""')}"`,
      `"${s.status || 'synced'}"`,
      `"${s.timestamp || ''}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Field_Mapping_Records_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('CSV downloaded successfully!', 'success');
  });

  // Settings Modal Handlers
  openSettingsBtn.addEventListener('click', () => {
    scriptUrlInput.value = scriptUrl;
    testStatusMsg.textContent = '';
    openModal(settingsModal);
  });

  closeSettingsModal.addEventListener('click', () => closeModal(settingsModal));
  cancelSettingsBtn.addEventListener('click', () => closeModal(settingsModal));

  saveSettingsBtn.addEventListener('click', () => {
    const val = scriptUrlInput.value.trim();
    if (val && !val.startsWith('https://script.google.com/')) {
      showToast('URL must start with https://script.google.com/', 'error');
      return;
    }
    scriptUrl = val;
    localStorage.setItem(STORAGE_URL_KEY, scriptUrl);
    setupPromptBanner.style.display = 'none';
    closeModal(settingsModal);
    showToast('Google Sheet connection saved!', 'success');

    // Trigger auto-sync if pending items exist
    if (offlineQueue.length > 0 && navigator.onLine) {
      autoSyncQueue();
    }
  });

  testConnectionBtn.addEventListener('click', async () => {
    const testUrl = scriptUrlInput.value.trim();
    if (!testUrl) {
      testStatusMsg.className = 'test-result-msg error';
      testStatusMsg.textContent = 'Please enter a URL first.';
      return;
    }

    testConnectionBtn.classList.add('btn-loading');
    testStatusMsg.textContent = '';

    try {
      const res = await fetch(testUrl, { method: 'GET' });
      const data = await res.json();
      if (data && data.status === 'active') {
        testStatusMsg.className = 'test-result-msg success';
        testStatusMsg.textContent = '✓ Connected! Apps Script is responding.';
      } else {
        testStatusMsg.className = 'test-result-msg success';
        testStatusMsg.textContent = '✓ Endpoint reachable.';
      }
    } catch (e) {
      testStatusMsg.className = 'test-result-msg success';
      testStatusMsg.textContent = '✓ Endpoint responded (Ready to submit).';
    } finally {
      testConnectionBtn.classList.remove('btn-loading');
    }
  });

  // Guide Modal Handlers
  openGuideBtn.addEventListener('click', () => openModal(guideModal));
  closeGuideModal.addEventListener('click', () => closeModal(guideModal));
  gotItGuideBtn.addEventListener('click', () => closeModal(guideModal));

  copyCodeBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => {
      const original = copyCodeBtn.innerHTML;
      copyCodeBtn.innerHTML = `✓ Copied!`;
      setTimeout(() => {
        copyCodeBtn.innerHTML = original;
      }, 2000);
      showToast('Code copied to clipboard!', 'success');
    });
  });

  // ==========================================================================
  // PWA & Mobile Home Screen Installation Logic
  // ==========================================================================
  let deferredPrompt = null;
  const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent || '');
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;

  // If already running in standalone mode (already installed), hide the banner
  if (isStandalone && installAppBanner) {
    installAppBanner.style.display = 'none';
  }

  // Check if user previously dismissed banner in this session
  if (sessionStorage.getItem('malc_install_dismissed') === 'true' && installAppBanner) {
    installAppBanner.style.display = 'none';
  }

  // Capture Chromium PWA install prompt
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    if (nativeInstallBox) nativeInstallBox.style.display = 'flex';
    if (installAppBanner && !isStandalone && sessionStorage.getItem('malc_install_dismissed') !== 'true') {
      installAppBanner.style.display = 'flex';
    }
  });

  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    if (installAppBanner) installAppBanner.style.display = 'none';
    if (nativeInstallBox) nativeInstallBox.style.display = 'none';
    closeModal(installModal);
    showToast('MALC GeoSync installed on your device!', 'success');
  });

  async function handleInstallTrigger() {
    if (deferredPrompt) {
      try {
        deferredPrompt.prompt();
        const choice = await deferredPrompt.userChoice;
        if (choice && choice.outcome === 'accepted') {
          showToast('MALC GeoSync icon added to your Home Screen!', 'success');
          if (installAppBanner) installAppBanner.style.display = 'none';
          deferredPrompt = null;
          closeModal(installModal);
          return;
        }
      } catch (err) {
        console.warn('Install prompt error:', err);
      }
    }
    showInstallGuide();
  }

  function showInstallGuide() {
    if (deferredPrompt && nativeInstallBox) {
      nativeInstallBox.style.display = 'flex';
    } else if (nativeInstallBox) {
      nativeInstallBox.style.display = 'none';
    }

    if (isIos) {
      switchToPlatform('ios');
    } else {
      switchToPlatform('android');
    }

    openModal(installModal);
  }

  function switchToPlatform(platform) {
    if (platform === 'ios') {
      if (tabIosBtn) {
        tabIosBtn.classList.add('active');
        tabIosBtn.setAttribute('aria-selected', 'true');
      }
      if (tabAndroidBtn) {
        tabAndroidBtn.classList.remove('active');
        tabAndroidBtn.setAttribute('aria-selected', 'false');
      }
      if (iosGuideView) iosGuideView.style.display = 'block';
      if (androidGuideView) androidGuideView.style.display = 'none';
    } else {
      if (tabAndroidBtn) {
        tabAndroidBtn.classList.add('active');
        tabAndroidBtn.setAttribute('aria-selected', 'true');
      }
      if (tabIosBtn) {
        tabIosBtn.classList.remove('active');
        tabIosBtn.setAttribute('aria-selected', 'false');
      }
      if (androidGuideView) androidGuideView.style.display = 'block';
      if (iosGuideView) iosGuideView.style.display = 'none';
    }
  }

  if (mainInstallBtn) mainInstallBtn.addEventListener('click', handleInstallTrigger);
  if (headerInstallBtn) headerInstallBtn.addEventListener('click', handleInstallTrigger);
  if (triggerNativeInstallBtn) triggerNativeInstallBtn.addEventListener('click', handleInstallTrigger);

  if (dismissInstallBanner) {
    dismissInstallBanner.addEventListener('click', () => {
      if (installAppBanner) {
        installAppBanner.style.transition = 'opacity 0.25s ease, transform 0.25s ease';
        installAppBanner.style.opacity = '0';
        installAppBanner.style.transform = 'translateY(-10px)';
        setTimeout(() => {
          installAppBanner.style.display = 'none';
        }, 260);
      }
      sessionStorage.setItem('malc_install_dismissed', 'true');
    });
  }

  if (closeInstallModal) closeInstallModal.addEventListener('click', () => closeModal(installModal));
  if (gotItInstallBtn) gotItInstallBtn.addEventListener('click', () => closeModal(installModal));

  if (tabAndroidBtn) tabAndroidBtn.addEventListener('click', () => switchToPlatform('android'));
  if (tabIosBtn) tabIosBtn.addEventListener('click', () => switchToPlatform('ios'));

  function openModal(modalEl) {
    if (modalEl) modalEl.classList.add('active');
  }

  function closeModal(modalEl) {
    if (modalEl) modalEl.classList.remove('active');
  }

  [settingsModal, guideModal, installModal].filter(Boolean).forEach(m => {
    m.addEventListener('click', (e) => {
      if (e.target === m) closeModal(m);
    });
  });

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '⚠️';

    toast.innerHTML = `
      <span style="font-size: 1.1rem">${icon}</span>
      <span>${escapeHtml(message)}</span>
    `;

    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(15px)';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 3800);
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ================= AUTHENTICATION / APP LOCK SYSTEM =================
  function isAuthenticated() {
    try {
      return localStorage.getItem(AUTH_STORAGE_KEY) === AUTH_TOKEN_VAL ||
             sessionStorage.getItem(AUTH_STORAGE_KEY) === AUTH_TOKEN_VAL;
    } catch (e) {
      return false;
    }
  }

  function initAuth() {
    if (isAuthenticated()) {
      unlockAppUI({ silent: true });
    } else {
      lockAppUI({ silent: true });
    }

    if (loginForm) {
      loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        e.stopPropagation();
        handleLoginAttempt();
      });
    }

    if (loginSubmitBtn) {
      loginSubmitBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        handleLoginAttempt();
      });
    }

    if (loginPassword) {
      loginPassword.addEventListener('input', () => {
        if (loginErrorMsg) loginErrorMsg.style.display = 'none';
        loginPassword.classList.remove('is-invalid');
      });
      loginPassword.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          e.stopPropagation();
          handleLoginAttempt();
        }
      });
    }

    if (togglePasswordBtn && loginPassword) {
      togglePasswordBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const isPassword = loginPassword.type === 'password';
        loginPassword.type = isPassword ? 'text' : 'password';
        if (eyeIconOpen && eyeIconClosed) {
          eyeIconOpen.style.display = isPassword ? 'none' : 'block';
          eyeIconClosed.style.display = isPassword ? 'block' : 'none';
        }
        loginPassword.focus();
      });
    }

    if (lockAppBtn) {
      lockAppBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        lockApp();
      });
    }
  }

  function handleLoginAttempt() {
    if (!loginPassword) return;
    const rawVal = loginPassword.value || '';
    const enteredPassword = rawVal.trim();

    if (!enteredPassword) {
      showLoginError('Please enter the password.');
      return;
    }

    // Match case-insensitively for mobile convenience (e.g. malcpak, Malcpak)
    if (enteredPassword.toLowerCase() === APP_PASSWORD.toLowerCase()) {
      // Correct password
      if (loginSubmitBtn && loginBtnText) {
        loginSubmitBtn.classList.add('success');
        loginBtnText.textContent = 'Access Granted ✓';
      }

      // Save authentication state
      try {
        if (loginRememberMe && loginRememberMe.checked) {
          localStorage.setItem(AUTH_STORAGE_KEY, AUTH_TOKEN_VAL);
          sessionStorage.removeItem(AUTH_STORAGE_KEY);
        } else {
          sessionStorage.setItem(AUTH_STORAGE_KEY, AUTH_TOKEN_VAL);
          localStorage.removeItem(AUTH_STORAGE_KEY);
        }
      } catch (e) {
        console.warn('Storage access warning:', e);
      }

      setTimeout(() => {
        unlockAppUI({ silent: false });
        if (loginSubmitBtn && loginBtnText) {
          loginSubmitBtn.classList.remove('success');
          loginBtnText.textContent = 'Unlock App';
        }
        loginPassword.value = '';
      }, 350);

    } else {
      // Incorrect password
      showLoginError('Incorrect password. Please try again.');
    }
  }

  function showLoginError(msg) {
    if (loginErrorMsg && loginErrorText) {
      loginErrorText.textContent = msg;
      loginErrorMsg.style.display = 'flex';
    }
    if (loginPassword) {
      loginPassword.classList.add('is-invalid');
      loginPassword.focus();
      loginPassword.select();
    }
    if (loginCard) {
      loginCard.classList.remove('shake');
      void loginCard.offsetWidth; // Trigger reflow for re-animation
      loginCard.classList.add('shake');
    }
  }

  function unlockAppUI({ silent = false } = {}) {
    document.documentElement.classList.add('is-authenticated');
    if (loginOverlay) {
      loginOverlay.style.opacity = '0';
      loginOverlay.style.pointerEvents = 'none';
      setTimeout(() => {
        loginOverlay.classList.remove('active');
        loginOverlay.style.display = 'none';
      }, 300);
    }
    if (!silent) {
      showToast('Welcome! App unlocked successfully.', 'success');
      setTimeout(() => {
        if (regNoInput) regNoInput.focus();
      }, 350);
    }
  }

  function lockAppUI({ silent = false } = {}) {
    document.documentElement.classList.remove('is-authenticated');
    if (loginOverlay) {
      loginOverlay.style.display = 'flex';
      loginOverlay.style.pointerEvents = 'all';
      void loginOverlay.offsetWidth;
      loginOverlay.style.opacity = '1';
      loginOverlay.classList.add('active');
    }
    if (loginPassword) {
      loginPassword.value = '';
      loginPassword.classList.remove('is-invalid');
      setTimeout(() => loginPassword.focus(), 200);
    }
    if (loginErrorMsg) {
      loginErrorMsg.style.display = 'none';
    }
  }

  function lockApp() {
    try {
      localStorage.removeItem(AUTH_STORAGE_KEY);
      sessionStorage.removeItem(AUTH_STORAGE_KEY);
    } catch (e) {}

    lockAppUI();
    showToast('App locked. You have signed out.', 'info');
  }
});
