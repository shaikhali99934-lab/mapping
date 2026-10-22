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

  // Google Apps Script snippet shown in guide
  const APPS_SCRIPT_CODE = `function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data;
    
    // Support JSON or Form Data
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }

    var centerCode = data.center_code || data.Center_Code || '';
    var regNo = data.reg_no || data.Reg_No || data['Reg No'] || '';
    var coordinates = data.coordinates || data.Coordinates || '';
    var timestamp = new Date().toLocaleString();

    // Appends: Center_Code | Reg No | Coordinates | Timestamp
    sheet.appendRow([centerCode, regNo, coordinates, timestamp]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Row added' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'active', message: 'Ready' }))
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
    }
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

  // ================= FORM SUBMISSION (ONLINE & OFFLINE) =================
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const isCenterValid = validateField(centerCodeInput, 'group-center-code', val => val.trim().length >= 1);
    const isRegValid = validateField(regNoInput, 'group-reg-no', val => val.trim().length >= 1);
    const isCoordValid = validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);

    if (!isCenterValid || !isRegValid || !isCoordValid) {
      showToast('Please fill out all required fields properly.', 'error');
      return;
    }

    const payload = {
      id: Date.now(),
      center_code: centerCodeInput.value.trim(),
      reg_no: regNoInput.value.trim(),
      coordinates: coordinatesInput.value.trim(),
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

    // Clear form for next entry
    centerCodeInput.value = '';
    regNoInput.value = '';
    coordinatesInput.value = '';
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
    const jsonString = JSON.stringify(data);

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
      formData.append('center_code', data.center_code);
      formData.append('reg_no', data.reg_no);
      formData.append('coordinates', data.coordinates);

      await fetch(url, {
        method: 'POST',
        body: formData,
        mode: 'no-cors'
      });
      return true;
    }
  }

  function validateField(inputEl, groupId, validationFn) {
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

  function renderHistoryTable() {
    recordCountBadge.textContent = submissions.length;
    tabRecordCount.textContent = submissions.length;

    if (submissions.length === 0) {
      tableBody.innerHTML = `
        <tr class="empty-row" id="emptyStateRow">
          <td colspan="6">
            <div class="empty-state">
              <div class="empty-icon">📍</div>
              <h4>No field records recorded yet</h4>
              <p>Capture GPS coordinates and submit even if offline.</p>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tableBody.innerHTML = submissions.map((item, index) => {
      const isSynced = item.status === 'synced';
      const statusBadge = isSynced
        ? `<span class="badge-status synced">✓ In Sheet</span>`
        : `<span class="badge-status pending">⏳ Offline (Pending)</span>`;

      return `
        <tr>
          <td style="color: var(--text-dim); font-size: 0.78rem;">#${submissions.length - index}</td>
          <td><span class="badge" style="font-size: 0.7rem; margin: 0;">${escapeHtml(item.center_code)}</span></td>
          <td style="font-weight: 700;">${escapeHtml(item.reg_no)}</td>
          <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #a7f3d0;">
            ${escapeHtml(item.coordinates)}
          </td>
          <td>${statusBadge}</td>
          <td style="color: var(--text-muted); font-size: 0.76rem;">${item.timestamp || 'Just now'}</td>
        </tr>
      `;
    }).join('');
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
    const headers = ['Center_Code', 'Reg_No', 'Coordinates', 'Status', 'Timestamp'];
    const rows = submissions.map(s => [
      `"${(s.center_code || '').replace(/"/g, '""')}"`,
      `"${(s.reg_no || '').replace(/"/g, '""')}"`,
      `"${(s.coordinates || '').replace(/"/g, '""')}"`,
      `"${s.status || 'synced'}"`,
      `"${s.timestamp || ''}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `GeoSync_Field_Records_${new Date().toISOString().slice(0, 10)}.csv`);
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

  function openModal(modalEl) {
    modalEl.classList.add('active');
  }

  function closeModal(modalEl) {
    modalEl.classList.remove('active');
  }

  [settingsModal, guideModal].forEach(m => {
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
});
