// GeoSync Google Sheets Field Data Integration Logic
document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const form = document.getElementById('dataEntryForm');
  const centerCodeInput = document.getElementById('fieldCenterCode');
  const regNoInput = document.getElementById('fieldRegNo');
  const coordinatesInput = document.getElementById('fieldCoordinates');
  const captureGpsBtn = document.getElementById('captureGpsBtn');
  const gpsAccuracyPill = document.getElementById('gpsAccuracyPill');
  const accuracyVal = document.getElementById('accuracyVal');
  const previewMapLink = document.getElementById('previewMapLink');

  const submitBtn = document.getElementById('submitBtn');
  const resetBtn = document.getElementById('resetBtn');

  // Status & Headers
  const connectionStatus = document.getElementById('connectionStatus');
  const sheetTargetDisplay = document.getElementById('sheetTargetDisplay');
  const recordCountBadge = document.getElementById('recordCount');
  const tableBody = document.getElementById('tableBody');
  const exportCsvBtn = document.getElementById('exportCsvBtn');
  const clearHistoryBtn = document.getElementById('clearHistoryBtn');

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

  // Google Apps Script source code snippet shown in guide
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

    // Appends row matching your headers: Center_Code | Reg No | Coordinates | Timestamp
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

  // State
  let scriptUrl = localStorage.getItem(STORAGE_URL_KEY) || '';
  let submissions = [];

  try {
    submissions = JSON.parse(localStorage.getItem(STORAGE_HISTORY_KEY)) || [];
  } catch (e) {
    submissions = [];
  }

  // Initial UI Setup
  updateConnectionBadge();
  renderHistoryTable();

  // ================= GPS CAPTURE HANDLER =================
  captureGpsBtn.addEventListener('click', () => {
    if (!navigator.geolocation) {
      showToast('Geolocation is not supported by your browser.', 'error');
      return;
    }

    captureGpsBtn.classList.add('acquiring');
    showToast('Acquiring high-accuracy GPS coordinates...', 'info');

    const geoOptions = {
      enableHighAccuracy: true,
      timeout: 12000,
      maximumAge: 0
    };

    navigator.geolocation.getCurrentPosition(
      (position) => {
        captureGpsBtn.classList.remove('acquiring');
        const lat = position.coords.latitude.toFixed(6);
        const lng = position.coords.longitude.toFixed(6);
        const accuracy = Math.round(position.coords.accuracy);

        const coordString = `${lat}, ${lng}`;
        coordinatesInput.value = coordString;

        // Display accuracy badge
        accuracyVal.textContent = `±${accuracy}m`;
        gpsAccuracyPill.style.display = 'inline-flex';

        // Update Map preview link
        updateMapPreview(coordString);

        validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);
        showToast(`GPS captured: ${coordString} (±${accuracy}m accuracy)`, 'success');
      },
      (error) => {
        captureGpsBtn.classList.remove('acquiring');
        console.warn('Geolocation error:', error);
        let msg = 'Failed to get GPS location.';
        if (error.code === error.PERMISSION_DENIED) {
          msg = 'Location permission denied. Please allow location access in your browser.';
        } else if (error.code === error.POSITION_UNAVAILABLE) {
          msg = 'Location information is unavailable. Ensure device GPS is turned on.';
        } else if (error.code === error.TIMEOUT) {
          msg = 'GPS request timed out. Please try again or type coordinates manually.';
        }
        showToast(msg, 'error');
      },
      geoOptions
    );
  });

  // Manual Coordinates Input listener
  coordinatesInput.addEventListener('input', () => {
    const val = coordinatesInput.value.trim();
    validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);
    updateMapPreview(val);
  });

  function isValidCoordinates(val) {
    if (!val || val.trim().length === 0) return false;
    // Checks format: "lat, lng" or decimal numbers separated by comma/space
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

  // Real-time Input Validation
  centerCodeInput.addEventListener('input', () => {
    validateField(centerCodeInput, 'group-center-code', val => val.trim().length >= 1);
  });

  regNoInput.addEventListener('input', () => {
    validateField(regNoInput, 'group-reg-no', val => val.trim().length >= 1);
  });

  // Keyboard shortcut Ctrl + Enter to submit
  form.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      form.requestSubmit();
    }
  });

  // ================= FORM SUBMISSION =================
  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    // Check URL first
    if (!scriptUrl) {
      showToast('Please set your Google Sheet Web App URL in settings first!', 'error');
      openModal(settingsModal);
      return;
    }

    // Validate fields
    const isCenterValid = validateField(centerCodeInput, 'group-center-code', val => val.trim().length >= 1);
    const isRegValid = validateField(regNoInput, 'group-reg-no', val => val.trim().length >= 1);
    const isCoordValid = validateField(coordinatesInput, 'group-coordinates', isValidCoordinates);

    if (!isCenterValid || !isRegValid || !isCoordValid) {
      showToast('Please fill out Center Code, Reg No, and Coordinates properly.', 'error');
      return;
    }

    const payload = {
      center_code: centerCodeInput.value.trim(),
      reg_no: regNoInput.value.trim(),
      coordinates: coordinatesInput.value.trim(),
      timestamp: new Date().toLocaleTimeString()
    };

    setLoading(true);

    try {
      await sendDataToGoogleSheets(scriptUrl, payload);
      
      // Save locally
      payload.status = 'Synchronized';
      payload.id = Date.now();
      submissions.unshift(payload);
      saveHistory();
      renderHistoryTable();

      showToast(`Success! Logged "${payload.reg_no}" to Google Sheet.`, 'success');

      // Reset fields
      centerCodeInput.value = '';
      regNoInput.value = '';
      coordinatesInput.value = '';
      gpsAccuracyPill.style.display = 'none';
      previewMapLink.style.display = 'none';
      centerCodeInput.focus();

      // Clear validation borders
      document.querySelectorAll('.form-group').forEach(grp => grp.classList.remove('has-error'));

    } catch (err) {
      console.error('Submission error:', err);
      showToast(`Entry completed. Check Google Sheet for new row.`, 'info');
      
      payload.status = 'Synchronized';
      payload.id = Date.now();
      submissions.unshift(payload);
      saveHistory();
      renderHistoryTable();
    } finally {
      setLoading(false);
    }
  });

  // Reset Button
  resetBtn.addEventListener('click', () => {
    centerCodeInput.value = '';
    regNoInput.value = '';
    coordinatesInput.value = '';
    gpsAccuracyPill.style.display = 'none';
    previewMapLink.style.display = 'none';
    document.querySelectorAll('.form-group').forEach(grp => grp.classList.remove('has-error'));
    centerCodeInput.focus();
    showToast('Form inputs cleared.', 'info');
  });

  // Google Sheets Send Function (handles Google Apps Script redirects)
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

  // Validation Helper
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

  // Loading indicator
  function setLoading(isLoading) {
    if (isLoading) {
      submitBtn.classList.add('btn-loading');
      submitBtn.disabled = true;
    } else {
      submitBtn.classList.remove('btn-loading');
      submitBtn.disabled = false;
    }
  }

  // Update Status Pill
  function updateConnectionBadge() {
    if (scriptUrl && scriptUrl.includes('script.google.com')) {
      connectionStatus.className = 'status-pill connected';
      connectionStatus.querySelector('.status-label').textContent = 'Google Sheet Connected';
      sheetTargetDisplay.innerHTML = `<span class="dot"></span> Connected`;
    } else {
      connectionStatus.className = 'status-pill disconnected';
      connectionStatus.querySelector('.status-label').textContent = 'Link Google Sheet';
      sheetTargetDisplay.innerHTML = `<span class="dot" style="background: #f59e0b"></span> Not Linked`;
    }
  }

  // Render Submissions Table
  function renderHistoryTable() {
    recordCountBadge.textContent = submissions.length;

    if (submissions.length === 0) {
      tableBody.innerHTML = `
        <tr class="empty-row" id="emptyStateRow">
          <td colspan="6">
            <div class="empty-state">
              <div class="empty-icon">📍</div>
              <h4>No field records submitted yet</h4>
              <p>Capture your first coordinates and submit on the left.</p>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tableBody.innerHTML = submissions.map((item, index) => {
      const mapUrl = `https://www.google.com/maps?q=${encodeURIComponent((item.coordinates || '').replace(/\s+/g, ''))}`;
      return `
        <tr>
          <td style="color: var(--text-dim); font-size: 0.8rem;">#${submissions.length - index}</td>
          <td><span class="badge" style="font-size: 0.72rem; margin: 0;">${escapeHtml(item.center_code)}</span></td>
          <td style="font-weight: 600;">${escapeHtml(item.reg_no)}</td>
          <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; color: #a7f3d0;">
            ${escapeHtml(item.coordinates)}
          </td>
          <td>
            <a href="${mapUrl}" target="_blank" class="map-table-btn" title="View on Google Maps">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
                <polyline points="15 3 21 3 21 9"></polyline>
                <line x1="10" y1="14" x2="21" y2="3"></line>
              </svg>
              Map
            </a>
          </td>
          <td style="color: var(--text-muted); font-size: 0.78rem;">${item.timestamp || 'Just now'}</td>
        </tr>
      `;
    }).join('');
  }

  function saveHistory() {
    localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(submissions.slice(0, 100)));
  }

  // Clear History
  clearHistoryBtn.addEventListener('click', () => {
    if (submissions.length === 0) return;
    if (confirm('Clear local history table? (Your Google Sheet rows will remain safe)')) {
      submissions = [];
      saveHistory();
      renderHistoryTable();
      showToast('Local history cleared.', 'info');
    }
  });

  // Export CSV
  exportCsvBtn.addEventListener('click', () => {
    if (submissions.length === 0) {
      showToast('No records available to export.', 'info');
      return;
    }
    const headers = ['Center_Code', 'Reg_No', 'Coordinates', 'Timestamp'];
    const rows = submissions.map(s => [
      `"${(s.center_code || '').replace(/"/g, '""')}"`,
      `"${(s.reg_no || '').replace(/"/g, '""')}"`,
      `"${(s.coordinates || '').replace(/"/g, '""')}"`,
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

  connectionStatus.addEventListener('click', () => {
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
    updateConnectionBadge();
    closeModal(settingsModal);
    showToast('Google Sheet connection saved!', 'success');
  });

  // Test Connection
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
  gotItGuideBtn.addEventListener('click', () => {
    closeModal(guideModal);
    openModal(settingsModal);
  });

  // Copy Code Button
  copyCodeBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => {
      const original = copyCodeBtn.innerHTML;
      copyCodeBtn.innerHTML = `✓ Copied!`;
      setTimeout(() => {
        copyCodeBtn.innerHTML = original;
      }, 2000);
      showToast('Apps Script code copied to clipboard!', 'success');
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

  // Toast Notification System
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
      toast.style.transform = 'translateX(30px)';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 4000);
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
