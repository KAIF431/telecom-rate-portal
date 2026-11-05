// content_routetrust.js - Automated Actions on mcm.routetrust.com (Background Tab)

(function () {
  'use strict';

  console.log('[RouteTrust Auto-Pilot] Content script running on:', window.location.href);

  function base64ToBlob(base64, type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    const binStr = atob(base64);
    const len = binStr.length;
    const arr = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      arr[i] = binStr.charCodeAt(i);
    }
    return new Blob([arr], { type: type });
  }

  // ==========================================
  // 1. UPLOAD PAGE (/rateimport/upload/id/*)
  // ==========================================
  if (window.location.href.includes('/rateimport/upload/id/')) {
    console.log('[RouteTrust Auto-Pilot] On upload page. Checking for pending file from portal...');

    function ensureAtoZ() {
      const radioInputs = document.querySelectorAll('input[type="radio"]');
      radioInputs.forEach(r => {
        const label = r.parentElement ? r.parentElement.textContent.trim().toLowerCase() : '';
        if (label.includes('a to z') || r.value.toLowerCase().includes('a to z')) {
          r.checked = true;
        }
      });
    }
    ensureAtoZ();

    chrome.runtime.sendMessage({ action: 'GET_RT_FILE' }, (response) => {
      if (response && response.hasData && response.fileBase64) {
        console.log('[RouteTrust Auto-Pilot] Received file data for:', response.fileName);

        const fileInput = document.querySelector('input[type="file"]');
        if (fileInput) {
          try {
            const blob = base64ToBlob(response.fileBase64);
            const file = new File([blob], response.fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            const dt = new DataTransfer();
            dt.items.add(file);
            fileInput.files = dt.files;
            fileInput.dispatchEvent(new Event('change', { bubbles: true }));

            ensureAtoZ();

            chrome.runtime.sendMessage({
              action: 'UPDATE_STAGE',
              stage: 'SUBMITTING',
              message: `Uploading <b>${response.fileName}</b> to RouteTrust...`
            });

            setTimeout(() => {
              const submitBtn = Array.from(document.querySelectorAll('button, input[type="submit"]')).find(el => {
                const t = (el.textContent || el.value || '').toLowerCase();
                return t.includes('submit');
              });

              if (submitBtn) {
                submitBtn.click();
              } else {
                const form = fileInput.closest('form');
                if (form) form.submit();
              }
            }, 1000);
          } catch (e) {
            console.error('[RouteTrust Auto-Pilot] Error attaching file:', e);
          }
        }
      }
    });
  }

  // ==========================================
  // 2. STATUS & APPROVAL PAGE (/rateimport/status/id/*)
  // ==========================================
  if (window.location.href.includes('/rateimport/status/id/')) {
    console.log('[RouteTrust Auto-Pilot] On status page. Monitoring analysis & buttons...');

    function checkAndClickActionButtons() {
      const clickables = Array.from(document.querySelectorAll('button, a, input[type="button"], input[type="submit"]'));

      // Check for "Approve & Load"
      const approveBtn = clickables.find(el => {
        const txt = (el.textContent || el.value || '').trim().toLowerCase();
        return txt.includes('approve & load') || txt.includes('approve and load') || txt.includes('approve');
      });

      if (approveBtn && !approveBtn.disabled && approveBtn.offsetParent !== null) {
        console.log('[RouteTrust Auto-Pilot] Found "Approve & Load"! Clicking...');
        chrome.runtime.sendMessage({
          action: 'UPDATE_STAGE',
          stage: 'APPROVING',
          message: 'Found "Approve & Load" — Approving and loading rates now...'
        });

        setTimeout(() => {
          approveBtn.click();
        }, 1200);
        return true;
      }

      // Check for "Use Vendor Date"
      const vendorDateBtn = clickables.find(el => {
        const txt = (el.textContent || el.value || '').trim().toLowerCase();
        return txt.includes('use vendor date') || txt.includes('vendor date');
      });

      if (vendorDateBtn && !vendorDateBtn.disabled && vendorDateBtn.offsetParent !== null) {
        console.log('[RouteTrust Auto-Pilot] Found "Use Vendor Date"! Clicking...');
        chrome.runtime.sendMessage({
          action: 'UPDATE_STAGE',
          stage: 'VENDOR_DATE',
          message: 'Clearing exceptions with "Use Vendor Date"...'
        });

        setTimeout(() => {
          vendorDateBtn.click();
        }, 1200);
        return true;
      }

      // Check if Completed OK
      const bodyText = document.body.innerText.toLowerCase();
      if (bodyText.includes('rates loaded') || (bodyText.includes('stage status') && bodyText.includes('completed ok'))) {
        console.log('[RouteTrust Auto-Pilot] Rates Loaded successfully!');
        const subIdMatch = window.location.href.match(/status\/id\/(\d+)/);
        const subId = subIdMatch ? subIdMatch[1] : '29168';

        chrome.runtime.sendMessage({
          action: 'RT_COMPLETED',
          subId: subId
        });
        return true;
      }

      return false;
    }

    // Check immediately and periodically
    const hasAction = checkAndClickActionButtons();
    if (!hasAction) {
      const intervalId = setInterval(() => {
        const acted = checkAndClickActionButtons();
        if (acted) clearInterval(intervalId);
      }, 2000);
    }
  }
})();
