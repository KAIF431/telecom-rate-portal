// content_portal.js - Injected into Telecom Rate Portal
// Listens to "Auto Upload to RouteTrust" and coordinates background execution

(function () {
  'use strict';

  console.log('[RouteTrust Auto-Pilot] Portal listener active.');

  // Listen to inline onclick postMessage from button
  window.addEventListener('message', function (event) {
    if (event.data && event.data.action === 'PORTAL_BTN_CLICKED') {
      const { subName, fileName, b64 } = event.data;
      console.log('[RouteTrust Auto-Pilot] Button clicked for sub:', subName);

      try {
        chrome.runtime.sendMessage({
          action: 'START_AUTO_UPLOAD',
          subName: subName,
          fileName: fileName,
          fileBase64: b64
        }, (response) => {
          console.log('[RouteTrust Auto-Pilot] Background response:', response);
        });
      } catch (err) {
        console.error('[RouteTrust Auto-Pilot] Extension error:', err);
      }
    }
  });

  // Also bind direct clicks if any
  function bindClicks() {
    const buttons = document.querySelectorAll('.rt-btn-autoupload');
    buttons.forEach(btn => {
      if (btn.dataset.bound) return;
      btn.dataset.bound = 'true';
      btn.addEventListener('click', function () {
        const subName = btn.dataset.subname || 'KAMAHIRAF (VENDOR CLI)';
        const fileName = btn.dataset.filename || 'rates.xlsx';
        const fileBase64 = btn.dataset.b64 || '';
        try {
          chrome.runtime.sendMessage({
            action: 'START_AUTO_UPLOAD',
            subName: subName,
            fileName: fileName,
            fileBase64: fileBase64
          });
        } catch (e) {}
      });
    });
  }
  setInterval(bindClicks, 1000);

  // Listen for Live Stage Updates & Success from Background
  chrome.runtime.onMessage.addListener((msg) => {
    console.log('[RouteTrust Auto-Pilot] Portal stage update:', msg);

    const statusBoxes = document.querySelectorAll('.rt-card-status');
    const autoBtns = document.querySelectorAll('.rt-btn-autoupload');

    if (msg.action === 'PORTAL_UPDATE_STAGE') {
      statusBoxes.forEach(sb => {
        sb.style.display = 'block';
        sb.style.background = 'rgba(245, 158, 11, 0.15)';
        sb.style.color = '#fbbf24';
        sb.style.border = '1px solid rgba(245, 158, 11, 0.3)';
        sb.innerHTML = `⏳ ${msg.message}`;
      });
      autoBtns.forEach(btn => {
        btn.innerHTML = `⏳ ${msg.message}`;
      });
    }

    if (msg.action === 'PORTAL_SUCCESS') {
      statusBoxes.forEach(sb => {
        sb.style.display = 'block';
        sb.style.background = 'rgba(34, 197, 94, 0.15)';
        sb.style.color = '#4ade80';
        sb.style.border = '1px solid rgba(34, 197, 94, 0.3)';
        sb.innerHTML = `🎉 <b>SUCCESS!</b> Rates Approved & Loaded into RouteTrust (<b>${msg.subName}</b>)!`;
      });
      autoBtns.forEach(btn => {
        btn.disabled = false;
        btn.style.background = 'linear-gradient(135deg, #16a34a 0%, #15803d 100%)';
        btn.style.borderColor = '#22c55e';
        btn.innerHTML = `✅ Rates Approved & Loaded!`;
      });
    }
  });
})();
