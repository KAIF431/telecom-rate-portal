// content_portal.js - Injected into Telecom Rate Portal
// Adds "Auto Upload to RouteTrust" button and manages background execution

(function () {
  'use strict';

  console.log('[RouteTrust Auto-Pilot] Portal script initialized.');

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result.split(',')[1];
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  // Inject Auto Upload Button next to Download Excel button
  function injectAutoUploadButtons() {
    const downloadLinks = Array.from(document.querySelectorAll('a[download]'));

    for (const link of downloadLinks) {
      const downloadAttr = link.getAttribute('download') || '';
      const href = link.getAttribute('href') || '';

      // Only target cleaned .xlsx files
      if (!downloadAttr.endsWith('.xlsx')) continue;

      // Check if button already injected for this link
      const parentContainer = link.closest('[data-testid="column"]') || link.parentElement;
      if (!parentContainer || parentContainer.querySelector('.rt-injected-container')) continue;

      // Extract Subscription/Vendor Name from filename
      let subName = downloadAttr.split(' A to Z ')[0].trim();
      if (!subName || subName.length < 2) {
        subName = downloadAttr.split('.')[0].trim();
      }

      console.log('[RouteTrust Auto-Pilot] Found Excel sheet for sub:', subName);

      const wrapper = document.createElement('div');
      wrapper.className = 'rt-injected-container';
      wrapper.innerHTML = `
        <div class="rt-card">
          <div class="rt-card-header">
            <span class="rt-badge-target">🎯 Target: <b>${subName}</b></span>
          </div>
          <button class="rt-btn-autoupload" data-filename="${downloadAttr}" data-href="${href}" data-subname="${subName}">
            🚀 Auto Upload to RouteTrust
          </button>
          <div class="rt-card-status" style="display:none;"></div>
        </div>
      `;

      // Insert right below the download button
      parentContainer.appendChild(wrapper);

      // Handle Click on Auto Upload
      const btn = wrapper.querySelector('.rt-btn-autoupload');
      const statusBox = wrapper.querySelector('.rt-card-status');

      btn.addEventListener('click', async function (e) {
        e.preventDefault();
        e.stopPropagation();

        btn.disabled = true;
        btn.classList.add('rt-loading');
        btn.innerHTML = `⏳ Uploading to RouteTrust in Background...`;
        statusBox.style.display = 'block';
        statusBox.className = 'rt-card-status rt-status-working';
        statusBox.innerHTML = `⏳ Searching <b>${subName}</b> & uploading rate deck...`;

        try {
          // Fetch the Excel blob from the download link
          const res = await fetch(href);
          const blob = await res.blob();
          const base64Data = await blobToBase64(blob);

          chrome.runtime.sendMessage({
            action: 'START_AUTO_UPLOAD',
            subName: subName,
            fileName: downloadAttr,
            fileBase64: base64Data
          }, (response) => {
            console.log('[RouteTrust Auto-Pilot] Started background upload:', response);
          });
        } catch (err) {
          console.error('[RouteTrust Auto-Pilot] Error initiating upload:', err);
          btn.disabled = false;
          btn.innerHTML = `❌ Upload Failed - Try Again`;
          statusBox.className = 'rt-card-status rt-status-error';
          statusBox.innerHTML = `⚠️ Error reading Excel blob: ${err.message}`;
        }
      });
    }
  }

  // Listen for Stage Updates and Success from Background
  chrome.runtime.onMessage.addListener((msg) => {
    console.log('[RouteTrust Auto-Pilot] Portal message received:', msg);

    const statusBoxes = document.querySelectorAll('.rt-card-status');
    const autoBtns = document.querySelectorAll('.rt-btn-autoupload');

    if (msg.action === 'PORTAL_UPDATE_STAGE') {
      statusBoxes.forEach(sb => {
        sb.style.display = 'block';
        sb.className = 'rt-card-status rt-status-working';
        sb.innerHTML = `⏳ ${msg.message}`;
      });
      autoBtns.forEach(btn => {
        btn.innerHTML = `⏳ ${msg.message}`;
      });
    }

    if (msg.action === 'PORTAL_SUCCESS') {
      statusBoxes.forEach(sb => {
        sb.style.display = 'block';
        sb.className = 'rt-card-status rt-status-success';
        sb.innerHTML = `🎉 <b>SUCCESS!</b> Rates Approved & Loaded into RouteTrust (<b>${msg.subName}</b>)!`;
      });
      autoBtns.forEach(btn => {
        btn.disabled = false;
        btn.classList.remove('rt-loading');
        btn.classList.add('rt-success');
        btn.innerHTML = `✅ Rates Approved & Loaded!`;
      });
    }
  });

  // Watch for DOM changes in Streamlit
  setInterval(injectAutoUploadButtons, 1200);
})();
