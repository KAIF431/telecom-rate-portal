// content_portal.js - Injected into Telecom Rate Portal
// Relays START_AUTO_UPLOAD from Streamlit to background.js

(function () {
  'use strict';

  console.log('[RouteTrust Auto-Pilot] Portal content script active.');

  // Listen to postMessage from Streamlit
  window.addEventListener('message', function (event) {
    if (event.data && event.data.action === 'START_AUTO_UPLOAD') {
      const { subName, fileName, fileBase64 } = event.data;
      console.log('[RouteTrust Auto-Pilot] Received upload request for sub:', subName);

      try {
        chrome.runtime.sendMessage({
          action: 'START_AUTO_UPLOAD',
          subName: subName,
          fileName: fileName,
          fileBase64: fileBase64
        }, (response) => {
          console.log('[RouteTrust Auto-Pilot] Background upload acknowledged:', response);
        });
      } catch (err) {
        console.error('[RouteTrust Auto-Pilot] Error communicating with extension:', err);
      }
    }
  });
})();
