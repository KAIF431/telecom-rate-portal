// popup.js

document.addEventListener('DOMContentLoaded', function () {
  const subIdInput = document.getElementById('sub_id');
  const openBtn = document.getElementById('btn_open_upload');

  // Load saved subscription ID
  chrome.storage.local.get(['last_sub_id'], function (result) {
    if (result.last_sub_id) {
      subIdInput.value = result.last_sub_id;
    }
  });

  // Save when changed
  subIdInput.addEventListener('input', function () {
    const val = subIdInput.value.trim();
    if (val) {
      chrome.storage.local.set({ last_sub_id: val });
    }
  });

  // Open upload page
  openBtn.addEventListener('click', function () {
    const id = subIdInput.value.trim() || '29168';
    chrome.storage.local.set({ last_sub_id: id });
    const targetUrl = `https://mcm.routetrust.com/netpro/sub/rateimport/upload/id/${id}`;
    chrome.tabs.create({ url: targetUrl });
  });
});
