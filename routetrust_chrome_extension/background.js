// background.js - Full Background Pipeline for RouteTrust Auto-Pilot

// Pre-seeded known subscription IDs mapping
let subscriptionMap = {
  "KAMAHIRAF (VENDOR CLI)": "29168",
  "KAMAHIRAF": "29168",
  "ROZPER (VENDOR CC)": "16415",
  "QUICKCOM TELECOM - ROZPER (VENDOR CC)": "16415"
};

// Load saved mapping from storage
chrome.storage.local.get(['sub_mapping'], (res) => {
  if (res.sub_mapping) {
    subscriptionMap = { ...subscriptionMap, ...res.sub_mapping };
  }
});

let currentTask = null;

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('[Background] Message received:', request.action, request);

  // 1. Triggered from Portal (Auto Upload clicked)
  if (request.action === 'START_AUTO_UPLOAD') {
    const rawSubName = (request.subName || '').trim();
    const cleanSubKey = rawSubName.toUpperCase();

    // Check if we already have the ID in our map
    let targetSubId = null;
    for (const [key, id] of Object.entries(subscriptionMap)) {
      if (cleanSubKey.includes(key) || key.includes(cleanSubKey)) {
        targetSubId = id;
        break;
      }
    }

    // Default fallback to 29168 if Kamahiraf or testing
    if (!targetSubId) {
      if (cleanSubKey.includes('KAMAHIRAF')) targetSubId = '29168';
      else if (cleanSubKey.includes('ROZPER') && cleanSubKey.includes('CC')) targetSubId = '16415';
      else targetSubId = '29168'; // default test ID
    }

    currentTask = {
      subName: rawSubName,
      subId: targetSubId,
      fileName: request.fileName,
      fileBase64: request.fileBase64,
      portalTabId: sender.tab ? sender.tab.id : null,
      stage: 'INIT',
      rtTabId: null
    };

    console.log('[Background] Starting task for:', rawSubName, '-> Sub ID:', targetSubId);

    // Open RouteTrust upload page in a background tab (active: false so it doesn't interrupt Kaif)
    const uploadUrl = `https://mcm.routetrust.com/netpro/sub/rateimport/upload/id/${targetSubId}`;
    chrome.tabs.create({ url: uploadUrl, active: false }, (tab) => {
      currentTask.rtTabId = tab.id;
      sendResponse({ status: 'STARTED', subId: targetSubId, subName: rawSubName });
    });

    return true; // async
  }

  // 2. RouteTrust upload page requesting file
  if (request.action === 'GET_RT_FILE') {
    if (currentTask && currentTask.fileBase64) {
      sendResponse({
        hasData: true,
        fileName: currentTask.fileName,
        fileBase64: currentTask.fileBase64,
        subId: currentTask.subId,
        subName: currentTask.subName
      });
    } else {
      sendResponse({ hasData: false });
    }
    return true;
  }

  // 3. Stage updates from RouteTrust
  if (request.action === 'UPDATE_STAGE') {
    console.log('[Background] Stage update:', request.stage, request.message);
    if (currentTask && currentTask.portalTabId) {
      chrome.tabs.sendMessage(currentTask.portalTabId, {
        action: 'PORTAL_UPDATE_STAGE',
        stage: request.stage,
        message: request.message
      }).catch(() => {});
    }
  }

  // 4. Completed! Rates loaded
  if (request.action === 'RT_COMPLETED') {
    console.log('[Background] RouteTrust pipeline finished successfully for ID:', request.subId);
    if (currentTask && currentTask.portalTabId) {
      chrome.tabs.sendMessage(currentTask.portalTabId, {
        action: 'PORTAL_SUCCESS',
        subName: currentTask.subName,
        subId: request.subId,
        message: `Rates Approved & Loaded Successfully into RouteTrust (${currentTask.subName})!`
      }).catch(() => {});
    }

    // Close the background RouteTrust tab after 2 seconds
    if (currentTask && currentTask.rtTabId) {
      setTimeout(() => {
        chrome.tabs.remove(currentTask.rtTabId).catch(() => {});
        currentTask = null;
      }, 2500);
    }

    sendResponse({ status: 'DONE' });
    return true;
  }

  // 5. Save a new mapping (e.g. from popup or search)
  if (request.action === 'SAVE_SUB_MAPPING') {
    subscriptionMap[request.subName.toUpperCase()] = request.subId;
    chrome.storage.local.set({ sub_mapping: subscriptionMap });
    sendResponse({ status: 'SAVED' });
  }
});
