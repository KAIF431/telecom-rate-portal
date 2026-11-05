// background.js - Service Worker for RouteTrust Auto-Pilot

let subscriptionMap = {
  "KAMAHIRAF (VENDOR CLI)": "29168",
  "KAMAHIRAF": "29168",
  "ROZPER (VENDOR CC)": "16415",
  "QUICKCOM TELECOM - ROZPER (VENDOR CC)": "16415"
};

chrome.storage.local.get(['sub_mapping'], (res) => {
  if (res.sub_mapping) {
    subscriptionMap = { ...subscriptionMap, ...res.sub_mapping };
  }
});

let currentTask = null;

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('[Background] Message received:', request.action, request);

  if (request.action === 'START_AUTO_UPLOAD') {
    const rawSubName = (request.subName || '').trim();
    const cleanSubKey = rawSubName.toUpperCase();

    let targetSubId = null;
    for (const [key, id] of Object.entries(subscriptionMap)) {
      if (cleanSubKey.includes(key) || key.includes(cleanSubKey)) {
        targetSubId = id;
        break;
      }
    }

    if (!targetSubId) {
      if (cleanSubKey.includes('KAMAHIRAF')) targetSubId = '29168';
      else if (cleanSubKey.includes('ROZPER') && cleanSubKey.includes('CC')) targetSubId = '16415';
      else targetSubId = '29168';
    }

    currentTask = {
      subName: rawSubName,
      subId: targetSubId,
      fileName: request.fileName,
      fileBase64: request.fileBase64,
      portalTabId: sender.tab ? sender.tab.id : null,
      rtTabId: null
    };

    console.log('[Background] Opening RouteTrust for:', rawSubName, '-> Sub ID:', targetSubId);

    // Open RouteTrust tab actively so Kaif can watch the magic live!
    const uploadUrl = `https://mcm.routetrust.com/netpro/sub/rateimport/upload/id/${targetSubId}`;
    chrome.tabs.create({ url: uploadUrl, active: true }, (tab) => {
      currentTask.rtTabId = tab.id;
      sendResponse({ status: 'STARTED', subId: targetSubId, subName: rawSubName });
    });

    return true;
  }

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

  if (request.action === 'RT_COMPLETED') {
    console.log('[Background] Pipeline finished successfully for ID:', request.subId);
    if (currentTask && currentTask.portalTabId) {
      chrome.tabs.sendMessage(currentTask.portalTabId, {
        action: 'PORTAL_SUCCESS',
        subName: currentTask.subName,
        subId: request.subId,
        message: `Rates Approved & Loaded Successfully into RouteTrust (${currentTask.subName})!`
      }).catch(() => {});
    }

    sendResponse({ status: 'DONE' });
    return true;
  }
});
