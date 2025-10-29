// Background service worker for Fabric Capacity Extension
// Currently minimal: reserved for future proactive refresh alarms.
// Provides a message endpoint to force refresh if needed.

const REFRESH_CHECK_ALARM = 'fabric_refresh_check';
const REFRESH_INTERVAL_MIN = 55; // ~55 minutes to stay ahead of 60m access token expiry

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create(REFRESH_CHECK_ALARM, { periodInMinutes: REFRESH_INTERVAL_MIN });
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== REFRESH_CHECK_ALARM) return;
  // Ask popup (if open) to ensure token freshness; if not open nothing happens.
  chrome.runtime.sendMessage({ type: 'BACKGROUND_REFRESH_PING' });
});

// Placeholder to allow future token refresh logic decoupled from popup.
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type === 'PING') {
    sendResponse({ ok: true });
  }
});
