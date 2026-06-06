// AFS Portal — Service Worker v1.0
// Polls SharePoint via PA every 2 minutes for new bookings

const CACHE_NAME = 'afs-portal-v1';
const POLL_INTERVAL = 2 * 60 * 1000; // 2 minutes

// ── Install & activate ──────────────────────────
self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
  // Start polling loop
  startPolling();
});

// ── Message from page (sends config on login) ───
self.addEventListener('message', e => {
  if (e.data?.type === 'INIT') {
    // Store portal type, user, and PA URL in SW scope
    self.portalType = e.data.portalType;  // 'logistics' or 'workshop'
    self.currentUser = e.data.currentUser;
    self.paGetUrl = e.data.paGetUrl;
    self.lastSeenRefs = e.data.lastSeenRefs || [];
    startPolling();
  }
  if (e.data?.type === 'LOGOUT') {
    self.portalType = null;
    self.currentUser = null;
    if (self.pollTimer) clearInterval(self.pollTimer);
  }
  if (e.data?.type === 'UPDATE_SEEN') {
    self.lastSeenRefs = e.data.refs || [];
  }
});

// ── Polling ─────────────────────────────────────
function startPolling() {
  if (self.pollTimer) clearInterval(self.pollTimer);
  self.pollTimer = setInterval(poll, POLL_INTERVAL);
}

async function poll() {
  if (!self.paGetUrl || !self.portalType) return;
  try {
    const res = await fetch(self.paGetUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'get' })
    });
    if (!res.ok) return;
    const text = await res.text();
    let items;
    try { items = JSON.parse(text); } catch(e) { items = JSON.parse(JSON.parse(text)); }
    if (!Array.isArray(items)) return;

    const activeItems = items.filter(i => i.Status !== 'Deleted');
    const seen = self.lastSeenRefs || [];

    if (self.portalType === 'workshop') {
      // Workshop: notify on any new booking submission
      const newBookings = activeItems.filter(i =>
        !seen.includes(i.BookingRef || i.Title) &&
        i.Status === 'Pending'
      );
      for (const b of newBookings) {
        const isUrgent = b.Priority === 'ASAP' || b.Priority === 'Urgent';
        await showNotification({
          title: isUrgent ? `🚨 ${b.Priority} — New Booking` : '📋 New Booking Request',
          body: `${b.JobType} — ${b.VehicleReg || 'No reg'} · ${b.DateRequired || ''} · Raised by ${b.RaisedBy}`,
          ref: b.BookingRef || b.Title,
          urgent: isUrgent
        });
      }
    } else if (self.portalType === 'logistics') {
      // Logistics: notify when status changes on bookings raised by this user
      const myBookings = activeItems.filter(i => i.RaisedBy === self.currentUser);
      for (const b of myBookings) {
        const key = `${b.BookingRef || b.Title}-${b.Status}`;
        if (!seen.includes(key) && (b.Status === 'Accepted' || b.Status === 'Completed' || b.Status === 'Declined')) {
          const emoji = b.Status === 'Accepted' ? '✅' : b.Status === 'Completed' ? '🏁' : '❌';
          await showNotification({
            title: `${emoji} Booking ${b.Status}`,
            body: `${b.BookingRef || b.Title} — ${b.JobType}, ${b.VehicleReg || 'No reg'}`,
            ref: b.BookingRef || b.Title,
            urgent: false
          });
        }
      }
    }

    // Update seen refs — include all refs and status keys
    const newSeen = activeItems.map(i => i.BookingRef || i.Title);
    const statusKeys = activeItems.map(i => `${i.BookingRef || i.Title}-${i.Status}`);
    self.lastSeenRefs = [...new Set([...newSeen, ...statusKeys])];

    // Notify the page to update its cache
    const allClients = await clients.matchAll({ type: 'window' });
    for (const client of allClients) {
      client.postMessage({ type: 'UPDATE_SEEN', refs: self.lastSeenRefs });
    }

  } catch(e) {
    console.warn('[SW] Poll failed:', e);
  }
}

async function showNotification({ title, body, ref, urgent }) {
  const options = {
    body,
    icon: '/afs-portal/icon-192.png',
    badge: '/afs-portal/icon-192.png',
    tag: ref,
    renotify: true,
    vibrate: urgent ? [200, 100, 200, 100, 200] : [200],
    data: { ref }
  };
  await self.registration.showNotification(title, options);
}

// ── Notification click → open portal ────────────
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(windowClients => {
      for (const client of windowClients) {
        if (client.url.includes('afs-portal') && 'focus' in client) {
          return client.focus();
        }
      }
      return clients.openWindow('/afs-portal/');
    })
  );
});
