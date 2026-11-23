/* Vision portal service worker — makes the portal installable and quick on weak signal.
   Pages: network first (always the latest when there's signal), saved copy when offline.
   Code libraries: saved copy first. Microsoft sign-in and SharePoint data are NEVER cached. */
const CACHE = "vision-v20";
const PRECACHE = ["/", "/feedback.js", "/manifest.webmanifest", "/icon-192.png", "/icon-512.png", "/offline.html", "/absence.html","/adminreview.html","/assigntask.html","/certifications.html","/closecall.html","/closeout.html","/compliance.html","/dashboard.html","/display.html","/employees.html","/equipment.html","/expense.html","/fieldfiles.html","/fitting.html","/fleet.html","/incident.html","/index.html","/jobsetup.html","/library.html","/lookahead.html","/lookaheadreview.html","/manpower.html","/mysubmissions.html","/mytasks.html","/phasecodes.html","/pmvisit.html","/portal.html","/pressuretest.html","/pretask.html","/rentals.html","/servicereport.html","/shopbillout.html","/shopdash.html","/shoplog.html","/shopproduction.html","/shopqueue.html","/startupreport.html","/taskschedule.html","/tcadmin.html","/timecard.html","/tmticket.html","/toolbox.html","/training.html","/transmittal.html","/tv.html","/bids.html","/orders.html","/punch.html","/weeklyreport.html"];
self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(PRECACHE.map(u => c.add(new Request(u, {cache: "reload"})).catch(() => null)))).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
const LIB_HOSTS = ["cdn.jsdelivr.net", "cdnjs.cloudflare.com"];
self.addEventListener("fetch", e => {
  const req = e.request; if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (LIB_HOSTS.includes(url.hostname)) {                       // libraries: cache first
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => { if (res.ok || res.type === "opaque") { const c = res.clone(); caches.open(CACHE).then(x => x.put(req, c)); } return res; })));
    return;
  }
  if (url.origin !== self.location.origin) return;               // Microsoft login, Graph, SharePoint: straight to the network
  const isPage = req.mode === "navigate" || /\.html?$/.test(url.pathname) || url.pathname === "/";
  if (isPage) {                                                 // pages: network first, fall back to saved copy, then offline page
    e.respondWith(fetch(req).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(url.pathname, c)); } return res; })
      .catch(() => caches.match(url.pathname).then(hit => hit || caches.match("/offline.html"))));
    return;
  }
  e.respondWith(fetch(req).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(req, c)); } return res; }).catch(() => caches.match(req)));
});
