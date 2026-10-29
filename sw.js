// DABS Night service worker: keeps the app opening without signal.
const CACHE='dabs-night-2.0.10292107-test';
const SHELL=['./','./index.html','./app.js?v=2.0.10292107-test','./firebase-config.js','./manifest.webmanifest','./icons/icon-192.png','./icons/icon-512.png','./icons/apple-touch-icon.png'];
const CDN=['www.gstatic.com','cdnjs.cloudflare.com'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const req=e.request;if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin===location.origin){
    // always ask GitHub for the newest file (no stale browser copy), fall back to the saved copy offline
    e.respondWith(fetch(new Request(req.url,{cache:'no-cache',credentials:'same-origin'})).then(r=>{if(r.ok){const c=r.clone();caches.open(CACHE).then(ca=>ca.put(req,c))}return r})
      .catch(()=>caches.match(req,{ignoreSearch:req.mode==='navigate'}).then(r=>r||caches.match('./index.html'))));
    return;
  }
  if(CDN.includes(url.hostname)){
    e.respondWith(caches.match(req).then(r=>r||fetch(req).then(n=>{if(n.ok){const c=n.clone();caches.open(CACHE).then(ca=>ca.put(req,c))}return n})));
  }
});
