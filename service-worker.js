const CACHE_NAME = "buku-warung-shell-v2";
const APP_SHELL = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon.svg"];

self.addEventListener("install", event => {
	event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
	self.skipWaiting();
});

self.addEventListener("activate", event => {
	event.waitUntil(
		caches.keys().then(keys => Promise.all(
			keys.filter(key => key.startsWith("buku-warung-shell-") && key !== CACHE_NAME)
				.map(key => caches.delete(key))
		))
	);
	self.clients.claim();
});

self.addEventListener("fetch", event => {
	if (event.request.method !== "GET" || new URL(event.request.url).origin !== self.location.origin) return;
	event.respondWith((async () => {
		const cached = await caches.match(event.request);
		try {
			const response = await fetch(event.request);
			if (response.ok) {
				const cache = await caches.open(CACHE_NAME);
				cache.put(event.request, response.clone());
			}
			return response;
		} catch {
			return cached || (event.request.mode === "navigate" ? caches.match("./index.html") : Response.error());
		}
	})());
});