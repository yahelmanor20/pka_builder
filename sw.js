/* Service worker מינימלי — נדרש כדי לאפשר התקנה (Add to Home Screen).
   ללא caching בכוונה: האפליקציה מתעדכנת לעיתים קרובות, וקאש היה מגיש גרסה ישנה.
   ה-fetch handler קיים (pass-through) כי קיומו הוא תנאי להופעת beforeinstallprompt. */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => { /* pass-through לרשת */ });
