// ============================================================================
// Service Worker para Sistema de Control de Cobranza CFDI (PWA)
// Versión: 1.2.0
// ============================================================================

const CACHE_NAME = 'cobranza-cfdi-v2';
const ASSETS_TO_CACHE = [
  './',
  './standalone.html',
  './icon.svg',
  './manifest.json'
];

// 1. Instalación del Service Worker y precaching de recursos esenciales
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE).catch((err) => {
        console.warn('Advertencia en precache de recursos:', err);
      });
    })
  );
});

// 2. Activación y limpieza de caches antiguos
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

// 3. Estrategia de Red Primero (Network-First) con respaldo en Cache
self.addEventListener('fetch', (event) => {
  // Ignorar peticiones que no sean GET o que sean hacia servicios en tiempo real (Firestore / Google APIs)
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (
    url.origin.includes('firestore') || 
    url.origin.includes('googleapis') || 
    url.origin.includes('identitytoolkit')
  ) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      })
      .catch(() => caches.match(event.request))
  );
});

// ============================================================================
// GESTIÓN DE NOTIFICACIONES PERSISTENTES Y POSTMESSAGE
// ============================================================================

// 4. Recepción de mensajes desde la app mediante postMessage ('SHOW_NOTIFICATION')
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SHOW_NOTIFICATION') {
    const { title, options } = event.data;
    
    // Configuración para notificación persistente con acción de apertura
    const notificationOptions = {
      body: options?.body || 'Tienes avisos pendientes en tu sistema de cobranza.',
      icon: options?.icon || './icon.svg',
      badge: options?.badge || './icon.svg',
      tag: options?.tag || 'cobranza-cfdi-alert',
      renotify: true,
      requireInteraction: true, // Notificación persistente: permanece visible hasta que el usuario interactúa
      vibrate: [200, 100, 200, 100, 200],
      data: {
        url: options?.url || './',
        dateOfArrival: Date.now()
      },
      actions: [
        {
          action: 'open_app',
          title: '📱 Abrir Sistema'
        },
        {
          action: 'dismiss',
          title: 'Cerrar'
        }
      ],
      ...options
    };

    self.registration.showNotification(title || 'Control de Cobranza CFDI', notificationOptions);
  }
});

// 5. Soporte para eventos Push del servidor
self.addEventListener('push', (event) => {
  let payload = {
    title: 'Control de Cobranza CFDI',
    body: 'Tienes facturas próximas a vencer o comprobantes nuevos por revisar.'
  };

  if (event.data) {
    try {
      payload = event.data.json();
    } catch (e) {
      payload.body = event.data.text();
    }
  }

  const pushOptions = {
    body: payload.body,
    icon: './icon.svg',
    badge: './icon.svg',
    tag: payload.tag || 'cfdi-push-alert',
    renotify: true,
    requireInteraction: true, // Notificación persistente
    vibrate: [200, 100, 200, 100, 200],
    data: {
      url: payload.url || './',
      dateOfArrival: Date.now()
    },
    actions: [
      {
        action: 'open_app',
        title: '📱 Abrir Sistema'
      }
    ]
  };

  event.waitUntil(
    self.registration.showNotification(payload.title || 'Control de Cobranza CFDI', pushOptions)
  );
});

// 6. Interacción del usuario con la notificación (Abrir o enfocar la aplicación)
self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  // Si el usuario presionó 'Cerrar', simplemente se descarta
  if (event.action === 'dismiss') {
    return;
  }

  const targetUrl = event.notification.data?.url || './';

  // Buscar ventanas existentes para enfocar, o abrir una nueva
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      // 1. Si ya hay una ventana activa de la aplicación, darle foco
      for (const client of clientList) {
        if ('focus' in client) {
          return client.focus();
        }
      }
      // 2. Si no hay ventanas activas, abrir una nueva con la URL de la app
      if (self.clients.openWindow) {
        return self.clients.openWindow(targetUrl);
      }
    })
  );
});
