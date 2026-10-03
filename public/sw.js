// ============================================================================
// Service Worker para Sistema de Control de Cobranza CFDI (PWA)
// Versión: 2.0.0 - Notificaciones Personalizadas y Posposición de Recordatorios
// ============================================================================

const CACHE_NAME = 'cobranza-cfdi-v2.0';
const ASSETS_TO_CACHE = [
  './',
  './standalone.html',
  './icon.svg',
  './manifest.json'
];

// 1. Instalación y Precaching
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

// 2. Activación y Reclamo de Clientes
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

// 3. Estrategia Network-First con fallback a Cache
self.addEventListener('fetch', (event) => {
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
// GESTIÓN DE NOTIFICACIONES PERSONALIZADAS Y POSPOSICIÓN (SNOOZE)
// ============================================================================

// 4. Recepción de mensajes desde la app mediante postMessage
self.addEventListener('message', (event) => {
  if (!event.data) return;

  // A. Mostrar notificación inmediata o recordatorio de pago
  if (event.data.type === 'SHOW_NOTIFICATION') {
    const { title, options } = event.data;
    
    // Si la notificación es un recordatorio de pago con opción a posponer
    const defaultActions = options?.isPaymentReminder ? [
      {
        action: 'open_app',
        title: '📱 Ver Factura'
      },
      {
        action: 'snooze_24h',
        title: '⏳ Posponer 24h'
      },
      {
        action: 'snooze_1h',
        title: '⏱️ Posponer 1h'
      }
    ] : [
      {
        action: 'open_app',
        title: '📱 Abrir Sistema'
      },
      {
        action: 'dismiss',
        title: 'Cerrar'
      }
    ];

    const notificationOptions = {
      body: options?.body || 'Tienes avisos pendientes en tu sistema de cobranza.',
      icon: options?.icon || './icon.svg',
      badge: options?.badge || './icon.svg',
      tag: options?.tag || 'cobranza-cfdi-alert',
      renotify: true,
      requireInteraction: true, // Notificación persistente
      vibrate: [250, 100, 250, 100, 250],
      data: {
        url: options?.url || './',
        reminderId: options?.reminderId || null,
        invoiceId: options?.invoiceId || null,
        folio: options?.folio || '',
        monto: options?.monto || '',
        empresa: options?.empresa || '',
        dateOfArrival: Date.now(),
        ...(options?.data || {})
      },
      actions: options?.actions || defaultActions,
      ...options
    };

    self.registration.showNotification(title || 'Control de Cobranza CFDI', notificationOptions);
  }

  // B. Programación diferida a nivel Service Worker (para cuando la pestaña está en segundo plano)
  if (event.data.type === 'SCHEDULE_SNOOZE_TIMER') {
    const { delayMs, title, options } = event.data;
    if (delayMs > 0 && delayMs < 2147483647) {
      setTimeout(() => {
        self.registration.showNotification(title, {
          ...options,
          icon: './icon.svg',
          badge: './icon.svg',
          requireInteraction: true
        });
      }, delayMs);
    }
  }
});

// 5. Soporte para eventos Push
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
    requireInteraction: true,
    vibrate: [200, 100, 200, 100, 200],
    data: {
      url: payload.url || './',
      dateOfArrival: Date.now(),
      ...payload.data
    },
    actions: payload.actions || [
      {
        action: 'open_app',
        title: '📱 Abrir Sistema'
      },
      {
        action: 'snooze_24h',
        title: '⏳ Posponer 24h'
      }
    ]
  };

  event.waitUntil(
    self.registration.showNotification(payload.title || 'Control de Cobranza CFDI', pushOptions)
  );
});

// 6. Interacción del usuario: Clic en Notificación o Botón de Acción ('Posponer' / 'Abrir')
self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  const action = event.action;
  const notifData = event.notification.data || {};
  const targetUrl = notifData.url || './';
  const folio = notifData.folio || 'seleccionada';

  // Caso 1: El usuario seleccionó POSPONER 24 Horas
  if (action === 'snooze_24h' || action === 'snooze') {
    event.waitUntil(
      handleSnoozeAction(notifData, 24, folio)
    );
    return;
  }

  // Caso 2: El usuario seleccionó POSPONER 1 Hora
  if (action === 'snooze_1h') {
    event.waitUntil(
      handleSnoozeAction(notifData, 1, folio)
    );
    return;
  }

  // Caso 3: Descartar / Cerrar
  if (action === 'dismiss') {
    return;
  }

  // Caso 4: 'open_app' o clic en el cuerpo de la notificación
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      // 1. Enfocar ventana existente si está disponible
      for (const client of clientList) {
        if ('focus' in client) {
          // Notificar a la app abierta qué factura se pulsó
          client.postMessage({
            type: 'NOTIFICATION_FOCUSED_INVOICE',
            invoiceId: notifData.invoiceId,
            folio: notifData.folio
          });
          return client.focus();
        }
      }
      // 2. Abrir nueva ventana si no hay ninguna activa
      if (self.clients.openWindow) {
        return self.clients.openWindow(targetUrl);
      }
    })
  );
});

// Función auxiliar para gestionar la posposición y notificar a los clientes
async function handleSnoozeAction(notifData, hours, folio) {
  const snoozeLabel = hours === 1 ? '1 hora' : `${hours} horas`;

  // 1. Mostrar confirmación inmediata al usuario de que el recordatorio fue pospuesto
  await self.registration.showNotification('⏳ Recordatorio Pospuesto', {
    body: `El aviso para la factura ${folio} se ha pospuesto por ${snoozeLabel}.`,
    icon: './icon.svg',
    badge: './icon.svg',
    tag: `snooze-ack-${notifData.invoiceId || Date.now()}`,
    renotify: true,
    requireInteraction: false
  });

  // 2. Notificar a todas las ventanas abiertas de la app para que actualicen el almacenamiento local
  const clientList = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  for (const client of clientList) {
    client.postMessage({
      type: 'REMINDER_SNOOZED',
      reminderId: notifData.reminderId,
      invoiceId: notifData.invoiceId,
      folio: notifData.folio,
      hours: hours
    });
  }

  // 3. Programar temporizador de respaldo en el Service Worker para volver a alertar
  const delayMs = hours * 60 * 60 * 1000;
  if (delayMs > 0 && delayMs < 2147483647) {
    setTimeout(() => {
      self.registration.showNotification(`🔔 Recordatorio Reanudado: Factura ${folio}`, {
        body: `Factura ${folio} pospuesta previamente. Fecha de pago pendiente.`,
        icon: './icon.svg',
        badge: './icon.svg',
        tag: `cfdi-reminder-${notifData.invoiceId || Date.now()}`,
        requireInteraction: true,
        vibrate: [250, 100, 250],
        data: notifData,
        actions: [
          { action: 'open_app', title: '📱 Ver Factura' },
          { action: 'snooze_24h', title: '⏳ Posponer 24h' }
        ]
      });
    }, delayMs);
  }
}
