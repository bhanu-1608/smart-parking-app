/**
 * DeccanPulse API & Live WebSocket Client.
 */

async function request(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.detail || `Request failed (${res.status})`);
  }
  return data;
}

export const api = {
  getDestinations(search = "", userX = null, userY = null) {
    const params = new URLSearchParams({ search });
    if (userX !== null && userY !== null) {
      params.set("user_x", userX);
      params.set("user_y", userY);
    }
    return request(`/api/destinations?${params.toString()}`);
  },
  getFacilities(params) {
    const clean = Object.fromEntries(Object.entries(params).filter(([, v]) => v !== null && v !== undefined && v !== ""));
    const qs = new URLSearchParams(clean).toString();
    return request(`/api/facilities?${qs}`);
  },
  getFacility(id, userX, userY) {
    return request(`/api/facilities/${encodeURIComponent(id)}?user_x=${userX}&user_y=${userY}`);
  },
  toggleWatchlist(id) {
    return request(`/api/facilities/${encodeURIComponent(id)}/watchlist`, { method: "POST" });
  },
  getNavigation(id, userX, userY) {
    return request(`/api/navigation/${encodeURIComponent(id)}?user_x=${userX}&user_y=${userY}`);
  },
  getVehicles() {
    return request("/api/vehicles");
  },
  createVehicle(payload) {
    return request("/api/vehicles", { method: "POST", body: JSON.stringify(payload) });
  },
  deleteVehicle(id) {
    return request(`/api/vehicles/${encodeURIComponent(id)}`, { method: "DELETE" });
  },
  getQuote(payload) {
    return request("/api/quote", { method: "POST", body: JSON.stringify(payload) });
  },
  getReservations(status = "all") {
    const qs = status && status !== "all" ? `?status=${encodeURIComponent(status)}` : "";
    return request(`/api/reservations${qs}`);
  },
  createReservation(payload) {
    return request("/api/reservations", { method: "POST", body: JSON.stringify(payload) });
  },
  enterReservation(id) {
    return request(`/api/reservations/${encodeURIComponent(id)}/enter`, { method: "POST" });
  },
  exitReservation(id) {
    return request(`/api/reservations/${encodeURIComponent(id)}/exit`, { method: "POST" });
  },
  extendReservation(id, extraMins, paymentMethod) {
    return request(`/api/reservations/${encodeURIComponent(id)}/extend`, {
      method: "POST",
      body: JSON.stringify({ extra_mins: extraMins, payment_method: paymentMethod }),
    });
  },
  cancelReservation(id) {
    return request(`/api/reservations/${encodeURIComponent(id)}/cancel`, { method: "POST" });
  },
  getNotifications() {
    return request("/api/notifications");
  },
  markNotificationsRead() {
    return request("/api/notifications/read", { method: "POST" });
  },
  pulseSensor() {
    return request("/api/simulate/pulse", { method: "POST" });
  },
};

export function connectLiveSocket(onMessage) {
  const proto = window.location.protocol === "https:" ? "wss:" : "ws:";
  const url = `${proto}//${window.location.host}/ws/live`;
  let ws;
  let timer;

  function open() {
    ws = new WebSocket(url);
    ws.onmessage = (evt) => {
      try {
        onMessage(JSON.parse(evt.data));
      } catch {
        // ignore invalid packet
      }
    };
    ws.onclose = () => {
      timer = setTimeout(open, 3500);
    };
    ws.onerror = () => ws.close();
  }

  open();
  return () => {
    clearTimeout(timer);
    if (ws) ws.close();
  };
}
