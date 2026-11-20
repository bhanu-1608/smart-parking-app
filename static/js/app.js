import { api, connectLiveSocket } from "./api.js";
import { projLatLng, renderHyderabadMap } from "./map.js";
import { generateQrSvg, renderNavFloat } from "./nav.js";

const VEHICLE_CATALOG = {
  car: [
    { make_model: "Maruti Suzuki Swift", car_type: "Hatchback", is_ev: false, charger_pref: "None" },
    { make_model: "Maruti Suzuki Baleno", car_type: "Hatchback", is_ev: false, charger_pref: "None" },
    { make_model: "Hyundai i20", car_type: "Hatchback", is_ev: false, charger_pref: "None" },
    { make_model: "Honda City", car_type: "Sedan", is_ev: false, charger_pref: "None" },
    { make_model: "Hyundai Verna", car_type: "Sedan", is_ev: false, charger_pref: "None" },
    { make_model: "Hyundai Creta", car_type: "SUV", is_ev: false, charger_pref: "None" },
    { make_model: "Kia Seltos", car_type: "SUV", is_ev: false, charger_pref: "None" },
    { make_model: "Mahindra XUV700", car_type: "SUV", is_ev: false, charger_pref: "None" },
    { make_model: "Mahindra Thar", car_type: "SUV", is_ev: false, charger_pref: "None" },
    { make_model: "Toyota Innova Hycross", car_type: "SUV", is_ev: false, charger_pref: "None" },
  ],
  ev_car: [
    { make_model: "Tata Tiago.ev", car_type: "EV Car", is_ev: true, charger_pref: "CCS2 DC" },
    { make_model: "Tata Nexon.ev", car_type: "EV Car", is_ev: true, charger_pref: "CCS2 DC" },
    { make_model: "MG ZS EV", car_type: "EV Car", is_ev: true, charger_pref: "CCS2 DC" },
    { make_model: "Mahindra XUV400 EV", car_type: "EV Car", is_ev: true, charger_pref: "CCS2 DC" },
    { make_model: "BYD Atto 3 EV", car_type: "EV Car", is_ev: true, charger_pref: "CCS2 DC" },
  ],
  two_wheeler: [
    { make_model: "Honda Activa 6G", car_type: "Scooty / Bike", is_ev: false, charger_pref: "None" },
    { make_model: "TVS Jupiter", car_type: "Scooty / Bike", is_ev: false, charger_pref: "None" },
    { make_model: "Suzuki Access 125", car_type: "Scooty / Bike", is_ev: false, charger_pref: "None" },
    { make_model: "Royal Enfield Classic 350", car_type: "Scooty / Bike", is_ev: false, charger_pref: "None" },
    { make_model: "Bajaj Pulsar 150", car_type: "Scooty / Bike", is_ev: false, charger_pref: "None" },
  ],
  ev_two_wheeler: [
    { make_model: "Ather 450X", car_type: "EV Scooty", is_ev: true, charger_pref: "3.3kW AC" },
    { make_model: "Ola S1 Pro", car_type: "EV Scooty", is_ev: true, charger_pref: "3.3kW AC" },
    { make_model: "TVS iQube", car_type: "EV Scooty", is_ev: true, charger_pref: "3.3kW AC" },
    { make_model: "Bajaj Chetak EV", car_type: "EV Scooty", is_ev: true, charger_pref: "3.3kW AC" },
    { make_model: "Revolt RV400 EV", car_type: "EV Scooty", is_ev: true, charger_pref: "3.3kW AC" },
  ],
};

const PAYMENT_OPTIONS = [
  { id: "UPI Instant (GPay / PhonePe)", label: "UPI Instant", sub: "GPay / PhonePe / Paytm" },
  { id: "FASTag Auto-Pay", label: "FASTag Auto-Pay", sub: "Barrier Auto-Debit" },
  { id: "RuPay / HDFC Card", label: "RuPay / Card", sub: "Debit or Credit Card" },
  { id: "NetBanking", label: "NetBanking", sub: "SBI / HDFC / ICICI" },
];

const state = {
  page: "map", // "map" | "hubs" | "book" | "ticket" | "vehicles" | "history" | "alerts"
  facilities: [],
  activeDestination: null,
  selectedFacilityId: null,
  selectedFacilityDetail: null,
  selectedSlotId: null,
  slotZoneTab: "regular_deck", // "regular_deck" | "two_wheeler_floor" | "ev_charging_zone"
  evVehicleType: "car", // "car" | "two_wheeler" inside EV Wing
  activeRegularLevel: null,
  bookingDuration: 120,
  bookingPayMethod: PAYMENT_OPTIONS[0].id,
  bookingPromo: "",
  bookingIncludeParking: true,
  bookingChargeEv: false,
  vehicles: [],
  activeVehicleId: null,
  reservations: [],
  notifications: [],
  locationEnabled: false,
  userPos: { x: 585, y: 217 },
  maxDistanceKm: "all",
  floorFilter: "all", // "all" | "car" | "2w" | "ev"
  onlyCovered: false,
  hubSearch: "",
  hubSort: "distance",
  mapView: { zoom: 1.0, panX: 0, panY: 0 },
  activeRoute: null,
  navStepIndex: 0,
  navSimPos: null,
  navTimer: null,
  carFormTab: "car", // "car" | "ev_car" | "two_wheeler"
  carFormCustom: "",
  carFormSelectedIdx: 0,
  carFormPlate: "TS 09 EV 4021",
};

function esc(v) {
  return String(v ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function inr(n) {
  return `₹${Math.round(Number(n || 0)).toLocaleString("en-IN")}`;
}

function toast(msg) {
  const area = document.getElementById("toast-area");
  const el = document.createElement("div");
  el.className = "toast-pill";
  el.textContent = msg;
  area.appendChild(el);
  setTimeout(() => el.remove(), 3400);
}

function activeVehicle() {
  return state.vehicles.find((v) => v.id === state.activeVehicleId) || state.vehicles[0] || null;
}

function formatTime(iso) {
  return new Date(iso).toLocaleString("en-IN", { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
}

function renderMetroBadges(fac) {
  const lines = fac.metro_lines || [];
  if (!lines.length && !fac.metro_station) return "";
  const badges = lines
    .map((line) => {
      const cls = line.toLowerCase();
      return `<span class="metro-badge"><span class="metro-dot metro-${cls}"></span>${esc(line)} Line</span>`;
    })
    .join("");
  const station = fac.metro_station ? `<span class="sub-txt">${esc(fac.metro_station)}</span>` : "";
  return `<div class="metro-row">${badges}${station}</div>`;
}

function goToPage(pageName) {
  state.page = pageName;
  document.querySelectorAll(".page-view").forEach((s) => s.classList.toggle("active", s.id === `page-${pageName}`));
  document.querySelectorAll(".rail-item, .m-btn").forEach((b) => {
    const highlight = pageName === "book" ? "hubs" : pageName;
    b.classList.toggle("active", b.dataset.page === highlight);
  });
  if (pageName === "alerts") {
    api.markNotificationsRead().then(() => {
      state.notifications.forEach((n) => (n.is_read = true));
      updateAlertBadge();
    });
  }
  renderActivePage();
}

async function refreshData() {
  const params = {
    user_x: state.userPos.x,
    user_y: state.userPos.y,
    only_ev: state.floorFilter === "ev",
    only_two_wheeler: state.floorFilter === "2w",
    only_covered: state.onlyCovered,
  };
  if (state.locationEnabled && state.maxDistanceKm !== "all") {
    params.max_distance_km = state.maxDistanceKm;
  }
  if (state.activeDestination) {
    params.dest_x = state.activeDestination.x;
    params.dest_y = state.activeDestination.y;
  }

  const [facs, resList, notifList, vehs] = await Promise.all([
    api.getFacilities(params),
    api.getReservations("all"),
    api.getNotifications(),
    api.getVehicles(),
  ]);

  state.facilities = facs;
  state.reservations = resList;
  state.notifications = notifList;
  state.vehicles = vehs;
  if (!activeVehicle() && vehs.length > 0) state.activeVehicleId = vehs[0].id;

  if (state.selectedFacilityId) {
    state.selectedFacilityDetail = await api.getFacility(state.selectedFacilityId, state.userPos.x, state.userPos.y);
  }
  updateAlertBadge();
  renderActivePage();
}

function updateAlertBadge() {
  const unread = state.notifications.filter((n) => !n.is_read).length;
  const badge = document.getElementById("rail-alert-badge");
  if (unread > 0) {
    badge.textContent = String(unread);
    badge.classList.remove("hidden");
  } else {
    badge.classList.add("hidden");
  }
}

function toggleUserLocation() {
  if (state.locationEnabled) {
    state.locationEnabled = false;
    state.maxDistanceKm = "all";
    document.getElementById("btn-location-toggle").textContent = "Location: OFF";
    document.getElementById("btn-location-toggle").classList.remove("active");
    document.getElementById("map-radius-select").classList.add("hidden");
    refreshData();
    return;
  }

  const activateAt = (x, y, label) => {
    state.locationEnabled = true;
    state.userPos = { x, y };
    document.getElementById("btn-location-toggle").textContent = "Location: ON";
    document.getElementById("btn-location-toggle").classList.add("active");
    document.getElementById("map-radius-select").classList.remove("hidden");
    toast(label);
    refreshData();
  };

  if (navigator.geolocation) {
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        const lat = Math.max(17.22, Math.min(17.51, pos.coords.latitude));
        const lng = Math.max(78.25, Math.min(78.63, pos.coords.longitude));
        const [px, py] = projLatLng(lat, lng);
        activateAt(px, py, "Location enabled — distances updated");
      },
      () => activateAt(470, 210, "Location enabled (tap anywhere on map to move pin)"),
      { timeout: 2500 }
    );
  } else {
    activateAt(470, 210, "Location enabled (tap anywhere on map to move pin)");
  }
}

function renderMapPage() {
  renderHyderabadMap(document.getElementById("hyderabad-svg"), state, {
    onSelectFacility: async (facId) => {
      state.selectedFacilityId = facId;
      state.selectedFacilityDetail = await api.getFacility(facId, state.userPos.x, state.userPos.y);
      renderMapPage();
    },
  });

  renderMapInspector();
  renderNavFloat(document.getElementById("map-nav-banner"), state.activeRoute, state.navStepIndex, Boolean(state.navTimer), {
    onSimulateDrive: startDriveSimulation,
    onStopNav: stopNavigation,
  });
}

function renderMapInspector() {
  const panel = document.getElementById("map-inspector");
  const d = state.selectedFacilityDetail;

  if (d) {
    panel.innerHTML = `
      <div class="row-between">
        <div class="area-title-row">
          <strong class="area-name">${esc(d.name)}</strong>
          <span class="area-tag">${esc(d.tag)}</span>
        </div>
        <button id="insp-close" class="btn btn-soft btn-sm">✕</button>
      </div>

      <div class="row-between sub-txt" style="align-items:center;">
        <span>${esc(d.address)}</span>
        <strong class="mono-val" style="color:var(--accent-gold);white-space:nowrap;">★ ${d.rating.toFixed(1)}</strong>
      </div>

      ${renderMetroBadges(d)}

      ${state.locationEnabled ? `<div class="sub-txt"><strong>${d.distance_km} km (${d.bearing})</strong> · ${d.drive_mins} min drive</div>` : ""}

      ${
        d.amenities && d.amenities.length
          ? `<div class="amenity-row">${d.amenities.map((a) => `<span class="amenity-chip">${esc(a)}</span>`).join("")}</div>`
          : ""
      }

      <div class="floor-metrics">
        <div class="metric-cell">
          <div class="sub-txt">Car Deck</div>
          <strong class="mono-val">${d.regular_available}/${d.regular_total}</strong>
          <div class="sub-txt mono-val">${inr(d.hourly_rate)}/hr</div>
        </div>
        <div class="metric-cell">
          <div class="sub-txt">Scooty/Bike</div>
          <strong class="mono-val">${d.two_wheeler_available}/${d.two_wheeler_total}</strong>
          <div class="sub-txt mono-val">${inr(d.two_wheeler_rate)}/hr</div>
        </div>
        <div class="metric-cell ev">
          <div class="sub-txt" style="color:var(--accent-teal);">EV Wing</div>
          <strong class="mono-val">C${d.ev_car_available} · 2W${d.ev_2w_available}</strong>
          <div class="sub-txt mono-val">₹${d.ev_kwh_rate}/₹${d.ev_2w_kwh_rate} kWh</div>
        </div>
      </div>

      ${
        d.ev_charger_types && d.ev_charger_types.length
          ? `<div class="sub-txt"><strong>Chargers:</strong> ${esc(d.ev_charger_types.join(" · "))}</div>`
          : ""
      }
      <div class="sub-txt"><strong>Gate:</strong> ${esc(d.entrance_hint)}</div>

      <div style="display:flex;gap:8px;margin-top:auto;">
        <button id="insp-book" class="btn btn-primary" style="flex:1;">Choose Slot</button>
        <button id="insp-dir" class="btn btn-soft">Directions</button>
        <button id="insp-watch" class="btn btn-soft">${d.watchlisted ? "Watching" : "Alert"}</button>
      </div>
    `;

    panel.querySelector("#insp-close").addEventListener("click", () => {
      state.selectedFacilityDetail = null;
      state.selectedFacilityId = null;
      renderMapPage();
    });
    panel.querySelector("#insp-book").addEventListener("click", () => openHubBookingPage(d.id));
    panel.querySelector("#insp-dir").addEventListener("click", () => startNavigation(d.id));
    panel.querySelector("#insp-watch").addEventListener("click", async () => {
      const r = await api.toggleWatchlist(d.id);
      toast(r.watchlisted ? "Spot alert enabled" : "Spot alert removed");
      await refreshData();
    });
    return;
  }

  panel.innerHTML = `
    <div class="row-between" style="align-items:center;">
      <strong class="area-name" style="font-size:16px;">${state.activeDestination ? `Near ${esc(state.activeDestination.name)}` : "Parking Areas"}</strong>
      <span class="sub-txt mono-val">${state.facilities.length} areas</span>
    </div>
    <p class="sub-txt">Select a P marker on the map or pick an area below for Metro, EV chargers &amp; bay details.</p>
    <div style="display:flex;flex-direction:column;gap:8px;">
      ${state.facilities
        .slice(0, 8)
        .map(
          (f) => `
          <button class="card soft" style="text-align:left;padding:11px;gap:6px;" data-pick-area="${f.id}">
            <div class="row-between" style="width:100%;">
              <div class="area-title-row">
                <strong class="area-name" style="font-size:14px;">${esc(f.name)}</strong>
                <span class="area-tag">${esc(f.tag)}</span>
              </div>
              <span class="rate-pill">${inr(f.hourly_rate)}/hr</span>
            </div>
            ${renderMetroBadges(f)}
            <div class="row-between sub-txt mono-val" style="width:100%;">
              <span>Car ${f.regular_available} · 2W ${f.two_wheeler_available} · EV ${f.ev_car_available}+${f.ev_2w_available}</span>
              <span>${state.locationEnabled ? `${f.distance_km} km` : `${f.drive_mins}m`}</span>
            </div>
          </button>`
        )
        .join("")}
    </div>
  `;

  panel.querySelectorAll("[data-pick-area]").forEach((b) =>
    b.addEventListener("click", async () => {
      state.selectedFacilityId = b.dataset.pickArea;
      state.selectedFacilityDetail = await api.getFacility(state.selectedFacilityId, state.userPos.x, state.userPos.y);
      renderMapPage();
    })
  );
}

async function openHubBookingPage(facId) {
  state.selectedFacilityId = facId;
  state.selectedFacilityDetail = await api.getFacility(facId, state.userPos.x, state.userPos.y);
  const veh = activeVehicle();
  if (veh && veh.car_type === "EV Scooty") {
    state.slotZoneTab = "ev_charging_zone";
    state.evVehicleType = "two_wheeler";
    state.bookingIncludeParking = false;
    state.bookingChargeEv = true;
  } else if (veh && veh.is_ev) {
    state.slotZoneTab = "ev_charging_zone";
    state.evVehicleType = "car";
    state.bookingIncludeParking = false;
    state.bookingChargeEv = true;
  } else if (veh && veh.car_type === "Scooty / Bike") {
    state.slotZoneTab = "two_wheeler_floor";
    state.evVehicleType = "two_wheeler";
    state.bookingIncludeParking = true;
    state.bookingChargeEv = false;
  } else {
    state.slotZoneTab = "regular_deck";
    state.evVehicleType = "car";
    state.bookingIncludeParking = true;
    state.bookingChargeEv = false;
  }
  state.bookingPromo = "";
  state.selectedSlotId = null;
  const regLevels = state.selectedFacilityDetail.levels.filter((l) => l !== "EV-Z1" && !l.startsWith("2W-"));
  state.activeRegularLevel = regLevels[0] || "L1";
  goToPage("book");
}

function renderHubsPage() {
  const el = document.getElementById("hubs-page-content");
  let list = [...state.facilities];
  if (state.hubSearch.trim()) {
    const q = state.hubSearch.trim().toLowerCase();
    list = list.filter(
      (f) => f.name.toLowerCase().includes(q) || f.tag.toLowerCase().includes(q) || f.address.toLowerCase().includes(q)
    );
  }
  if (state.hubSort === "price") list.sort((a, b) => a.hourly_rate - b.hourly_rate);
  else if (state.hubSort === "available") list.sort((a, b) => b.available_bays - a.available_bays);

  el.innerHTML = `
    <div class="row-between" style="align-items:center;flex-wrap:wrap;">
      <div class="phase-title-wrap">
        <span class="phase-eyebrow phase-hubs-tag">Directory &amp; Live Grid</span>
        <h2 class="page-heading">Parking Hubs (${list.length})</h2>
        <p class="sub-txt">${state.locationEnabled ? "Distances shown from your enabled location" : "Select an area to inspect Car, Scooty/Bike, and EV floors"}</p>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <button id="hubs-loc-btn" class="ctrl-btn ${state.locationEnabled ? "active" : ""}">
          ${state.locationEnabled ? "Location: ON" : "Enable Location"}
        </button>
        <select id="hubs-sort" class="ctrl-btn">
          <option value="distance" ${state.hubSort === "distance" ? "selected" : ""}>${state.locationEnabled ? "Nearest First" : "Default Order"}</option>
          <option value="price" ${state.hubSort === "price" ? "selected" : ""}>Lowest Rate</option>
          <option value="available" ${state.hubSort === "available" ? "selected" : ""}>Most Open</option>
        </select>
      </div>
    </div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;">
      <input id="hubs-search-inp" class="input-box" style="flex:1;min-width:220px;" placeholder="Filter by area name or tag..." value="${esc(state.hubSearch)}" />
      <select id="hubs-floor-filter" class="ctrl-btn">
        <option value="all" ${state.floorFilter === "all" ? "selected" : ""}>All Floors</option>
        <option value="car" ${state.floorFilter === "car" ? "selected" : ""}>Car Decks</option>
        <option value="2w" ${state.floorFilter === "2w" ? "selected" : ""}>Scooty &amp; Bike Floors</option>
        <option value="ev" ${state.floorFilter === "ev" ? "selected" : ""}>EV Chargers</option>
      </select>
      <button id="hubs-cov-chk" class="ctrl-btn ${state.onlyCovered ? "active" : ""}">Covered</button>
    </div>

    <div class="bento-grid">
      ${list
        .map(
          (f) => `
          <article class="card">
            <div class="row-between">
              <div class="area-title-row">
                <strong class="area-name">${esc(f.name)}</strong>
                <span class="area-tag">${esc(f.tag)}</span>
              </div>
              <span class="rate-pill">${inr(f.hourly_rate)}/hr</span>
            </div>

            <div class="row-between sub-txt" style="align-items:center;">
              <span>${esc(f.address)}</span>
              <strong class="mono-val" style="color:var(--accent-gold);">★ ${f.rating.toFixed(1)}</strong>
            </div>
            ${renderMetroBadges(f)}
            ${state.locationEnabled ? `<div class="sub-txt"><strong class="mono-val">${f.distance_km} km (${f.bearing})</strong> · ${f.drive_mins} min drive</div>` : ""}
            ${
              f.amenities && f.amenities.length
                ? `<div class="amenity-row">${f.amenities.map((a) => `<span class="amenity-chip">${esc(a)}</span>`).join("")}</div>`
                : ""
            }

            <div class="floor-metrics">
              <div class="metric-cell">
                <div class="sub-txt">Cars</div>
                <strong class="mono-val">${f.regular_available}/${f.regular_total}</strong>
                <div class="sub-txt mono-val">${inr(f.hourly_rate)}/hr</div>
              </div>
              <div class="metric-cell">
                <div class="sub-txt">Scooty/Bike</div>
                <strong class="mono-val">${f.two_wheeler_available}/${f.two_wheeler_total}</strong>
                <div class="sub-txt mono-val">${inr(f.two_wheeler_rate)}/hr</div>
              </div>
              <div class="metric-cell ev">
                <div class="sub-txt" style="color:var(--accent-teal);">EV Wing</div>
                <strong class="mono-val">C${f.ev_car_available} · 2W${f.ev_2w_available}</strong>
                <div class="sub-txt mono-val">₹${f.ev_kwh_rate}/₹${f.ev_2w_kwh_rate} kWh</div>
              </div>
            </div>

            <div style="display:flex;gap:8px;margin-top:auto;">
              <button class="btn btn-primary btn-sm" style="flex:1;" data-hub-book="${f.id}">Choose Slot</button>
              <button class="btn btn-soft btn-sm" data-hub-map="${f.id}">On Map</button>
              <button class="btn btn-soft btn-sm" data-hub-watch="${f.id}">${f.watchlisted ? "Watching" : "Notify"}</button>
            </div>
          </article>`
        )
        .join("")}
    </div>
  `;

  el.querySelector("#hubs-loc-btn").addEventListener("click", toggleUserLocation);
  el.querySelector("#hubs-sort").addEventListener("change", (e) => {
    state.hubSort = e.target.value;
    renderHubsPage();
  });
  el.querySelector("#hubs-search-inp").addEventListener("input", (e) => {
    state.hubSearch = e.target.value;
    renderHubsPage();
  });
  el.querySelector("#hubs-floor-filter").addEventListener("change", (e) => {
    state.floorFilter = e.target.value;
    refreshData();
  });
  el.querySelector("#hubs-cov-chk").addEventListener("click", () => {
    state.onlyCovered = !state.onlyCovered;
    refreshData();
  });
  el.querySelectorAll("[data-hub-book]").forEach((b) => b.addEventListener("click", () => openHubBookingPage(b.dataset.hubBook)));
  el.querySelectorAll("[data-hub-map]").forEach((b) =>
    b.addEventListener("click", async () => {
      state.selectedFacilityId = b.dataset.hubMap;
      state.selectedFacilityDetail = await api.getFacility(state.selectedFacilityId, state.userPos.x, state.userPos.y);
      goToPage("map");
    })
  );
  el.querySelectorAll("[data-hub-watch]").forEach((b) =>
    b.addEventListener("click", async () => {
      const r = await api.toggleWatchlist(b.dataset.hubWatch);
      toast(r.watchlisted ? "Spot alert enabled" : "Spot alert removed");
      await refreshData();
    })
  );
}

async function renderBookPage() {
  const el = document.getElementById("book-page-content");
  const d = state.selectedFacilityDetail;
  if (!d) {
    el.innerHTML = `<div class="card"><p class="sub-txt">Select a parking area from Map or Parking Hubs first.</p><button id="go-hubs-btn" class="btn btn-primary btn-sm">Browse Hubs</button></div>`;
    el.querySelector("#go-hubs-btn").addEventListener("click", () => goToPage("hubs"));
    return;
  }

  const isEvTab = state.slotZoneTab === "ev_charging_zone";
  const isTwTab = state.slotZoneTab === "two_wheeler_floor";
  const regLevels = d.levels.filter((l) => l !== "EV-Z1" && !l.startsWith("2W-"));
  if (!state.activeRegularLevel || !regLevels.includes(state.activeRegularLevel)) {
    state.activeRegularLevel = regLevels[0] || "L1";
  }

  const slots = isEvTab
    ? d.bays.filter((b) => b.zone === "ev_charging_zone")
    : isTwTab
      ? d.bays.filter((b) => b.zone === "two_wheeler_floor")
      : d.bays.filter((b) => b.zone === "regular_deck" && b.level === state.activeRegularLevel);

  const chosenSlot = d.bays.find((b) => b.id === state.selectedSlotId);
  const veh = activeVehicle();
  const evCarSlots = d.bays.filter((b) => b.zone === "ev_charging_zone" && b.vehicle_category === "car");
  const ev2wSlots = d.bays.filter((b) => b.zone === "ev_charging_zone" && b.vehicle_category === "two_wheeler");

  let quoteCategory = "car";
  if (isTwTab) {
    quoteCategory = "two_wheeler";
  } else if (isEvTab) {
    quoteCategory = chosenSlot && chosenSlot.zone === "ev_charging_zone" ? chosenSlot.vehicle_category : state.evVehicleType;
  }

  const includeParking = isEvTab ? state.bookingIncludeParking : true;
  const chargeEv = isEvTab ? state.bookingChargeEv : false;

  const quote = await api.getQuote({
    facility_id: d.id,
    duration_mins: state.bookingDuration,
    zone: state.slotZoneTab,
    vehicle_category: quoteCategory,
    include_parking: includeParking,
    charge_ev: chargeEv,
    promo_code: state.bookingPromo,
  });

  const activeHourly = quoteCategory === "two_wheeler" ? d.two_wheeler_rate : d.hourly_rate;
  const activeEvRate = quoteCategory === "two_wheeler" ? d.ev_2w_kwh_rate : d.ev_kwh_rate;

  const renderSlotBtn = (s) => {
    const avail = s.status === "available";
    const sel = s.id === state.selectedSlotId;
    return `
      <button class="slot-btn ${avail ? "available" : ""} ${sel ? "selected" : ""}" data-bk-slot="${s.id}" ${avail ? "" : "disabled"}>
        <div class="row-between">
          <span class="slot-code-txt">${s.level}-${s.code}</span>
          <span style="font-size:12px;font-weight:700;">${avail ? "Open" : s.status === "reserved" ? "Hold" : "Busy"}</span>
        </div>
        <span class="sub-txt">${isEvTab ? esc(s.charger_standard) : `${s.walk_meters}m walk`}</span>
      </button>`;
  };

  el.innerHTML = `
    <div class="row-between" style="align-items:center;flex-wrap:wrap;">
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
        <button id="bk-back" class="btn btn-soft btn-sm">← Hubs</button>
        <div class="phase-title-wrap">
          <span class="phase-eyebrow phase-book-tag">Slot &amp; Charger Allocation</span>
          <div class="area-title-row">
            <h2 class="page-heading">${esc(d.name)}</h2>
            <span class="area-tag">${esc(d.tag)}</span>
          </div>
        </div>
      </div>
      <span class="rate-pill">${includeParking ? `${inr(activeHourly)}/hr Park` : `${inr(activeEvRate)}/kWh Charge`}</span>
    </div>

    <div class="two-col">
      <!-- Left: Car Deck / Scooty & Bike Floor / EV Wing (Split EV Car & EV 2W) -->
      <div class="card ${isEvTab ? "ev-wing" : ""}">
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:6px;">
          <button id="bk-tab-reg" class="btn btn-sm ${!isEvTab && !isTwTab ? "btn-primary" : "btn-soft"}">Cars (${d.regular_available})</button>
          <button id="bk-tab-2w" class="btn btn-sm ${isTwTab ? "btn-primary" : "btn-soft"}">Scooty/Bike (${d.two_wheeler_available})</button>
          <button id="bk-tab-ev" class="btn btn-sm ${isEvTab ? "btn-teal" : "btn-soft"}">EV Wing (${d.ev_available})</button>
        </div>

        ${
          !isEvTab && !isTwTab && regLevels.length > 1
            ? `<div style="display:flex;gap:6px;align-items:center;">
                <span class="sub-txt">Floor:</span>
                ${regLevels.map((lv) => `<button class="ctrl-btn ${lv === state.activeRegularLevel ? "active" : ""}" data-lvl="${lv}">Level ${lv}</button>`).join("")}
              </div>`
            : `<div class="sub-txt">${isEvTab ? `${esc(d.ev_zone_name)} · Cars ₹${d.ev_kwh_rate}/kWh · Scooty/Bike ₹${d.ev_2w_kwh_rate}/kWh` : isTwTab ? `Dedicated Two-Wheeler Floor · ${inr(d.two_wheeler_rate)}/hr` : esc(d.entrance_hint)}</div>`
        }

        ${
          isEvTab
            ? `<div style="display:flex;flex-direction:column;gap:12px;">
                <div class="ev-split-section">
                  <div class="ev-split-head">
                    <span>EV Car Chargers (${d.ev_car_available}/${d.ev_car_total} open)</span>
                    <span class="rate-pill">₹${d.ev_kwh_rate}/kWh · Park ${inr(d.hourly_rate)}/hr</span>
                  </div>
                  <div class="slot-grid">
                    ${evCarSlots.map(renderSlotBtn).join("")}
                  </div>
                </div>

                <div class="ev-split-section">
                  <div class="ev-split-head">
                    <span>EV Scooty &amp; Bike Chargers (${d.ev_2w_available}/${d.ev_2w_total} open)</span>
                    <span class="rate-pill">₹${d.ev_2w_kwh_rate}/kWh · Park ${inr(d.two_wheeler_rate)}/hr</span>
                  </div>
                  <div class="slot-grid">
                    ${ev2wSlots.map(renderSlotBtn).join("")}
                  </div>
                </div>
              </div>`
            : `<div class="slot-grid">
                ${slots.map(renderSlotBtn).join("")}
              </div>`
        }
      </div>

      <!-- Right: Summary & Separate Individual Option Selection -->
      <div class="card">
        <div class="row-between" style="align-items:center;">
          <div>
            <strong>Active Vehicle</strong>
            <div class="sub-txt">${veh ? `${esc(veh.make_model)} <span class="plate-badge">${esc(veh.plate)}</span>` : "Add your vehicle in the Vehicles page"}</div>
          </div>
          ${!veh ? `<button id="bk-go-vehicles" class="btn btn-soft btn-sm">Open Vehicles</button>` : ""}
        </div>

        <div>
          <div class="sub-txt" style="margin-bottom:6px;font-weight:600;">Duration</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;">
            ${[60, 120, 180, 240]
              .map((m) => `<button class="btn btn-sm ${state.bookingDuration === m ? "btn-primary" : "btn-soft"}" data-bk-dur="${m}">${m / 60}h</button>`)
              .join("")}
          </div>
        </div>

        <div style="display:flex;flex-direction:column;gap:8px;">
          <div class="sub-txt" style="font-weight:600;">Select Individual Service(s)</div>
          ${
            isEvTab
              ? `<label class="service-option ${state.bookingIncludeParking ? "active" : ""}">
                  <span style="display:flex;align-items:center;gap:8px;">
                    <input id="bk-park-chk" type="checkbox" ${state.bookingIncludeParking ? "checked" : ""} />
                    <span><strong>Parking Space Fee</strong> <span class="sub-txt">(${quoteCategory === "two_wheeler" ? "2W" : "Car"} · ${inr(activeHourly)}/hr)</span></span>
                  </span>
                  <strong class="mono-val">${state.bookingIncludeParking ? inr(quote.base_fee) : "₹0"}</strong>
                </label>
                <label class="service-option ${state.bookingChargeEv ? "active" : ""}">
                  <span style="display:flex;align-items:center;gap:8px;">
                    <input id="bk-ev-chk" type="checkbox" ${state.bookingChargeEv ? "checked" : ""} />
                    <span><strong>EV Charging Power</strong> <span class="sub-txt">(${inr(activeEvRate)}/kWh only)</span></span>
                  </span>
                  <strong class="mono-val">${state.bookingChargeEv ? inr(quote.ev_fee) : "₹0"}</strong>
                </label>`
              : `<div class="service-option active" style="cursor:default;">
                  <span><strong>Parking Space Fee Only</strong> <span class="sub-txt">(${quoteCategory === "two_wheeler" ? "Scooty/Bike" : "Car"} · ${inr(activeHourly)}/hr)</span></span>
                  <strong class="mono-val">${inr(quote.base_fee)}</strong>
                </div>`
          }
        </div>

        <div>
          <div class="sub-txt" style="margin-bottom:6px;font-weight:600;">Payment Method</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
            ${PAYMENT_OPTIONS.map(
              (p) => `
              <button class="slot-btn ${state.bookingPayMethod === p.id ? "selected" : ""}" data-bk-pay="${esc(p.id)}">
                <strong style="font-size:12px;">${p.label}</strong>
                <span class="sub-txt">${p.sub}</span>
              </button>`
            ).join("")}
          </div>
        </div>

        <div style="display:flex;gap:6px;">
          <input id="bk-promo-inp" class="input-box" value="${esc(state.bookingPromo)}" placeholder="Optional Promo: HYDEV20, FASTAG15, NAMASTE50" />
          <button id="bk-promo-btn" class="btn btn-soft btn-sm">Apply</button>
        </div>
        ${quote.promo_note ? `<div class="sub-txt" style="color:var(--accent-teal);">${esc(quote.promo_note)}</div>` : ""}

        <div class="card soft" style="padding:12px;">
          ${includeParking ? `<div class="row-between"><span>${quoteCategory === "two_wheeler" ? "2W Parking" : "Car Parking"} (${state.bookingDuration / 60}h × ${inr(activeHourly)})</span><strong class="mono-val">${inr(quote.base_fee)}</strong></div>` : ""}
          ${chargeEv ? `<div class="row-between"><span>EV Charging Only (${quoteCategory === "two_wheeler" ? "2W AC" : "Car EV"})</span><strong class="mono-val">${inr(quote.ev_fee)}</strong></div>` : ""}
          ${quote.discount_amount > 0 ? `<div class="row-between" style="color:var(--accent-teal);"><span>Discount</span><strong class="mono-val">-${inr(quote.discount_amount)}</strong></div>` : ""}
          <div class="row-between" style="border-top:1px solid var(--border);padding-top:6px;font-size:15px;">
            <strong>Total</strong><strong class="mono-val">${inr(quote.total_fee)}</strong>
          </div>
        </div>

        <button id="bk-confirm" class="btn btn-primary btn-full" ${chosenSlot && veh ? "" : "disabled"}>
          ${!veh ? "Add Vehicle in Vehicles First" : chosenSlot ? `Confirm & Pay ${inr(quote.total_fee)} (${chosenSlot.level}-${chosenSlot.code})` : "Select an Open Slot"}
        </button>
      </div>
    </div>
  `;

  el.querySelector("#bk-back").addEventListener("click", () => goToPage("hubs"));
  el.querySelector("#bk-go-vehicles")?.addEventListener("click", () => goToPage("vehicles"));
  el.querySelector("#bk-tab-reg").addEventListener("click", () => {
    state.slotZoneTab = "regular_deck";
    state.evVehicleType = "car";
    state.bookingIncludeParking = true;
    state.bookingChargeEv = false;
    state.selectedSlotId = null;
    renderBookPage();
  });
  el.querySelector("#bk-tab-2w").addEventListener("click", () => {
    state.slotZoneTab = "two_wheeler_floor";
    state.evVehicleType = "two_wheeler";
    state.bookingIncludeParking = true;
    state.bookingChargeEv = false;
    state.selectedSlotId = null;
    renderBookPage();
  });
  el.querySelector("#bk-tab-ev").addEventListener("click", () => {
    state.slotZoneTab = "ev_charging_zone";
    const v = activeVehicle();
    state.evVehicleType = v && (v.car_type === "Scooty / Bike" || v.car_type === "EV Scooty") ? "two_wheeler" : "car";
    state.bookingIncludeParking = false;
    state.bookingChargeEv = true;
    state.selectedSlotId = null;
    renderBookPage();
  });
  el.querySelectorAll("[data-lvl]").forEach((b) =>
    b.addEventListener("click", () => {
      state.activeRegularLevel = b.dataset.lvl;
      state.selectedSlotId = null;
      renderBookPage();
    })
  );
  el.querySelectorAll("[data-bk-slot]").forEach((b) =>
    b.addEventListener("click", () => {
      state.selectedSlotId = b.dataset.bkSlot;
      const picked = d.bays.find((x) => x.id === state.selectedSlotId);
      if (picked) state.evVehicleType = picked.vehicle_category || "car";
      renderBookPage();
    })
  );
  el.querySelectorAll("[data-bk-dur]").forEach((b) =>
    b.addEventListener("click", () => {
      state.bookingDuration = Number(b.dataset.bkDur);
      renderBookPage();
    })
  );
  el.querySelector("#bk-park-chk")?.addEventListener("change", (e) => {
    if (!e.target.checked && !state.bookingChargeEv) {
      state.bookingChargeEv = true;
      toast("Kept EV Charging active (select at least one service)");
    }
    state.bookingIncludeParking = e.target.checked;
    renderBookPage();
  });
  el.querySelector("#bk-ev-chk")?.addEventListener("change", (e) => {
    if (!e.target.checked && !state.bookingIncludeParking) {
      state.bookingIncludeParking = true;
      toast("Kept Parking Space active (select at least one service)");
    }
    state.bookingChargeEv = e.target.checked;
    renderBookPage();
  });
  el.querySelectorAll("[data-bk-pay]").forEach((b) =>
    b.addEventListener("click", () => {
      state.bookingPayMethod = b.dataset.bkPay;
      renderBookPage();
    })
  );
  el.querySelector("#bk-promo-btn").addEventListener("click", () => {
    state.bookingPromo = el.querySelector("#bk-promo-inp").value.trim();
    renderBookPage();
  });
  el.querySelector("#bk-confirm").addEventListener("click", async () => {
    if (!chosenSlot || !veh) return;
    try {
      const res = await api.createReservation({
        facility_id: d.id,
        bay_id: chosenSlot.id,
        vehicle_id: veh.id,
        duration_mins: state.bookingDuration,
        payment_method: state.bookingPayMethod,
        promo_code: state.bookingPromo || null,
        include_parking: chosenSlot.zone === "ev_charging_zone" ? state.bookingIncludeParking : true,
        charge_ev: state.bookingChargeEv && chosenSlot.zone === "ev_charging_zone",
      });
      toast(`Booked ${res.level}-${res.bay_code} · PIN: ${res.gate_pin}`);
      await refreshData();
      await startNavigation(d.id, false);
      goToPage("ticket");
    } catch (err) {
      toast(err.message);
    }
  });
}

function renderTicketPage() {
  const el = document.getElementById("ticket-page-content");
  const active = state.reservations.find((r) => r.status === "reserved" || r.status === "entered");
  if (!active) {
    el.innerHTML = `
      <div class="card">
        <div class="phase-title-wrap">
          <span class="phase-eyebrow phase-ticket-tag">Gate Boarding Pass</span>
          <h2 class="page-heading">My Ticket</h2>
        </div>
        <p class="sub-txt">No active booking right now. Select an area on Map or Parking Hubs.</p>
        <div style="display:flex;gap:8px;">
          <button id="tk-map" class="btn btn-primary btn-sm">Open Map</button>
          <button id="tk-hubs" class="btn btn-soft btn-sm">Parking Hubs</button>
        </div>
      </div>`;
    el.querySelector("#tk-map").addEventListener("click", () => goToPage("map"));
    el.querySelector("#tk-hubs").addEventListener("click", () => goToPage("hubs"));
    return;
  }

  const isEntered = active.status === "entered";
  el.innerHTML = `
    <div class="phase-title-wrap">
      <span class="phase-eyebrow phase-ticket-tag">Gate Boarding Pass</span>
      <h2 class="page-heading">My Ticket</h2>
    </div>
    <div class="two-col">
      <div class="card" style="border-color:var(--border-strong);">
        <div class="row-between">
          <div>
            <span style="font-family:var(--font-mono);font-size:12px;font-weight:700;color:var(--accent-teal);">${isEntered ? "PARKED INSIDE · LIVE" : "RESERVED · READY AT GATE"}</span>
            <h3 class="area-name" style="font-size:18px;margin-top:4px;">${esc(active.facility_name)}</h3>
            <div class="sub-txt">Slot <strong class="slot-code-txt">${active.level}-${active.bay_code}</strong> · ${esc(active.vehicle_name)} <span class="plate-badge">${esc(active.vehicle_plate)}</span></div>
          </div>
          <span class="rate-pill">${inr(active.total_fee)}</span>
        </div>

        <div class="qr-card">
          ${generateQrSvg(active.qr_payload, 88)}
          <div>
            <div style="font-size:12px;font-weight:700;color:var(--accent-gold);">GATE ENTRY PIN</div>
            <div class="pin-big">${active.gate_pin}</div>
            <div style="font-size:12px;color:var(--text-sub);">Booking ${active.id} · ${active.duration_mins} mins</div>
          </div>
        </div>

        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          ${
            !isEntered
              ? `<button id="tk-enter" class="btn btn-teal" style="flex:1;">Enter Gate</button>
                 <button id="tk-cancel" class="btn btn-danger">Cancel</button>`
              : `<button id="tk-ext" class="btn btn-soft">+30m Extend</button>
                 <button id="tk-exit" class="btn btn-primary" style="flex:1;">Exit Parking &amp; Free Slot</button>`
          }
        </div>
      </div>

      <div class="card">
        <strong style="font-size:15px;">Directions</strong>
        <p class="sub-txt">${esc(active.entrance_hint)}</p>
        ${
          state.activeRoute
            ? state.activeRoute.steps
                .map(
                  (s, idx) => `
                  <div class="card soft" style="padding:10px;">
                    <strong>Step ${idx + 1}: ${esc(s.instruction)}</strong>
                    <span class="sub-txt">${s.distance_m}m</span>
                  </div>`
                )
                .join("")
            : ""
        }
        <button id="tk-show-map" class="btn btn-soft btn-sm">View Route on Map</button>
      </div>
    </div>
  `;

  el.querySelector("#tk-enter")?.addEventListener("click", async () => {
    await api.enterReservation(active.id);
    toast("Gate opened — slot marked occupied");
    await refreshData();
  });
  el.querySelector("#tk-cancel")?.addEventListener("click", async () => {
    await api.cancelReservation(active.id);
    toast("Booking cancelled and slot released");
    await refreshData();
  });
  el.querySelector("#tk-ext")?.addEventListener("click", async () => {
    await api.extendReservation(active.id, 30, active.payment_method);
    toast("Extended +30 mins");
    await refreshData();
  });
  el.querySelector("#tk-exit")?.addEventListener("click", async () => {
    await api.exitReservation(active.id);
    toast("Exited — slot freed");
    await refreshData();
    goToPage("history");
  });
  el.querySelector("#tk-show-map").addEventListener("click", () => startNavigation(active.facility_id, true));
}

function renderVehiclesPage() {
  const el = document.getElementById("vehicles-page-content");
  const catalog = VEHICLE_CATALOG[state.carFormTab] || VEHICLE_CATALOG.car;
  const selectedPreset = catalog[state.carFormSelectedIdx] || catalog[0];

  el.innerHTML = `
    <div class="phase-title-wrap">
      <span class="phase-eyebrow phase-garage-tag">Personal Garage</span>
      <h2 class="page-heading">Vehicles</h2>
    </div>
    <div class="two-col">
      <div class="card">
        <strong style="font-size:15px;">Saved Vehicles</strong>
        ${
          state.vehicles.length
            ? state.vehicles
                .map(
                  (v) => `
                  <div class="card soft" style="${v.id === state.activeVehicleId ? "border-color:var(--accent-saffron);" : ""}">
                    <div class="row-between">
                      <div>
                        <strong>${esc(v.make_model)}</strong>
                        <div class="sub-txt"><span class="plate-badge">${esc(v.plate)}</span> · ${esc(v.car_type)} ${v.is_ev ? `· ${esc(v.charger_pref)}` : ""}</div>
                      </div>
                      <div style="display:flex;gap:6px;">
                        <button class="btn btn-sm ${v.id === state.activeVehicleId ? "btn-primary" : "btn-soft"}" data-use-c="${v.id}">${v.id === state.activeVehicleId ? "Active" : "Select"}</button>
                        <button class="btn btn-danger btn-sm" data-del-c="${v.id}">✕</button>
                      </div>
                    </div>
                  </div>`
                )
                .join("")
            : `<p class="sub-txt">No vehicles saved yet. Use Add Vehicle on the right to choose yours.</p>`
        }
      </div>

      <div class="card">
        <strong style="font-size:15px;">Add Vehicle</strong>
        <p class="sub-txt">Choose Car, EV Car, Scooty / Bike, or EV Scooty / Bike and enter your plate.</p>

        <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:6px;">
          <button type="button" class="btn btn-sm ${state.carFormTab === "car" ? "btn-primary" : "btn-soft"}" data-vt="car">Car</button>
          <button type="button" class="btn btn-sm ${state.carFormTab === "ev_car" ? "btn-teal" : "btn-soft"}" data-vt="ev_car">EV Car</button>
          <button type="button" class="btn btn-sm ${state.carFormTab === "two_wheeler" ? "btn-primary" : "btn-soft"}" data-vt="two_wheeler">Scooty / Bike</button>
          <button type="button" class="btn btn-sm ${state.carFormTab === "ev_two_wheeler" ? "btn-teal" : "btn-soft"}" data-vt="ev_two_wheeler">EV Scooty / Bike</button>
        </div>

        <div style="display:flex;flex-wrap:wrap;gap:6px;">
          ${catalog
            .map(
              (c, idx) => `
              <button type="button" class="ctrl-btn ${state.carFormSelectedIdx === idx && !state.carFormCustom ? "active" : ""}" data-cat-i="${idx}">
                ${esc(c.make_model)}
              </button>`
            )
            .join("")}
        </div>
        <input id="car-cust-inp" class="input-box" value="${esc(state.carFormCustom)}" placeholder="Or type another model (optional)" />
        <input id="car-plate-inp" class="input-box" value="${esc(state.carFormPlate)}" placeholder="Registration Plate (e.g. TS 09 EV 4021)" />
        <button id="car-add-btn" class="btn btn-primary btn-full">Add Vehicle</button>
      </div>
    </div>
  `;

  el.querySelectorAll("[data-use-c]").forEach((b) =>
    b.addEventListener("click", () => {
      state.activeVehicleId = b.dataset.useC;
      toast("Active vehicle updated");
      renderVehiclesPage();
    })
  );
  el.querySelectorAll("[data-del-c]").forEach((b) =>
    b.addEventListener("click", async () => {
      await api.deleteVehicle(b.dataset.delC);
      await refreshData();
    })
  );
  el.querySelectorAll("[data-vt]").forEach((b) =>
    b.addEventListener("click", () => {
      state.carFormTab = b.dataset.vt;
      state.carFormSelectedIdx = 0;
      state.carFormCustom = "";
      renderVehiclesPage();
    })
  );
  el.querySelectorAll("[data-cat-i]").forEach((b) =>
    b.addEventListener("click", () => {
      state.carFormSelectedIdx = Number(b.dataset.catI);
      state.carFormCustom = "";
      renderVehiclesPage();
    })
  );
  el.querySelector("#car-cust-inp").addEventListener("input", (e) => {
    state.carFormCustom = e.target.value;
  });
  el.querySelector("#car-plate-inp").addEventListener("input", (e) => {
    state.carFormPlate = e.target.value;
  });
  el.querySelector("#car-add-btn").addEventListener("click", async () => {
    const model = state.carFormCustom.trim() || selectedPreset.make_model;
    const plate = (state.carFormPlate.trim() || "TS 09 AB 1234").toUpperCase();
    const created = await api.createVehicle({
      make_model: model,
      plate,
      car_type: selectedPreset.car_type,
      is_ev: selectedPreset.is_ev,
      charger_pref: selectedPreset.charger_pref,
    });
    state.activeVehicleId = created.id;
    state.carFormCustom = "";
    toast(`Added ${created.make_model}`);
    await refreshData();
  });
}

function renderHistoryPage() {
  const el = document.getElementById("history-page-content");
  el.innerHTML = `
    <div class="row-between" style="align-items:center;">
      <div class="phase-title-wrap">
        <span class="phase-eyebrow phase-history-tag">Ledger &amp; Receipts</span>
        <h2 class="page-heading">History (${state.reservations.length})</h2>
      </div>
    </div>
    <div class="bento-grid">
      ${
        state.reservations.length
          ? state.reservations
              .map(
                (r) => `
                <article class="card">
                  <div class="row-between">
                    <div>
                      <strong class="area-name">${esc(r.facility_name)}</strong>
                      <div class="sub-txt"><span class="mono-val">${r.id} · ${r.level}-${r.bay_code}</span> · <span class="plate-badge">${esc(r.vehicle_plate)}</span></div>
                      <div class="sub-txt">${formatTime(r.start_time)} · ${esc(r.payment_method)}</div>
                    </div>
                    <div style="text-align:right;">
                      <span class="rate-pill">${inr(r.total_fee)}</span>
                      <div class="sub-txt mono-val" style="margin-top:4px;font-weight:700;">${r.status.toUpperCase()}</div>
                    </div>
                  </div>
                </article>`
              )
              .join("")
          : `<div class="card"><p class="sub-txt">No booking records yet.</p></div>`
      }
    </div>
  `;
}

function renderAlertsPage() {
  const el = document.getElementById("alerts-page-content");
  el.innerHTML = `
    <div class="row-between" style="align-items:center;">
      <div class="phase-title-wrap">
        <span class="phase-eyebrow phase-alerts-tag">Sensor &amp; Spot Stream</span>
        <h2 class="page-heading">Alerts (${state.notifications.length})</h2>
      </div>
      <button id="btn-mark-read-all" class="btn btn-soft btn-sm">Mark all read</button>
    </div>
    <div class="bento-grid">
      ${
        state.notifications.length
          ? state.notifications
              .map(
                (n) => `
                <article class="card">
                  <div class="row-between">
                    <strong>${esc(n.title)}</strong>
                    <span class="sub-txt">${formatTime(n.created_at)}</span>
                  </div>
                  <p class="sub-txt">${esc(n.message)}</p>
                </article>`
              )
              .join("")
          : `<div class="card"><p class="sub-txt">You are all caught up.</p></div>`
      }
    </div>
  `;

  el.querySelector("#btn-mark-read-all").addEventListener("click", async () => {
    await api.markNotificationsRead();
    state.notifications.forEach((n) => (n.is_read = true));
    updateAlertBadge();
    renderAlertsPage();
  });
}

function renderActivePage() {
  if (state.page === "map") renderMapPage();
  else if (state.page === "hubs") renderHubsPage();
  else if (state.page === "book") renderBookPage();
  else if (state.page === "ticket") renderTicketPage();
  else if (state.page === "vehicles") renderVehiclesPage();
  else if (state.page === "history") renderHistoryPage();
  else if (state.page === "alerts") renderAlertsPage();
}

async function startNavigation(facId, switchToMap = true) {
  stopNavigation();
  state.activeRoute = await api.getNavigation(facId, state.userPos.x, state.userPos.y);
  state.navStepIndex = 0;
  if (switchToMap) goToPage("map");
  else renderMapPage();
}

function stopNavigation() {
  clearInterval(state.navTimer);
  state.navTimer = null;
  state.navSimPos = null;
  state.activeRoute = null;
  renderMapPage();
}

function startDriveSimulation() {
  if (!state.activeRoute || state.navTimer) return;
  const pts = state.activeRoute.path_points;
  let seg = 0;
  let t = 0;
  state.navSimPos = { x: pts[0][0], y: pts[0][1] };

  state.navTimer = setInterval(() => {
    t += 0.14;
    if (t >= 1) {
      t = 0;
      seg += 1;
      state.navStepIndex = Math.min(state.activeRoute.steps.length - 1, seg);
    }
    if (seg >= pts.length - 1) {
      const endPt = pts[pts.length - 1];
      state.locationEnabled = true;
      state.userPos = { x: endPt[0], y: endPt[1] };
      stopNavigation();
      toast("Arrived at gate — open My Ticket to enter");
      refreshData();
      goToPage("ticket");
      return;
    }
    const a = pts[seg];
    const b = pts[seg + 1];
    state.navSimPos = { x: a[0] + (b[0] - a[0]) * t, y: a[1] + (b[1] - a[1]) * t };
    renderMapPage();
  }, 150);
}

async function handleMapSearch(query) {
  const dd = document.getElementById("map-search-dropdown");
  document.getElementById("btn-search-clear").classList.toggle("hidden", !query);
  if (!query.trim()) {
    dd.classList.add("hidden");
    return;
  }

  const ux = state.locationEnabled ? state.userPos.x : null;
  const uy = state.locationEnabled ? state.userPos.y : null;
  const [dests, hubs] = await Promise.all([
    api.getDestinations(query, ux, uy),
    api.getFacilities({ user_x: state.userPos.x, user_y: state.userPos.y, search: query }),
  ]);

  const rows = [
    ...dests.slice(0, 6).map(
      (d) => `
      <button class="dd-item" data-dd-dest="${d.id}">
        <div>
          <strong>${esc(d.name)}</strong>
          <div class="sub-txt">${esc(d.area)} · ${esc(d.category)}</div>
        </div>
        <span class="sub-txt">${d.distance_km !== null ? `${d.distance_km} km (${d.bearing}) · ${d.drive_mins}m` : "Nearby parking"}</span>
      </button>`
    ),
    ...hubs.slice(0, 5).map(
      (f) => `
      <button class="dd-item" data-dd-hub="${f.id}">
        <div>
          <strong>${esc(f.name)}</strong> <span class="area-tag">${esc(f.tag)}</span>
          <div class="sub-txt">Car ${f.regular_available} · 2W ${f.two_wheeler_available} · EV ${f.ev_available} ${state.locationEnabled ? `· ${f.distance_km} km` : ""}</div>
        </div>
        <span class="rate-pill">${inr(f.hourly_rate)}/hr</span>
      </button>`
    ),
  ];

  dd.innerHTML = rows.length ? rows.join("") : `<div class="sub-txt" style="padding:10px;">No places or areas found</div>`;
  dd.classList.remove("hidden");

  dd.querySelectorAll("[data-dd-dest]").forEach((b) =>
    b.addEventListener("click", async () => {
      const dest = dests.find((x) => x.id === b.dataset.ddDest);
      if (!dest) return;
      state.activeDestination = dest;
      state.selectedFacilityDetail = null;
      state.selectedFacilityId = null;
      dd.classList.add("hidden");
      document.getElementById("map-search").value = dest.name;
      await refreshData();
      if (state.facilities.length > 0) {
        state.selectedFacilityId = state.facilities[0].id;
        state.selectedFacilityDetail = await api.getFacility(state.selectedFacilityId, state.userPos.x, state.userPos.y);
        renderMapPage();
      }
      toast(`Nearest parking to ${dest.name}`);
    })
  );

  dd.querySelectorAll("[data-dd-hub]").forEach((b) =>
    b.addEventListener("click", async () => {
      dd.classList.add("hidden");
      state.selectedFacilityId = b.dataset.ddHub;
      state.selectedFacilityDetail = await api.getFacility(state.selectedFacilityId, state.userPos.x, state.userPos.y);
      renderMapPage();
    })
  );
}

function bindEvents() {
  document.querySelectorAll(".rail-item, .m-btn").forEach((b) => b.addEventListener("click", () => goToPage(b.dataset.page)));

  // Shared theme toggle logic
  function applyTheme(theme) {
    const html = document.documentElement;
    html.dataset.theme = theme;
    localStorage.setItem("dp-theme", theme);
    const isDark = theme === "dark";
    const topLabel = document.getElementById("theme-top-label");
    const sideLabel = document.getElementById("theme-label");
    if (topLabel) topLabel.textContent = isDark ? "Light Mode" : "Dark Mode";
    if (sideLabel) sideLabel.textContent = isDark ? "Light Classic" : "Dark Mode";
  }

  function toggleTheme() {
    const html = document.documentElement;
    const next = html.dataset.theme === "light" ? "dark" : "light";
    applyTheme(next);
  }

  // Top-right toggle
  const btnThemeTop = document.getElementById("btn-theme-top");
  if (btnThemeTop) btnThemeTop.addEventListener("click", toggleTheme);

  // Sidebar toggle (original)
  const btnThemeSide = document.getElementById("btn-theme");
  if (btnThemeSide) btnThemeSide.addEventListener("click", toggleTheme);

  // Restore saved theme
  const savedTheme = localStorage.getItem("dp-theme");
  if (savedTheme) applyTheme(savedTheme);

  document.getElementById("btn-location-toggle").addEventListener("click", toggleUserLocation);
  document.getElementById("map-radius-select").addEventListener("change", (e) => {
    state.maxDistanceKm = e.target.value;
    refreshData();
  });

  document.getElementById("map-vehicle-filter").addEventListener("change", (e) => {
    state.floorFilter = e.target.value;
    refreshData();
  });

  const searchInp = document.getElementById("map-search");
  searchInp.addEventListener("input", (e) => handleMapSearch(e.target.value));
  document.getElementById("btn-search-clear").addEventListener("click", async () => {
    searchInp.value = "";
    state.activeDestination = null;
    document.getElementById("map-search-dropdown").classList.add("hidden");
    document.getElementById("btn-search-clear").classList.add("hidden");
    await refreshData();
  });

  document.getElementById("btn-pulse-map").addEventListener("click", async () => {
    const res = await api.pulseSensor();
    if (res?.summary) {
      document.getElementById("map-live-text").textContent = res.summary;
      toast(res.summary);
    }
    await refreshData();
  });

  document.getElementById("btn-zoom-in").addEventListener("click", () => {
    state.mapView.zoom = Math.min(2.8, +(state.mapView.zoom + 0.3).toFixed(2));
    renderMapPage();
  });
  document.getElementById("btn-zoom-out").addEventListener("click", () => {
    state.mapView.zoom = Math.max(1, +(state.mapView.zoom - 0.3).toFixed(2));
    renderMapPage();
  });
  document.getElementById("btn-zoom-reset").addEventListener("click", () => {
    state.mapView = { zoom: 1.0, panX: 0, panY: 0 };
    renderMapPage();
  });

  const compassBtn = document.getElementById("map-compass");
  if (compassBtn) {
    compassBtn.addEventListener("click", () => {
      state.mapView = { zoom: 1.0, panX: 0, panY: 0 };
      toast("Map oriented to True North (0°)");
      renderMapPage();
    });
  }

  const legendCollapseBtn = document.getElementById("btn-legend-collapse");
  if (legendCollapseBtn) {
    legendCollapseBtn.addEventListener("click", () => {
      const body = document.getElementById("legend-body");
      if (body) {
        const isCollapsed = body.classList.toggle("collapsed");
        legendCollapseBtn.textContent = isCollapsed ? "+" : "−";
        legendCollapseBtn.title = isCollapsed ? "Expand grid key" : "Collapse grid key";
      }
    });
  }

  const svg = document.getElementById("hyderabad-svg");
  let dragging = false;
  let moved = false;
  let sx = 0;
  let sy = 0;

  svg.addEventListener("pointerdown", (e) => {
    dragging = true;
    moved = false;
    sx = e.clientX;
    sy = e.clientY;
  });
  window.addEventListener("pointermove", (e) => {
    if (!dragging) return;
    const dx = e.clientX - sx;
    const dy = e.clientY - sy;
    if (Math.hypot(dx, dy) > 4) moved = true;
    if (moved && state.mapView.zoom > 1) {
      state.mapView.panX = Math.max(-350, Math.min(350, state.mapView.panX - dx * 0.6));
      state.mapView.panY = Math.max(-250, Math.min(250, state.mapView.panY - dy * 0.6));
      sx = e.clientX;
      sy = e.clientY;
      renderMapPage();
    }
  });
  window.addEventListener("pointerup", (e) => {
    if (!dragging) return;
    dragging = false;
    if (!moved && e.target.closest("#hyderabad-svg") && state.locationEnabled) {
      const ctm = svg.getScreenCTM();
      if (!ctm) return;
      const pt = svg.createSVGPoint();
      pt.x = e.clientX;
      pt.y = e.clientY;
      const p = pt.matrixTransform(ctm.inverse());
      state.userPos = { x: Math.max(40, Math.min(960, Math.round(p.x))), y: Math.max(40, Math.min(660, Math.round(p.y))) };
      refreshData();
    }
  });
}

async function bootstrap() {
  bindEvents();
  await refreshData();
  connectLiveSocket(async (msg) => {
    if (msg.summary) document.getElementById("map-live-text").textContent = msg.summary;
    await refreshData();
  });
}

bootstrap();
