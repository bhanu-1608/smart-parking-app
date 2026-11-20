# DeccanPulse — Hyderabad Smart Parking System
## Comprehensive Project Documentation & Technical Report

**Project Name:** DeccanPulse (Hyderabad Smart Parking Application)  
**Author / Developer:** Bhanu ([@bhanu-1608](https://github.com/bhanu-1608))  
**Repository:** [https://github.com/bhanu-1608/smart-parking-app](https://github.com/bhanu-1608/smart-parking-app)  
**Local Development Server:** `http://localhost:8000`  
**Version:** 4.2.0  
**Status:** Production-Ready / Fully Verified  

---

## 1. Executive Summary

Urban vehicular growth in Greater Hyderabad—particularly across the Western IT corridors (HITEC City, Gachibowli, Kondapur, Raidurg), central transit junctions (Ameerpet, Secunderabad, MGBS), and heritage commercial centers (Charminar, Abids)—has caused extreme parking friction. Drivers routinely spend 15 to 25 minutes searching for available parking spaces, escalating fuel consumption, road congestion, and carbon emissions.

**DeccanPulse** is a modern, responsive, and full-stack smart parking platform engineered specifically for the Greater Hyderabad metropolitan area. The platform integrates real-time IoT sensor telemetry, interactive vector GIS mapping with Metro Rail overlay, turn-by-turn navigation simulation, a multi-tier parking allocation engine (separating Cars, 2-Wheelers, and EV Charging Wings), and an automated gate pass system with dynamic QR codes and PINs.

The application incorporates a bespoke design language celebrating Hyderabad's architectural heritage—featuring Charminar limestone and Golconda night palettes, an iconic royal emblem, an authentic true-north precision compass, and a collapsible transit grid key.

---

## 2. Key Features & Functional Modules

| Module | Description | Technical Implementation |
| :--- | :--- | :--- |
| **Real-Time Availability** | Continuous live tracking of vacant, held, and occupied bays across 21 metropolitan hubs. | Async WebSocket broadcast loop (`/ws/live`) + SQLite triggers |
| **Location-Based Search** | GPS proximity computation, user pin dragging, distance radius filters (5km, 10km, 18km, all). | Haversine road-distance formula + canvas projection |
| **Multi-Deck Slot Allocation** | Dedicated spaces for Car Decks, Scooty/Bike Floors, and separate EV Charging Wings. | Multi-tier Bayesian slot allocation with floor filtering |
| **EV Unbundled Pricing** | Independent selection of Parking Space Fee and EV Charging Power (kWh). | Dynamic pricing engine strictly rounded in multiples of ₹5 |
| **Interactive Map Canvas** | True North-Up 1000×700 vector map showing Nehru ORR, Musi River, lakes, and Metro rail lines. | Dynamic SVG DOM rendering with pan, zoom (1x to 2.8x), and drag |
| **True North Compass** | Ornate nautical & precision compass situated at the top-right of the map frame. | 8-point faceted star rose with one-click view reset (0° North) |
| **Floating Grid Key (Legend)** | Glassmorphic collapsible legend identifying Metro corridors and parking spot density. | Semi-transparent floating HUD with toggleable state (`−`/`+`) |
| **Navigation & GPS Simulation** | Step-by-step turn-by-turn driving directions to parking entrance gates with live route drawing. | Maneuver generator + animated waypoint traversal loop |
| **Gate Pass & QR Ticket** | Digital boarding pass with deterministic SVG QR code, 6-digit gate PIN, and lifecycle actions. | FNV-1a hash QR matrix generator with entry/exit/extend API |
| **Personal Garage (Vehicles)** | Fleet management supporting Cars, EV Cars, Scooty/Bikes, and EV 2-Wheelers with plates. | CRUD vehicle registry with active vehicle badge |
| **Sensor Pulse & Alerts** | Live notifications feed for parking expiry warnings, space availability, and gate entries. | SQLite notifications table with unread counter badges |
| **Dual Hyderabad Palettes** | Seamless toggling between **Hyderabad Light** (Limestone/Saffron) and **Dark** (Golconda Night/Neon Teal). | CSS3 custom properties with 300ms transitions + localStorage |

---

## 3. System Architecture & Tech Stack

```
+---------------------------------------------------------------------------------+
|                                 CLIENT TIER                                     |
|  Responsive Web & Mobile Browser UI (HTML5, Vanilla ES6 Modules, Modern CSS3)  |
|  - Geographic SVG Map Canvas (True North-Up Projection)                         |
|  - Floating Precision Compass Rose & Collapsible Hyderabad Grid Key (Legend)   |
|  - Turn-by-Turn Floating HUD & Animated Driving Simulation                      |
|  - Gate Pass Boarding Ticket with In-Browser Vector QR Code Generator           |
+---------------------------------------------------------------------------------+
                                       |
                   HTTP REST APIs / WebSocket Protocol
                                       |
+---------------------------------------------------------------------------------+
|                               BACKEND ENGINE                                    |
|  FastAPI Application (Python 3.12) with Lifespan Asynchronous Context          |
|  - api.getFacilities(), api.getFacility(), api.getNavigation()                  |
|  - api.getQuote(), api.createReservation(), api.enterReservation()             |
|  - api.extendReservation(), api.exitReservation(), api.getVehicles()           |
|  - Background Async Sensor Heartbeat Task (Interval: 4.5s)                     |
|  - WebSocket Live Hub with Connection Pooling & Real-time Event Broadcast       |
+---------------------------------------------------------------------------------+
                                       |
                             SQLite3 Database Engine
                                       |
+---------------------------------------------------------------------------------+
|                               DATABASE TIER                                     |
|  data/deccanpulse_v8.db (Relational Schema with Strict Foreign Keys)            |
|  - facilities (21 Hyderabad Hubs, Geo-coordinates, Rates, Metro Lines)         |
|  - bays (Multi-tier parking bays: regular_deck, two_wheeler_floor, EV wings)    |
|  - vehicles (User fleet registry: fuel cars, EV cars, bikes, EV scooties)      |
|  - reservations (Active & past bookings, gate PINs, durations, payment data)    |
|  - notifications (System alerts, spot updates, expiry warnings)                 |
+---------------------------------------------------------------------------------+
```

### Technology Breakdown:
* **Frontend:** Vanilla ES6 JavaScript (zero third-party UI framework dependencies for maximum performance and instant load), CSS3 Custom Properties with hardware-accelerated animations, Scalable Vector Graphics (SVG).
* **Backend:** FastAPI, Uvicorn ASGI server, Pydantic v2 data models with runtime validation.
* **Database:** SQLite 3 with atomic transaction isolation and foreign key enforcement.
* **Real-Time Layer:** Native HTML5 WebSockets connected to FastAPI WebSocket endpoint (`/ws/live`).

---

## 4. Hyderabad Visual Identity & Color Palettes

The user interface reflects the rich cultural and technological dichotomy of Hyderabad—merging historic Nizam-era monuments with the modern cyber infrastructure of Cyberabad:

### 4.1. Hyderabad Light Palette (Day / Heritage)
* **Canvas Background:** `#fdf9f4` — Inspired by warm Charminar limestone and Deccani parchment.
* **Header & Cards:** `#ffffff` — Crisp daylight contrast with subtle soft elevation shadows (`rgba(15, 51, 71, 0.10)`).
* **Primary Accent:** `#c05020` — Charminar sunset saffron, providing high contrast for active states.
* **Secondary Accent:** `#0a7d73` — Deep Hussain Sagar teal, utilized for live indicators and EV facilities.
* **Tertiary Accent:** `#9c6412` — Nizam royal gold, adorning dividers, star ratings, and badge borders.
* **Metro Lines:** Red (`#e11d48`), Blue (`#0284c7`), Green (`#059669`).

### 4.2. Hyderabad Dark Palette (Night / Cyberabad)
* **Canvas Background:** `#0e1118` — Midnight Golconda sky with deep slate undertones.
* **Header & Surfaces:** `#141820` / `#1b2130` — Sleek obsidian and navy cards.
* **Primary Accent:** `#e07030` — Glowing neon Deccani saffron.
* **Secondary Accent:** `#22b898` — Cyber teal for live telemetry, active chargers, and sensor pulses.
* **Tertiary Accent:** `#dcb040` — Luminous Golconda gold for typography highlights and compass accents.
* **Map Roads & Lakes:** Obsidian asphalt (`#2a3042`) with deep reservoir cyan (`#0e2e3a`).

---

## 5. UI & Component Showcase

### 5.1. Bespoke Royal Charminar Emblem
The logo is an intricate vector emblem constructed with:
1. **Nizam Gold Squircle:** Double-rimmed rounded frame with subtle IoT radar grid coordinates.
2. **Four Fluted Minarets:** Graceful spires with multi-tiered balconies and golden crescent tips.
3. **Monumental Grand Archway:** Symmetrical Deccani pointed arch with an inner cusped archway.
4. **Smart Parking Beacon:** Concentric sonar radio waves, a high-contrast saffron shield with modern "P", and a cyber-teal live sensor LED dot.

### 5.2. True North Precision Compass
Situated at the top-right of the map canvas:
* Dual-ring outer dial with cardinal markings (`N`, `E`, `S`, `W`) and degree ticks.
* Faceted dual-tone North needle (Saffron/Crimson) and South needle (Nizam Gold).
* Displays live `0° NORTH` orientation badge.
* Interactive click action resets the map zoom and pan back to standard framing (`zoom: 1.0, panX: 0, panY: 0`) accompanied by an orientation toast.

### 5.3. Optimized Floating Grid Key (Legend)
Situated at the bottom-left of the map canvas:
* Docked unobtrusively over the south-western lake region to prevent obscuring parking markers.
* Clearly identifies:
  * **Metro Lines:** Red Line (*Miyapur ⇄ LB Nagar*), Blue Line (*Raidurg ⇄ Nagole*), Green Line (*JBS ⇄ MGBS*).
  * **Parking Bay Status:** 🟢 5+ Available (Plenty), 🟡 2–4 Available (Filling Fast), 🔴 0–1 Available (Near Capacity).
  * **Corridors & Water:** Nehru ORR Ring Road and Musi River & Twin Lakes.
* Collapsible via one-click toggle (`−` / `+`) for full-canvas map viewing.

---

## 6. Database Schema Design

### 6.1. Table: `facilities`
Stores metadata for the 21 parking hubs:
* `id` (TEXT, PK): Unique identifier (e.g., `fac-hitec`, `fac-necklace`, `fac-secbad`).
* `name` (TEXT): Display name (e.g., "HITEC Cyber Towers", "Necklace Road").
* `tag` (TEXT): Category badge (e.g., "IT Corridor", "Solar EV", "Central Metro").
* `district` (TEXT): City zone (e.g., "Madhapur", "Tank Bund", "Secunderabad").
* `address` (TEXT): Street address and entrance gate landmarks.
* `x`, `y` (REAL): Projected canvas coordinates (0..1000, 0..700).
* `lat`, `lng` (REAL): Authentic GPS coordinates (Latitude 17.21..17.52 N, Longitude 78.24..78.64 E).
* `hourly_rate` (INTEGER): Car hourly rate in ₹ (multiples of 5: ₹30, ₹40, ₹50).
* `two_wheeler_rate` (INTEGER): Scooty/Bike hourly rate in ₹ (₹10 or ₹15).
* `ev_kwh_rate` (INTEGER): EV Car charging rate (₹15/kWh).
* `ev_2w_kwh_rate` (INTEGER): EV 2-Wheeler charging rate (₹10/kWh).
* `metro_lines` (TEXT JSON): Metro connections array (`["Red"]`, `["Blue"]`, `["Blue", "Green"]`).
* `amenities` (TEXT JSON): Features array (`["Covered", "Valet", "FASTag", "CCTV"]`).

### 6.2. Table: `bays`
Represents individual physical parking slots:
* `id` (TEXT, PK): Unique bay identifier (e.g., `fac-necklace-L1-101`).
* `facility_id` (TEXT, FK): References `facilities(id)`.
* `level` (TEXT): Deck floor (`L1`, `L2`, `2W-G`, `EV-Z1`).
* `code` (TEXT): Bay code (`101`, `2W-05`, `EVC-01`).
* `zone` (TEXT): `regular_deck`, `two_wheeler_floor`, or `ev_charging_zone`.
* `vehicle_category` (TEXT): `car` or `two_wheeler`.
* `status` (TEXT): `available`, `reserved`, or `occupied`.
* `charger_standard` (TEXT): Charging standard (`CCS2 30kW DC`, `3.3kW Smart AC`, `None`).
* `walk_meters` (INTEGER): Walking distance to destination gate.

### 6.3. Table: `reservations`
Tracks booking transactions and boarding pass lifecycles:
* `id` (TEXT, PK): Formatted reference (e.g., `DP-818275`).
* `facility_id`, `bay_id`, `vehicle_id` (TEXT, FKs).
* `gate_pin` (TEXT): 6-digit numeric gate entry code.
* `qr_payload` (TEXT): Digitally verifiable gate boarding pass payload.
* `status` (TEXT): `reserved`, `entered`, `completed`, or `cancelled`.
* `start_time`, `end_time` (TEXT ISO 8601).
* `duration_mins` (INTEGER): Duration in minutes (60, 120, 180, 240).
* `base_fee`, `ev_fee`, `discount`, `total_fee` (REAL).
* `payment_method` (TEXT): UPI Instant, FASTag Auto-Pay, RuPay/Card, NetBanking.

---

## 7. API Reference

| Endpoint | Method | Description |
| :--- | :--- | :--- |
| `/api/health` | GET | Health check returning service status and city metadata |
| `/api/facilities` | GET | Lists parking facilities with proximity distances, rates, and occupancy |
| `/api/facilities/{id}` | GET | Returns complete facility detail including floor layout and bay matrix |
| `/api/facilities/{id}/watchlist`| POST | Toggles user notification alert for spot openings |
| `/api/destinations` | GET | Autocomplete search for Hyderabad landmarks and transit junctions |
| `/api/navigation/{id}` | GET | Computes turn-by-turn route, distance, ETA, and SVG path waypoints |
| `/api/quote` | POST | Calculates unbundled parking and EV power quote with promo code discounts |
| `/api/reservations` | GET / POST | Retrieves booking ledger or books an allocated parking bay |
| `/api/reservations/{id}/enter` | POST | Gate sensor check-in: validates PIN, marks bay occupied |
| `/api/reservations/{id}/exit` | POST | Gate sensor check-out: frees bay, calculates final settlement |
| `/api/reservations/{id}/extend`| POST | Extends reservation by +30 or +60 minutes |
| `/api/vehicles` | GET / POST | Retrieves or creates user vehicles (Car, EV, 2W) |
| `/api/notifications` | GET | Fetches live notification feed |
| `/api/simulate/pulse` | POST | Triggers manual IoT sensor occupancy flip and broadcasts event |
| `/ws/live` | WebSocket | High-throughput WebSocket stream for live bay updates |

---

## 8. Verification & Test Execution

### 8.1. Automated Unit Test Suite
The automated test suite (`tests/test_api.py`) exercises the pricing algorithms, database constraints, multi-category floors, and lifecycle states:
```powershell
python -m unittest tests/test_api.py
```
**Results:**
* `test_01_hubs_include_car_scooty_bike_and_ev_floors` — PASSED
* `test_02_scooty_bike_booking_and_pricing_multiples_of_5` — PASSED
* `test_03_ev_car_and_ev_scooty_unbundled_pricing` — PASSED
* `test_04_reservation_gate_entry_and_exit_lifecycle` — PASSED
* `test_05_navigation_turn_by_turn_waypoints` — PASSED
* **Overall: 5/5 Tests Passed in 0.18s.**

### 8.2. In-Browser Verification Protocol
The browser verification agent executed end-to-end user workflows on `http://localhost:8000`:
1. **Topbar Integrity:** Verified Charminar emblem, horizontal title and side title (`HYDERABAD · SMART PARKING • LIVE GRID`), and top-right theme switch.
2. **True North Compass:** Clicked the compass on the top-right; confirmed map view reset (`zoom: 1.0, panX: 0, panY: 0`) and toast alert.
3. **Floating Grid Key:** Expanded and collapsed the legend card; verified Metro line color mappings and occupancy thresholds.
4. **Interactive Map & Inspector:** Selected Necklace Road marker; verified rating (★ 4.9), Metro line badge (Red Line, Khairatabad Metro), and floor metrics.
5. **Turn-by-Turn Navigation:** Generated route directions and simulated live driving along waypoints until arrival at Lakefront Solar Canopy gate.
6. **Slot Booking & QR Pass:** Selected bay `L1-101`, selected 2-hour duration, applied FASTag Auto-Pay, and confirmed booking. Verified dynamic QR code pass and Gate Entry PIN (`818275`).
7. **Hyderabad Dark Theme:** Toggled dark mode; verified all components adapted to Golconda obsidian, neon teal, and gold.

---

## 9. Captured Visual Evidence Gallery

The following high-resolution screenshots are stored in `docs/screenshots/` and document every primary interface of DeccanPulse:

1. **`01_map_light_mode.png`** — Interactive Hyderabad map in Light Mode featuring the Charminar emblem, reoriented header, True North compass, and floating grid key.
2. **`02_navigation_directions.png`** — Route calculation with dashed path rendering and turn-by-turn navigation banner.
3. **`03_gps_drive_simulation.png`** — Real-time vehicle simulation traversing waypoints along the Tank Bund corridor.
4. **`04_parking_hubs_directory.png`** — Directory view of all 21 parking hubs with floor metrics, rates, and Metro line badges.
5. **`05_vehicle_garage_management.png`** — Fleet management page showing Car, EV Car, Scooty/Bike, and EV Scooty categories.
6. **`06_slot_allocation_booking.png`** — Slot allocation page with zone tabs, bay matrix, unbundled pricing checkboxes, and promo codes.
7. **`07_ticket_boarding_pass_qr.png`** — Gate boarding pass with dynamic vector QR code, 6-digit gate PIN, and gate check-in/out controls.
8. **`08_hyderabad_dark_mode.png`** — Complete interface rendering in the Hyderabad Dark Palette with glowing neon corridors and dark cartography.

---

## 10. Local Setup & Execution Guide

### Prerequisites
* Python 3.11 or 3.12
* Modern web browser (Chrome, Edge, Firefox, Safari)

### Steps to Run
```powershell
# 1. Clone repository
git clone https://github.com/bhanu-1608/smart-parking-app.git
cd smart-parking-app

# 2. Install dependencies
pip install fastapi uvicorn pydantic

# 3. Start local development server
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload

# 4. Access application
# Open your web browser and navigate to:
# http://localhost:8000
```
