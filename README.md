# DeccanPulse — Hyderabad Smart Parking

**DeccanPulse** is an organized multi-page smart parking application for **Greater Hyderabad** supporting **Car Decks**, **Scooty & Bike Floors**, and **Separate EV Charger Wings** across 21 city areas with Metro line integration, priced strictly in multiples of ₹5.

## Pages

- **Map**: Uncluttered North-Up Hyderabad map displaying parking symbols and area names alongside Red, Blue & Green Metro rail lines, while the side panel lists complete area details (Metro line colors, amenities, EV chargers, bay counts, rates, and live distances).
- **Parking Hubs**: Organized directory of all 21 Hyderabad areas (including Miyapur, Kompally, ECIL, Uppal, MGBS, Narsingi, Mehdipatnam, Nagole & Shamshabad Airport) with Car, Scooty/Bike, EV Car/2W availability, amenities, and Metro line color badges.
- **Book Slot**: 3-tab zone selector (`Cars`, `Scooty/Bike`, `EV Wing` split into **EV Car Chargers** & **EV Scooty/Bike Chargers**) with independent service toggles (`Parking Space Fee` and `EV Charging Power` selectable separately so charging-only sessions exclude parking fees).
- **My Ticket**: Gate QR/PIN pass with `Reserved → Enter Gate → Exit Parking` lifecycle and road directions.
- **Vehicles**: Dedicated vehicle management page (`Add Vehicle` for Cars, EV Cars, Scooty/Bike, and EV Scooty/Bike).
- **History** & **Alerts**: Separate dedicated pages for booking records and live alerts.

## Running Locally

```powershell
cd C:\Portfolio\smart_parking_app
C:\Portfolio\python_env\tools\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8010
```

## Running Tests

```powershell
C:\Portfolio\python_env\tools\python.exe -m unittest discover -s tests -v
```
