from __future__ import annotations

import os
import random
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional

from .models import (
    BayOut,
    DestinationPoint,
    FacilityDetail,
    FacilitySummary,
    NotificationOut,
    ReservationCreate,
    ReservationOut,
    VehicleCreate,
    VehicleOut,
    calculate_price_quote,
    canvas_to_lat_lng,
    compute_bearing_label,
    compute_haversine_road_km,
    lat_lng_to_canvas,
    round_inr_5,
    utc_now_iso,
)

DEFAULT_DB_PATH = Path(__file__).resolve().parents[1] / "data" / "deccanpulse_v8.db"


HYDERABAD_DESTINATIONS = [
    ("dest-rgia", "RGIA International Airport", "Shamshabad", "Airport", 17.2403, 78.4294),
    ("dest-nagole", "Nagole Metro Terminus", "Nagole", "Metro", 17.3728, 78.5595),
    ("dest-miyapur", "Miyapur Metro Terminus", "Miyapur", "Metro", 17.4960, 78.3610),
    ("dest-kompally", "Kompally NH-44 Junction", "Kompally", "Hub", 17.4950, 78.4880),
    ("dest-ecil", "ECIL X Roads", "ECIL / Sainikpuri", "Hub", 17.4720, 78.5620),
    ("dest-narsingi", "Narsingi / Kokapet ORR", "Narsingi", "IT / ORR", 17.3920, 78.3340),
    ("dest-mehdipatnam", "Mehdipatnam PVNR Gate", "Mehdipatnam", "Transit", 17.3920, 78.4360),
    ("dest-ikea", "IKEA Hyderabad", "HITEC / Raidurg", "Shopping", 17.4384, 78.3748),
    ("dest-inorbit", "Inorbit Mall Cyberabad", "Mindspace", "Mall", 17.4342, 78.3868),
    ("dest-cable", "Durgam Cheruvu Cable Bridge", "Madhapur", "Landmark", 17.4310, 78.3895),
    ("dest-shilpa", "Shilparamam", "HITEC City", "Landmark", 17.4526, 78.3782),
    ("dest-sarath", "Sarath City Capital Mall", "Kondapur", "Mall", 17.4573, 78.3638),
    ("dest-gvk", "GVK One Mall", "Banjara Hills", "Mall", 17.4193, 78.4484),
    ("dest-kbr", "KBR National Park", "Jubilee Hills", "Park", 17.4202, 78.4216),
    ("dest-tankbund", "Tank Bund & Buddha Statue", "Hussain Sagar", "Landmark", 17.4156, 78.4750),
    ("dest-birla", "Birla Mandir", "Khairatabad", "Temple", 17.4062, 78.4691),
    ("dest-charminar", "Charminar & Laad Bazaar", "Old City", "Heritage", 17.3616, 78.4747),
    ("dest-golconda", "Golconda Fort", "Ibrahim Bagh", "Heritage", 17.3833, 78.4011),
    ("dest-secbad", "Secunderabad Railway Station", "Secunderabad", "Transit", 17.4344, 78.5013),
    ("dest-ameerpet", "Ameerpet Metro Interchange", "Ameerpet", "Metro", 17.4375, 78.4483),
    ("dest-mgbs", "MGBS Bus Terminal", "Imlibun", "Transit", 17.3787, 78.4810),
    ("dest-uppal", "Rajiv Gandhi Cricket Stadium", "Uppal", "Stadium", 17.4065, 78.5504),
    ("dest-kphb", "Nexus Mall / KPHB Metro", "Kukatpally", "Mall", 17.4840, 78.3888),
    ("dest-begumpet", "Begumpet Airport Terminal", "Begumpet", "Transit", 17.4531, 78.4675),
    ("dest-apollo", "Apollo Hospital", "Jubilee Hills", "Hospital", 17.4153, 78.4126),
    ("dest-lbnagar", "LB Nagar Metro Junction", "LB Nagar", "Metro", 17.3457, 78.5522),
    ("dest-dlf", "DLF Cyber City", "Gachibowli", "IT Park", 17.4483, 78.3552),
]


def get_db_path() -> Path:
    custom = os.environ.get("BAYPULSE_DB_PATH")
    return Path(custom) if custom else DEFAULT_DB_PATH


@contextmanager
def get_conn() -> Generator[sqlite3.Connection, None, None]:
    db_path = get_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    """Create schema if missing and seed 14 Hyderabad hubs with Car, Scooty/Bike & EV Charger floors."""
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS facilities (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                tag TEXT NOT NULL,
                district TEXT NOT NULL,
                address TEXT NOT NULL,
                x REAL NOT NULL,
                y REAL NOT NULL,
                lat REAL NOT NULL,
                lng REAL NOT NULL,
                hourly_rate INTEGER NOT NULL,
                two_wheeler_rate INTEGER NOT NULL,
                ev_kwh_rate INTEGER NOT NULL,
                ev_2w_kwh_rate INTEGER NOT NULL,
                covered INTEGER NOT NULL,
                open_24_7 INTEGER NOT NULL,
                has_fastag INTEGER NOT NULL,
                has_valet INTEGER NOT NULL,
                ev_zone_name TEXT NOT NULL,
                demand_profile TEXT NOT NULL,
                rating REAL NOT NULL,
                entrance_hint TEXT NOT NULL,
                metro_station TEXT,
                metro_lines TEXT NOT NULL DEFAULT '',
                amenities TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS bays (
                id TEXT PRIMARY KEY,
                facility_id TEXT NOT NULL,
                level TEXT NOT NULL,
                code TEXT NOT NULL,
                row_idx INTEGER NOT NULL,
                col_idx INTEGER NOT NULL,
                zone TEXT NOT NULL,
                vehicle_category TEXT NOT NULL DEFAULT 'car',
                charger_standard TEXT NOT NULL,
                charger_kw INTEGER NOT NULL,
                walk_meters INTEGER NOT NULL,
                status TEXT NOT NULL,
                covered INTEGER NOT NULL,
                FOREIGN KEY (facility_id) REFERENCES facilities(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS vehicles (
                id TEXT PRIMARY KEY,
                make_model TEXT NOT NULL,
                plate TEXT NOT NULL,
                car_type TEXT NOT NULL,
                is_ev INTEGER NOT NULL,
                charger_pref TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS reservations (
                id TEXT PRIMARY KEY,
                facility_id TEXT NOT NULL,
                facility_name TEXT NOT NULL,
                bay_id TEXT NOT NULL,
                bay_code TEXT NOT NULL,
                level TEXT NOT NULL,
                zone TEXT NOT NULL,
                vehicle_id TEXT NOT NULL,
                vehicle_name TEXT NOT NULL,
                vehicle_plate TEXT NOT NULL,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                entered_at TEXT,
                exited_at TEXT,
                duration_mins INTEGER NOT NULL,
                status TEXT NOT NULL,
                include_parking INTEGER NOT NULL DEFAULT 1,
                base_fee INTEGER NOT NULL,
                ev_fee INTEGER NOT NULL,
                discount_amount INTEGER NOT NULL,
                total_fee INTEGER NOT NULL,
                promo_code TEXT,
                payment_method TEXT NOT NULL,
                transaction_id TEXT NOT NULL,
                gate_pin TEXT NOT NULL,
                qr_payload TEXT NOT NULL,
                entrance_hint TEXT NOT NULL,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS notifications (
                id TEXT PRIMARY KEY,
                type TEXT NOT NULL,
                title TEXT NOT NULL,
                message TEXT NOT NULL,
                facility_id TEXT,
                reservation_id TEXT,
                action_label TEXT,
                is_read INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS watchlists (
                facility_id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                FOREIGN KEY (facility_id) REFERENCES facilities(id) ON DELETE CASCADE
            );
            """
        )

        existing = conn.execute("SELECT COUNT(*) FROM facilities").fetchone()[0]
        if existing == 0:
            _seed_data(conn)


def _seed_data(conn: sqlite3.Connection) -> None:
    # 20 Spacious Greater Hyderabad Hubs covering all corridors + Metro Line colors
    # Standardized Hyderabad Rates matching commuter expectations:
    # - City/Metro/Outer Hubs: ₹30/hr Car · ₹10/hr Scooty/Bike
    # - IT/Prime Hubs: ₹40/hr Car · ₹15/hr Scooty/Bike
    # - Airport Express: ₹50/hr Car · ₹15/hr Scooty/Bike
    # - EV Charging Tariff: Uniform ₹15/kWh (Cars) · ₹10/kWh (Scooty/Bike)
    raw_hubs = [
        {
            "id": "fac-airport", "name": "Shamshabad Airport", "tag": "Aero Express", "district": "Shamshabad",
            "address": "RGIA Terminal Approach Rd, Shamshabad", "lat": 17.2425, "lng": 78.4310,
            "hr": 50, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 1,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "airport", "rat": 4.8,
            "hint": "PVNR Expy / ORR Exit 14 • Terminal Gate A", "metro_st": "Pushpak Aero Express (Airport Bus)",
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,Valet,CCTV", "lvls": 2, "reg": 24, "tw": 24, "evc": 5, "ev2w": 4,
        },
        {
            "id": "fac-nagole", "name": "Nagole", "tag": "Metro East", "district": "Nagole",
            "address": "Nagole Metro Station, Inner Ring Rd", "lat": 17.3728, "lng": 78.5595,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.6,
            "hint": "Blue Line Terminus • Concourse Gate", "metro_st": "Nagole Metro Station (120m)",
            "metro_lines": "Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 24, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-uppal", "name": "Uppal", "tag": "Stadium · Metro", "district": "Uppal",
            "address": "Uppal Metro & Stadium Link Rd, Uppal", "lat": 17.4040, "lng": 78.5600,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.7,
            "hint": "Uppal Metro Pillar C-121 • Stadium East Gate", "metro_st": "Uppal & Stadium Metro (150m)",
            "metro_lines": "Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 28, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-ecil", "name": "ECIL X Roads", "tag": "North-East Hub", "district": "ECIL / Kapra",
            "address": "ECIL Bus Terminus Rd, North-East Hyderabad", "lat": 17.4720, "lng": 78.5620,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "transit", "rat": 4.5,
            "hint": "Radhika Multiplex Corner • Covered West Entry", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 24, "evc": 3, "ev2w": 4,
        },
        {
            "id": "fac-kompally", "name": "Kompally", "tag": "NH-44 North", "district": "Kompally",
            "address": "NH-44 Medchal Highway, Kompally", "lat": 17.4950, "lng": 78.4880,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "transit", "rat": 4.6,
            "hint": "Cine Planet Service Rd • FASTag North Gate", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 24, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-miyapur", "name": "Miyapur", "tag": "Red Terminus", "district": "Miyapur",
            "address": "NH-65 Miyapur Metro Park & Ride", "lat": 17.4950, "lng": 78.3580,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.7,
            "hint": "Red Line NW Terminus • Feeder Gate 1", "metro_st": "Miyapur Terminus (80m)",
            "metro_lines": "Red", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 2, "reg": 20, "tw": 28, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-kphb", "name": "KPHB Colony", "tag": "Metro P+R", "district": "Kukatpally",
            "address": "NH-65, KPHB Metro Station", "lat": 17.4820, "lng": 78.3930,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.5,
            "hint": "KPHB Metro Pillar Entry • 24/7 FASTag", "metro_st": "KPHB Colony Metro (90m)",
            "metro_lines": "Red", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 28, "evc": 3, "ev2w": 4,
        },
        {
            "id": "fac-hitec", "name": "HITEC City", "tag": "Cyber Towers", "district": "Madhapur",
            "address": "Cyber Towers Junction, HITEC Main Rd", "lat": 17.4540, "lng": 78.3812,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "it_hub", "rat": 4.8,
            "hint": "West Ramp HITEC Main Rd • FASTag Gate 2", "metro_st": "HITEC City Metro (180m)",
            "metro_lines": "Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 2, "reg": 20, "tw": 28, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-mindspace", "name": "Mindspace", "tag": "Inorbit Zone", "district": "Raidurg",
            "address": "Inorbit Mall Rd, Mindspace", "lat": 17.4300, "lng": 78.3860,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "it_hub", "rat": 4.7,
            "hint": "Cable Bridge Approach • Level L1", "metro_st": "Raidurg Metro Terminus (250m)",
            "metro_lines": "Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 2, "reg": 24, "tw": 28, "evc": 5, "ev2w": 4,
        },
        {
            "id": "fac-gachibowli", "name": "Gachibowli", "tag": "Financial Dist", "district": "Nanakramguda",
            "address": "ISB Rd, Financial District", "lat": 17.4185, "lng": 78.3440,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "financial", "rat": 4.7,
            "hint": "Basement Spiral off ISB Rd • B1 Entry", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 2, "reg": 18, "tw": 24, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-narsingi", "name": "Narsingi", "tag": "ORR West", "district": "Narsingi / Kokapet",
            "address": "Nehru ORR Exit 18A Service Rd, Narsingi", "lat": 17.3920, "lng": 78.3340,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "it_hub", "rat": 4.6,
            "hint": "Kokapet Neopolis Link • ORR West Gate", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 24, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-kondapur", "name": "Kondapur", "tag": "Tech Plaza", "district": "Kondapur",
            "address": "Botanical Garden Rd, Kondapur", "lat": 17.4605, "lng": 78.3620,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "it_hub", "rat": 4.6,
            "hint": "Sarath City Side Gate • Level L1", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 24, "evc": 3, "ev2w": 3,
        },
        {
            "id": "fac-jubilee", "name": "Jubilee Hills", "tag": "Checkpost", "district": "Jubilee Hills",
            "address": "Road No. 36, Jubilee Checkpost", "lat": 17.4290, "lng": 78.4090,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.6,
            "hint": "Metro Pillar C-1632 Entry • Skybridge", "metro_st": "Jubilee Hills Checkpost (110m)",
            "metro_lines": "Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 20, "evc": 4, "ev2w": 3,
        },
        {
            "id": "fac-banjara", "name": "Banjara Hills", "tag": "Road No. 12", "district": "Banjara Hills",
            "address": "Road No. 12, Banjara Hills", "lat": 17.4130, "lng": 78.4445,
            "hr": 40, "twr": 15, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 1,
            "evz": "Valet EV Wing Z1 (Cars & 2W)", "prof": "upscale", "rat": 4.9,
            "hint": "Valet Porch on Road No. 12 • Covered Bay", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,Valet,CCTV", "lvls": 2, "reg": 16, "tw": 16, "evc": 4, "ev2w": 3,
        },
        {
            "id": "fac-mehdipatnam", "name": "Mehdipatnam", "tag": "PVNR Ramp", "district": "Mehdipatnam",
            "address": "PVNR Expy Pillar 12, Mehdipatnam", "lat": 17.3920, "lng": 78.4360,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "transit", "rat": 4.5,
            "hint": "Sarojini Eye Hospital Side • PVNR North Gate", "metro_st": None,
            "metro_lines": "", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 24, "evc": 3, "ev2w": 4,
        },
        {
            "id": "fac-necklace", "name": "Necklace Road", "tag": "Solar EV", "district": "Hussain Sagar",
            "address": "PVNR Marg, Tank Bund", "lat": 17.4235, "lng": 78.4720,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "Solar EV Plaza (Cars & 2W)", "prof": "lakefront", "rat": 4.9,
            "hint": "Lakefront Solar Canopy • EV Fast Lane", "metro_st": "Khairatabad Metro (550m walk)",
            "metro_lines": "Red", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 24, "evc": 6, "ev2w": 4,
        },
        {
            "id": "fac-panjagutta", "name": "Panjagutta", "tag": "Central Metro", "district": "Panjagutta",
            "address": "NIMS Flyover Link Rd, Panjagutta", "lat": 17.4318, "lng": 78.4510,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.5,
            "hint": "Panjagutta Metro Skywalk Gate", "metro_st": "Panjagutta Metro (70m • 1 stop to Ameerpet)",
            "metro_lines": "Red,Blue", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 20, "tw": 24, "evc": 3, "ev2w": 3,
        },
        {
            "id": "fac-secbad", "name": "Secunderabad", "tag": "Rail Terminus", "district": "Secunderabad",
            "address": "Station Rd, Secunderabad", "lat": 17.4348, "lng": 78.5015,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "Rail EV Zone Z1 (Cars & 2W)", "prof": "transit", "rat": 4.6,
            "hint": "Platform 1 Approach • 24/7 FASTag", "metro_st": "Secunderabad East & West Metro (100m)",
            "metro_lines": "Blue,Green", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 2, "reg": 24, "tw": 32, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-mgbs", "name": "MGBS / Abids", "tag": "Interchange", "district": "Imlibun / Old City",
            "address": "MGBS Terminal Approach, Musi Bank", "lat": 17.3815, "lng": 78.4860,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.6,
            "hint": "Imlibun Gate 3 • Direct Footbridge to Metro", "metro_st": "MGBS Interchange Station (100m)",
            "metro_lines": "Red,Green", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 28, "evc": 4, "ev2w": 4,
        },
        {
            "id": "fac-charminar", "name": "Charminar", "tag": "Old City", "district": "Charminar",
            "address": "Laad Bazaar Link Rd, Charminar", "lat": 17.3620, "lng": 78.4745,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 0, "o24": 0, "fst": 0, "val": 0,
            "evz": "E-Zone (Cars & 2W)", "prof": "heritage", "rat": 4.5,
            "hint": "Laad Bazaar Gate • Free E-Shuttle", "metro_st": "MGBS Green/Red Feeder (2.2 km)",
            "metro_lines": "Green", "amenities": "CCTV,E-Shuttle", "lvls": 1, "reg": 24, "tw": 24, "evc": 3, "ev2w": 3,
        },
        {
            "id": "fac-lbnagar", "name": "LB Nagar", "tag": "Ring Road", "district": "LB Nagar",
            "address": "Inner Ring Rd, LB Nagar", "lat": 17.3420, "lng": 78.5480,
            "hr": 30, "twr": 10, "evr": 15, "ev2wr": 10, "cov": 1, "o24": 1, "fst": 1, "val": 0,
            "evz": "EV Wing Z1 (Cars & 2W)", "prof": "metro", "rat": 4.6,
            "hint": "LB Nagar Terminus Ramp • Metro P+R", "metro_st": "LB Nagar Metro Terminus (100m)",
            "metro_lines": "Red", "amenities": "Covered,24/7,FASTag,CCTV", "lvls": 1, "reg": 24, "tw": 28, "evc": 4, "ev2w": 4,
        },
    ]

    facilities = []
    for h in raw_hubs:
        x, y = lat_lng_to_canvas(h["lat"], h["lng"])
        facilities.append((
            h["id"], h["name"], h["tag"], h["district"], h["address"], x, y, h["lat"], h["lng"],
            h["hr"], h["twr"], h["evr"], h["ev2wr"], h["cov"], h["o24"], h["fst"], h["val"],
            h["evz"], h["prof"], h["rat"], h["hint"], h["metro_st"], h["metro_lines"], h["amenities"],
        ))

    conn.executemany(
        """
        INSERT INTO facilities (
            id, name, tag, district, address, x, y, lat, lng, hourly_rate, two_wheeler_rate, ev_kwh_rate, ev_2w_kwh_rate,
            covered, open_24_7, has_fastag, has_valet, ev_zone_name, demand_profile, rating, entrance_hint,
            metro_station, metro_lines, amenities
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        facilities,
    )

    rng = random.Random(2026)
    bays_rows = []
    for h in raw_hubs:
        fac_id = h["id"]
        covered = h["cov"]
        num_levels, slots_per_level, tw_slots, ev_car_count, ev_2w_count = h["lvls"], h["reg"], h["tw"], h["evc"], h["ev2w"]
        ratio = 0.25 if fac_id == "fac-charminar" else 0.58

        # 1) Car Floors (L1, L2, B1, B2, or G1)
        is_underground = fac_id == "fac-gachibowli"
        for lvl_idx in range(num_levels):
            lvl_label = f"B{lvl_idx + 1}" if is_underground else (f"L{lvl_idx + 1}" if covered else "G1")
            base_num = (lvl_idx + 1) * 100
            for idx in range(slots_per_level):
                code = f"{base_num + idx + 1}"
                bay_id = f"{fac_id}-{lvl_label}-{code}"
                if fac_id == "fac-charminar":
                    status = "available" if idx in (1, 4, 9, 14, 19, 22) else "occupied"
                else:
                    status = "available" if rng.random() < ratio else "occupied"
                bays_rows.append(
                    (bay_id, fac_id, lvl_label, code, idx // 6, idx % 6, "regular_deck", "car", "None", 0, 12 + (idx % 8) * 6, status, covered)
                )

        # 2) Dedicated Scooty & Bike Floor (2W-G or 2W-B1)
        tw_level = "2W-B1" if is_underground else "2W-G"
        for idx in range(tw_slots):
            code = f"2W-{idx + 1:02d}"
            bay_id = f"{fac_id}-{tw_level}-{code}"
            status = "available" if (idx == 0 or rng.random() < max(0.45, ratio)) else "occupied"
            bays_rows.append(
                (bay_id, fac_id, tw_level, code, idx // 6, idx % 6, "two_wheeler_floor", "two_wheeler", "None", 0, 8 + (idx % 6) * 4, status, covered)
            )

        # 3) Dedicated EV Charging Wing (EV-Z1): split into EV Cars & EV Bikes/Scooty
        for idx in range(ev_car_count):
            code = f"EVC-{idx + 1:02d}"
            bay_id = f"{fac_id}-EV-{code}"
            if fac_id in ("fac-airport", "fac-mindspace", "fac-necklace") and idx < 2:
                std, kw = "CCS2 60kW DC", 60
            elif idx % 2 == 0:
                std, kw = "CCS2 30kW DC", 30
            else:
                std, kw = "Type-2 22kW AC", 22
            status = "available" if (idx == 0 or rng.random() < 0.65) else "occupied"
            bays_rows.append(
                (bay_id, fac_id, "EV-Z1", code, 0, idx, "ev_charging_zone", "car", std, kw, 10 + idx * 4, status, 1)
            )

        for idx in range(ev_2w_count):
            code = f"EV2W-{idx + 1:02d}"
            bay_id = f"{fac_id}-EV-{code}"
            std, kw = ("Ather / 15A AC", 3) if idx % 2 == 0 else ("3.3kW Smart AC", 3)
            status = "available" if (idx == 0 or rng.random() < 0.68) else "occupied"
            bays_rows.append(
                (bay_id, fac_id, "EV-Z1", code, 1, idx, "ev_charging_zone", "two_wheeler", std, kw, 8 + idx * 3, status, 1)
            )

    conn.executemany(
        """
        INSERT INTO bays (
            id, facility_id, level, code, row_idx, col_idx, zone, vehicle_category,
            charger_standard, charger_kw, walk_meters, status, covered
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        bays_rows,
    )

    now = datetime.now(timezone.utc)
    past_res = [
        (
            "HYD-70412",
            "fac-necklace",
            "Necklace Road",
            "fac-necklace-EV-EVC-01",
            "EVC-01",
            "EV-Z1",
            "ev_charging_zone",
            "user-car",
            "Tata Nexon.ev",
            "TS 09 EV 4021",
            (now - timedelta(days=1, hours=5)).isoformat(),
            (now - timedelta(days=1, hours=3)).isoformat(),
            (now - timedelta(days=1, hours=4, minutes=50)).isoformat(),
            (now - timedelta(days=1, hours=3)).isoformat(),
            120,
            "completed",
            60,
            90,
            30,
            120,
            "HYDEV20",
            "UPI Instant (GPay)",
            "UPI-992014",
            "482910",
            "DECCANPULSE|HYD-70412|fac-necklace|EV-01",
            "Lakefront Solar Canopy • EV Fast Lane",
            (now - timedelta(days=1, hours=6)).isoformat(),
        ),
        (
            "HYD-70289",
            "fac-secbad",
            "Secunderabad",
            "fac-secbad-2W-G-2W-02",
            "2W-02",
            "2W-G",
            "two_wheeler_floor",
            "user-2w",
            "Honda Activa 6G",
            "TS 09 SC 2219",
            (now - timedelta(days=2, hours=4)).isoformat(),
            (now - timedelta(days=2, hours=1)).isoformat(),
            (now - timedelta(days=2, hours=3, minutes=54)).isoformat(),
            (now - timedelta(days=2, hours=1)).isoformat(),
            180,
            "completed",
            30,
            0,
            5,
            25,
            "FASTAG15",
            "UPI Instant (PhonePe)",
            "UPI-981140",
            "731045",
            "DECCANPULSE|HYD-70289|fac-secbad|2W-02",
            "Platform 1 Approach • 24/7 FASTag",
            (now - timedelta(days=2, hours=5)).isoformat(),
        ),
    ]
    conn.executemany(
        """
        INSERT INTO reservations (
            id, facility_id, facility_name, bay_id, bay_code, level, zone, vehicle_id,
            vehicle_name, vehicle_plate, start_time, end_time, entered_at, exited_at, duration_mins, status,
            base_fee, ev_fee, discount_amount, total_fee, promo_code, payment_method,
            transaction_id, gate_pin, qr_payload, entrance_hint, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        past_res,
    )

    notifications = [
        (
            "notif-welcome",
            "system",
            "Live City Sensors Connected",
            "Car floors, Scooty & Bike floors, and EV Charging wings are updating live across Hyderabad.",
            "fac-hitec",
            None,
            "Open Map",
            0,
            utc_now_iso(),
        ),
        (
            "notif-promo",
            "payment",
            "HYDEV20: 20% Off EV Charging",
            "Dedicated EV charger zones are ready with CCS2, Type-2, and Bharat DC bays.",
            "fac-necklace",
            None,
            "View EV Wing",
            0,
            (now - timedelta(minutes=10)).isoformat(),
        ),
    ]
    conn.executemany(
        """
        INSERT INTO notifications (
            id, type, title, message, facility_id, reservation_id, action_label, is_read, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        notifications,
    )


def search_destinations(
    query: str = "",
    user_x: Optional[float] = None,
    user_y: Optional[float] = None,
) -> List[DestinationPoint]:
    q = query.strip().lower()
    user_lat_lng = canvas_to_lat_lng(user_x, user_y) if (user_x is not None and user_y is not None) else None
    results: List[DestinationPoint] = []
    for did, name, area, cat, lat, lng in HYDERABAD_DESTINATIONS:
        if q and q not in name.lower() and q not in area.lower() and q not in cat.lower():
            continue
        x, y = lat_lng_to_canvas(lat, lng)
        dist = None
        mins = None
        brg = None
        if user_lat_lng:
            dist = compute_haversine_road_km(user_lat_lng[0], user_lat_lng[1], lat, lng)
            mins = max(3, round(dist * 2.5 + 2))
            brg = compute_bearing_label(user_lat_lng[0], user_lat_lng[1], lat, lng)
        results.append(
            DestinationPoint(
                id=did,
                name=name,
                area=area,
                category=cat,
                lat=lat,
                lng=lng,
                x=x,
                y=y,
                distance_km=dist,
                drive_mins=mins,
                bearing=brg,
            )
        )
    if user_lat_lng:
        results.sort(key=lambda d: d.distance_km or 999.0)
    return results


def list_facilities(
    user_x: float = 585.0,
    user_y: float = 217.0,
    search: str = "",
    only_ev: bool = False,
    only_two_wheeler: bool = False,
    only_covered: bool = False,
    max_distance_km: Optional[float] = None,
    dest_x: Optional[float] = None,
    dest_y: Optional[float] = None,
) -> List[FacilitySummary]:
    user_lat, user_lng = canvas_to_lat_lng(user_x, user_y)
    dest_lat_lng = canvas_to_lat_lng(dest_x, dest_y) if (dest_x is not None and dest_y is not None) else None

    with get_conn() as conn:
        fac_rows = conn.execute("SELECT * FROM facilities").fetchall()
        bay_rows = conn.execute("SELECT facility_id, zone, vehicle_category, charger_standard, status FROM bays").fetchall()
        watch_rows = conn.execute("SELECT facility_id FROM watchlists").fetchall()

    watch_set = {r["facility_id"] for r in watch_rows}
    bays_by_fac: Dict[str, List[sqlite3.Row]] = {}
    for b in bay_rows:
        bays_by_fac.setdefault(b["facility_id"], []).append(b)

    q = search.strip().lower()
    results: List[FacilitySummary] = []

    for f in fac_rows:
        if q and q not in f["name"].lower() and q not in f["tag"].lower() and q not in f["district"].lower() and q not in f["address"].lower():
            continue
        if only_covered and not f["covered"]:
            continue

        dist_km = compute_haversine_road_km(user_lat, user_lng, f["lat"], f["lng"])
        if max_distance_km is not None and dist_km > max_distance_km:
            continue

        fac_bays = bays_by_fac.get(f["id"], [])
        total = len(fac_bays)
        available = sum(1 for b in fac_bays if b["status"] == "available")

        reg_bays = [b for b in fac_bays if b["zone"] == "regular_deck"]
        tw_bays = [b for b in fac_bays if b["zone"] == "two_wheeler_floor"]
        ev_bays = [b for b in fac_bays if b["zone"] == "ev_charging_zone"]

        reg_total = len(reg_bays)
        reg_avail = sum(1 for b in reg_bays if b["status"] == "available")
        tw_total = len(tw_bays)
        tw_avail = sum(1 for b in tw_bays if b["status"] == "available")
        ev_car_bays = [b for b in ev_bays if b["vehicle_category"] == "car"]
        ev_2w_bays = [b for b in ev_bays if b["vehicle_category"] == "two_wheeler"]
        ev_car_total = len(ev_car_bays)
        ev_car_avail = sum(1 for b in ev_car_bays if b["status"] == "available")
        ev_2w_total = len(ev_2w_bays)
        ev_2w_avail = sum(1 for b in ev_2w_bays if b["status"] == "available")
        ev_total = len(ev_bays)
        ev_avail = sum(1 for b in ev_bays if b["status"] == "available")
        charger_types = sorted({b["charger_standard"] for b in ev_bays if b["charger_standard"] != "None"})

        if only_ev and ev_avail == 0:
            continue
        if only_two_wheeler and tw_avail == 0:
            continue

        drive_mins = max(3, round(dist_km * 2.5 + 2))
        bearing = compute_bearing_label(user_lat, user_lng, f["lat"], f["lng"])
        walk_to_dest = None
        if dest_lat_lng:
            walk_km = compute_haversine_road_km(dest_lat_lng[0], dest_lat_lng[1], f["lat"], f["lng"])
            walk_to_dest = max(1, round(walk_km * 11))

        rate_5 = round_inr_5(f["hourly_rate"], minimum=20)
        tw_rate_5 = round_inr_5(f["two_wheeler_rate"], minimum=10)
        results.append(
            FacilitySummary(
                id=f["id"],
                name=f["name"],
                tag=f["tag"],
                district=f["district"],
                address=f["address"],
                x=f["x"],
                y=f["y"],
                lat=f["lat"],
                lng=f["lng"],
                hourly_rate=rate_5,
                two_wheeler_rate=tw_rate_5,
                ev_kwh_rate=round_inr_5(f["ev_kwh_rate"], minimum=10),
                ev_2w_kwh_rate=round_inr_5(f["ev_2w_kwh_rate"], minimum=10),
                covered=bool(f["covered"]),
                open_24_7=bool(f["open_24_7"]),
                has_fastag=bool(f["has_fastag"]),
                has_valet=bool(f["has_valet"]),
                ev_zone_name=f["ev_zone_name"],
                ev_charger_types=charger_types,
                demand_profile=f["demand_profile"],
                rating=f["rating"],
                entrance_hint=f["entrance_hint"],
                metro_station=f["metro_station"],
                metro_lines=[s.strip() for s in f["metro_lines"].split(",") if s.strip()],
                amenities=[s.strip() for s in f["amenities"].split(",") if s.strip()],
                total_bays=total,
                available_bays=available,
                regular_total=reg_total,
                regular_available=reg_avail,
                two_wheeler_total=tw_total,
                two_wheeler_available=tw_avail,
                ev_total=ev_total,
                ev_available=ev_avail,
                ev_car_total=ev_car_total,
                ev_car_available=ev_car_avail,
                ev_2w_total=ev_2w_total,
                ev_2w_available=ev_2w_avail,
                effective_rate=rate_5,
                distance_km=dist_km,
                drive_mins=drive_mins,
                bearing=bearing,
                walk_to_dest_mins=walk_to_dest,
                watchlisted=f["id"] in watch_set,
            )
        )

    if dest_lat_lng:
        results.sort(key=lambda item: (item.walk_to_dest_mins or 999, item.distance_km))
    else:
        results.sort(key=lambda item: item.distance_km)
    return results


def get_facility_detail(
    facility_id: str,
    user_x: float = 585.0,
    user_y: float = 217.0,
) -> Optional[FacilityDetail]:
    facilities = list_facilities(user_x=user_x, user_y=user_y)
    summary = next((f for f in facilities if f.id == facility_id), None)
    if not summary:
        return None

    with get_conn() as conn:
        bay_rows = conn.execute(
            "SELECT * FROM bays WHERE facility_id = ? ORDER BY zone, level, row_idx, col_idx",
            (facility_id,),
        ).fetchall()

    bays: List[BayOut] = []
    levels_seen: List[str] = []
    for b in bay_rows:
        if b["level"] not in levels_seen:
            levels_seen.append(b["level"])
        bays.append(
            BayOut(
                id=b["id"],
                facility_id=b["facility_id"],
                level=b["level"],
                code=b["code"],
                row_idx=b["row_idx"],
                col_idx=b["col_idx"],
                zone=b["zone"],
                vehicle_category=b["vehicle_category"],
                charger_standard=b["charger_standard"],
                charger_kw=b["charger_kw"],
                walk_meters=b["walk_meters"],
                status=b["status"],
                covered=bool(b["covered"]),
            )
        )
    return FacilityDetail(**summary.model_dump(), bays=bays, levels=levels_seen)


def list_vehicles() -> List[VehicleOut]:
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM vehicles").fetchall()
    return [
        VehicleOut(
            id=r["id"],
            make_model=r["make_model"],
            plate=r["plate"],
            car_type=r["car_type"],
            is_ev=bool(r["is_ev"]),
            charger_pref=r["charger_pref"],
        )
        for r in rows
    ]


def create_vehicle(payload: VehicleCreate) -> VehicleOut:
    vid = f"veh-{uuid.uuid4().hex[:6]}"
    is_ev = payload.is_ev or "EV" in payload.car_type
    charger = payload.charger_pref if is_ev and payload.charger_pref != "None" else ("CCS2 DC" if is_ev else "None")
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO vehicles (id, make_model, plate, car_type, is_ev, charger_pref) VALUES (?, ?, ?, ?, ?, ?)",
            (vid, payload.make_model.strip(), payload.plate.strip().upper(), payload.car_type, int(is_ev), charger),
        )
    return VehicleOut(
        id=vid,
        make_model=payload.make_model.strip(),
        plate=payload.plate.strip().upper(),
        car_type=payload.car_type,
        is_ev=is_ev,
        charger_pref=charger,
    )


def delete_vehicle(vehicle_id: str) -> None:
    with get_conn() as conn:
        conn.execute("DELETE FROM vehicles WHERE id = ?", (vehicle_id,))


def create_reservation(payload: ReservationCreate) -> ReservationOut:
    with get_conn() as conn:
        fac = conn.execute("SELECT * FROM facilities WHERE id = ?", (payload.facility_id,)).fetchone()
        if not fac:
            raise ValueError("Parking area not found")

        bay = conn.execute(
            "SELECT * FROM bays WHERE id = ? AND facility_id = ?",
            (payload.bay_id, payload.facility_id),
        ).fetchone()
        if not bay:
            raise ValueError("Slot not found in this area")
        if bay["status"] != "available":
            raise ValueError(f"Slot {bay['level']}-{bay['code']} was just taken")

        veh = conn.execute("SELECT * FROM vehicles WHERE id = ?", (payload.vehicle_id,)).fetchone()
        if not veh:
            raise ValueError("Please add a vehicle in Vehicles first")

        quote = calculate_price_quote(
            hourly_rate=fac["hourly_rate"],
            two_wheeler_rate=fac["two_wheeler_rate"],
            ev_kwh_rate=fac["ev_kwh_rate"],
            ev_2w_kwh_rate=fac["ev_2w_kwh_rate"],
            duration_mins=payload.duration_mins,
            zone=bay["zone"],
            vehicle_category=bay["vehicle_category"],
            charge_ev=payload.charge_ev and bay["zone"] == "ev_charging_zone",
            promo_code=payload.promo_code,
            include_parking=payload.include_parking,
        )

        now = datetime.now(timezone.utc)
        start_dt = now
        end_dt = now + timedelta(minutes=payload.duration_mins)
        res_id = f"HYD-{random.randint(10000, 99999)}"
        txn_id = f"UPI-{uuid.uuid4().hex[:8].upper()}"
        gate_pin = f"{random.randint(100000, 999999)}"
        qr_payload = f"DECCANPULSE|{res_id}|{fac['id']}|{bay['level']}-{bay['code']}|{veh['plate']}"

        conn.execute("UPDATE bays SET status = 'reserved' WHERE id = ?", (bay["id"],))
        conn.execute(
            """
            INSERT INTO reservations (
                id, facility_id, facility_name, bay_id, bay_code, level, zone, vehicle_id,
                vehicle_name, vehicle_plate, start_time, end_time, entered_at, exited_at, duration_mins, status,
                include_parking, base_fee, ev_fee, discount_amount, total_fee, promo_code, payment_method,
                transaction_id, gate_pin, qr_payload, entrance_hint, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, 'reserved', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                res_id,
                fac["id"],
                fac["name"],
                bay["id"],
                bay["code"],
                bay["level"],
                bay["zone"],
                veh["id"],
                veh["make_model"],
                veh["plate"],
                start_dt.isoformat(),
                end_dt.isoformat(),
                payload.duration_mins,
                int(payload.include_parking),
                quote.base_fee,
                quote.ev_fee,
                quote.discount_amount,
                quote.total_fee,
                quote.promo_applied,
                payload.payment_method,
                txn_id,
                gate_pin,
                qr_payload,
                fac["entrance_hint"],
                now.isoformat(),
            ),
        )

        _add_notification_conn(
            conn,
            notif_type="booking",
            title=f"Slot {bay['level']}-{bay['code']} Reserved",
            message=f"{fac['name']} held for {veh['plate']}. Gate PIN: {gate_pin} • Paid ₹{quote.total_fee}.",
            facility_id=fac["id"],
            reservation_id=res_id,
            action_label="Open Ticket",
        )
        row = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
    return _row_to_reservation(row)


def enter_reservation(res_id: str) -> ReservationOut:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
        if not row:
            raise ValueError("Booking not found")
        if row["status"] != "reserved":
            raise ValueError("Only reserved bookings can enter the gate")

        now_iso = utc_now_iso()
        conn.execute("UPDATE reservations SET status = 'entered', entered_at = ? WHERE id = ?", (now_iso, res_id))
        conn.execute("UPDATE bays SET status = 'occupied' WHERE id = ?", (row["bay_id"],))
        _add_notification_conn(
            conn,
            notif_type="booking",
            title=f"Gate Opened — Parked at {row['level']}-{row['bay_code']}",
            message=f"{row['vehicle_plate']} entered {row['facility_name']}.",
            facility_id=row["facility_id"],
            reservation_id=res_id,
            action_label="View Ticket",
        )
        updated = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
    return _row_to_reservation(updated)


def exit_reservation(res_id: str) -> ReservationOut:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
        if not row:
            raise ValueError("Booking not found")
        if row["status"] not in ("reserved", "entered"):
            raise ValueError("Session already closed")

        now_iso = utc_now_iso()
        conn.execute("UPDATE reservations SET status = 'completed', exited_at = ? WHERE id = ?", (now_iso, res_id))
        conn.execute("UPDATE bays SET status = 'available' WHERE id = ?", (row["bay_id"],))
        _add_notification_conn(
            conn,
            notif_type="payment",
            title=f"Exited {row['facility_name']}",
            message=f"Slot {row['level']}-{row['bay_code']} is now free. Bill ₹{row['total_fee']} settled.",
            facility_id=row["facility_id"],
            reservation_id=res_id,
            action_label="View History",
        )
        updated = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
    return _row_to_reservation(updated)


def cancel_reservation(res_id: str) -> ReservationOut:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
        if not row:
            raise ValueError("Booking not found")
        if row["status"] != "reserved":
            raise ValueError("Only reserved bookings before entry can be cancelled")

        conn.execute("UPDATE reservations SET status = 'cancelled' WHERE id = ?", (res_id,))
        conn.execute("UPDATE bays SET status = 'available' WHERE id = ?", (row["bay_id"],))
        _add_notification_conn(
            conn,
            notif_type="booking",
            title="Booking Cancelled",
            message=f"Slot {row['level']}-{row['bay_code']} released at {row['facility_name']}.",
            facility_id=row["facility_id"],
            reservation_id=res_id,
            action_label="View History",
        )
        updated = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
    return _row_to_reservation(updated)


def extend_reservation(res_id: str, extra_mins: int, payment_method: str) -> ReservationOut:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
        if not row:
            raise ValueError("Booking not found")
        if row["status"] not in ("reserved", "entered"):
            raise ValueError("Cannot extend a closed booking")

        fac = conn.execute("SELECT hourly_rate, two_wheeler_rate FROM facilities WHERE id = ?", (row["facility_id"],)).fetchone()
        bay_row = conn.execute("SELECT vehicle_category FROM bays WHERE id = ?", (row["bay_id"],)).fetchone()
        is_tw = row["zone"] == "two_wheeler_floor" or (bay_row and bay_row["vehicle_category"] == "two_wheeler")
        rate = fac["two_wheeler_rate"] if (fac and is_tw) else (fac["hourly_rate"] if fac else 30)
        extra_fee = round_inr_5(rate * (extra_mins / 60.0), minimum=10)

        end_dt = datetime.fromisoformat(row["end_time"]) + timedelta(minutes=extra_mins)
        new_dur = row["duration_mins"] + extra_mins
        new_base = row["base_fee"] + extra_fee
        new_total = row["total_fee"] + extra_fee

        conn.execute(
            "UPDATE reservations SET end_time = ?, duration_mins = ?, base_fee = ?, total_fee = ? WHERE id = ?",
            (end_dt.isoformat(), new_dur, new_base, new_total, res_id),
        )
        _add_notification_conn(
            conn,
            notif_type="payment",
            title=f"Extended +{extra_mins} mins",
            message=f"{row['facility_name']} extended via {payment_method} (+₹{extra_fee}).",
            facility_id=row["facility_id"],
            reservation_id=res_id,
            action_label="Open Ticket",
        )
        updated = conn.execute("SELECT * FROM reservations WHERE id = ?", (res_id,)).fetchone()
    return _row_to_reservation(updated)


def list_reservations(status_filter: Optional[str] = None) -> List[ReservationOut]:
    with get_conn() as conn:
        if status_filter and status_filter != "all":
            rows = conn.execute(
                "SELECT * FROM reservations WHERE status = ? ORDER BY created_at DESC",
                (status_filter,),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM reservations ORDER BY created_at DESC").fetchall()
    return [_row_to_reservation(r) for r in rows]


def toggle_watchlist(facility_id: str) -> bool:
    with get_conn() as conn:
        fac = conn.execute("SELECT name FROM facilities WHERE id = ?", (facility_id,)).fetchone()
        if not fac:
            raise ValueError("Parking area not found")
        exists = conn.execute("SELECT 1 FROM watchlists WHERE facility_id = ?", (facility_id,)).fetchone()
        if exists:
            conn.execute("DELETE FROM watchlists WHERE facility_id = ?", (facility_id,))
            return False
        conn.execute("INSERT INTO watchlists (facility_id, created_at) VALUES (?, ?)", (facility_id, utc_now_iso()))
        _add_notification_conn(
            conn,
            notif_type="availability",
            title="Spot Alert Enabled",
            message=f"We will alert you as soon as a slot opens at {fac['name']}.",
            facility_id=facility_id,
            action_label="View Area",
        )
        return True


def list_notifications() -> List[NotificationOut]:
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM notifications ORDER BY created_at DESC LIMIT 35").fetchall()
    return [
        NotificationOut(
            id=r["id"],
            type=r["type"],
            title=r["title"],
            message=r["message"],
            facility_id=r["facility_id"],
            reservation_id=r["reservation_id"],
            action_label=r["action_label"],
            is_read=bool(r["is_read"]),
            created_at=r["created_at"],
        )
        for r in rows
    ]


def mark_notifications_read(notif_id: Optional[str] = None) -> None:
    with get_conn() as conn:
        if notif_id:
            conn.execute("UPDATE notifications SET is_read = 1 WHERE id = ?", (notif_id,))
        else:
            conn.execute("UPDATE notifications SET is_read = 1")


def create_notification(
    notif_type: str,
    title: str,
    message: str,
    facility_id: Optional[str] = None,
    reservation_id: Optional[str] = None,
    action_label: Optional[str] = None,
) -> NotificationOut:
    with get_conn() as conn:
        nid = _add_notification_conn(conn, notif_type, title, message, facility_id, reservation_id, action_label)
        row = conn.execute("SELECT * FROM notifications WHERE id = ?", (nid,)).fetchone()
    return NotificationOut(
        id=row["id"],
        type=row["type"],
        title=row["title"],
        message=row["message"],
        facility_id=row["facility_id"],
        reservation_id=row["reservation_id"],
        action_label=row["action_label"],
        is_read=bool(row["is_read"]),
        created_at=row["created_at"],
    )


def simulate_sensor_pulse() -> Optional[Dict[str, Any]]:
    with get_conn() as conn:
        active_bay_ids = {
            r["bay_id"]
            for r in conn.execute("SELECT bay_id FROM reservations WHERE status IN ('reserved', 'entered')").fetchall()
        }
        bays = [b for b in conn.execute("SELECT * FROM bays").fetchall() if b["id"] not in active_bay_ids]
        if not bays:
            return None

        watch_set = {r["facility_id"] for r in conn.execute("SELECT facility_id FROM watchlists").fetchall()}
        target = random.choice(bays)
        old_status = target["status"]

        if old_status == "available":
            new_status = "reserved"
            action_text = "booked"
        elif old_status == "reserved":
            new_status = "occupied"
            action_text = "entered gate"
        else:
            new_status = "available"
            action_text = "exited (now open)"

        conn.execute("UPDATE bays SET status = ? WHERE id = ?", (new_status, target["id"]))
        fac = conn.execute("SELECT name FROM facilities WHERE id = ?", (target["facility_id"],)).fetchone()
        fac_name = fac["name"] if fac else target["facility_id"]

        alert_created = None
        if new_status == "available" and target["facility_id"] in watch_set:
            if target["zone"] == "ev_charging_zone":
                zone_label = "EV 2W Charger" if target["vehicle_category"] == "two_wheeler" else "EV Car Charger"
            elif target["zone"] == "two_wheeler_floor":
                zone_label = "Scooty/Bike"
            else:
                zone_label = "Car Slot"
            alert_created = _add_notification_conn(
                conn,
                notif_type="availability",
                title=f"Slot Freed at {fac_name}!",
                message=f"{zone_label} {target['level']}-{target['code']} just vacated.",
                facility_id=target["facility_id"],
                action_label="Book Now",
            )

    return {
        "event": "bay_update",
        "facility_id": target["facility_id"],
        "facility_name": fac_name,
        "bay_id": target["id"],
        "code": target["code"],
        "level": target["level"],
        "zone": target["zone"],
        "status": new_status,
        "summary": f"Vehicle {action_text} • {fac_name} ({target['level']}-{target['code']})",
        "alert_id": alert_created,
    }


def _add_notification_conn(
    conn: sqlite3.Connection,
    notif_type: str,
    title: str,
    message: str,
    facility_id: Optional[str] = None,
    reservation_id: Optional[str] = None,
    action_label: Optional[str] = None,
) -> str:
    nid = f"notif-{uuid.uuid4().hex[:8]}"
    conn.execute(
        """
        INSERT INTO notifications (
            id, type, title, message, facility_id, reservation_id, action_label, is_read, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?)
        """,
        (nid, notif_type, title, message, facility_id, reservation_id, action_label, utc_now_iso()),
    )
    return nid


def _row_to_reservation(row: sqlite3.Row) -> ReservationOut:
    return ReservationOut(
        id=row["id"],
        facility_id=row["facility_id"],
        facility_name=row["facility_name"],
        bay_id=row["bay_id"],
        bay_code=row["bay_code"],
        level=row["level"],
        zone=row["zone"],
        vehicle_id=row["vehicle_id"],
        vehicle_name=row["vehicle_name"],
        vehicle_plate=row["vehicle_plate"],
        start_time=row["start_time"],
        end_time=row["end_time"],
        entered_at=row["entered_at"],
        exited_at=row["exited_at"],
        duration_mins=row["duration_mins"],
        status=row["status"],
        include_parking=bool(row["include_parking"]),
        base_fee=row["base_fee"],
        ev_fee=row["ev_fee"],
        discount_amount=row["discount_amount"],
        total_fee=row["total_fee"],
        promo_code=row["promo_code"],
        payment_method=row["payment_method"],
        transaction_id=row["transaction_id"],
        gate_pin=row["gate_pin"],
        qr_payload=row["qr_payload"],
        entrance_hint=row["entrance_hint"],
        created_at=row["created_at"],
    )
