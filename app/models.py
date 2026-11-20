from __future__ import annotations

import math
from datetime import datetime, timezone
from typing import List, Literal, Optional
from pydantic import BaseModel, Field


BayZone = Literal["regular_deck", "two_wheeler_floor", "ev_charging_zone"]
BayVehicleCategory = Literal["car", "two_wheeler"]
BayStatus = Literal["available", "occupied", "reserved"]
VehicleKind = Literal["Hatchback", "Sedan", "SUV", "EV Car", "Scooty / Bike", "EV Scooty"]
ReservationStatus = Literal["reserved", "entered", "completed", "cancelled"]
NotificationType = Literal["booking", "expiry", "availability", "payment", "system"]

LNG_MIN, LNG_MAX = 78.24, 78.64
LAT_MIN, LAT_MAX = 17.21, 17.52


def round_inr_5(amount: float, minimum: int = 0) -> int:
    """Round any Rupee amount to a clean multiple of 5."""
    if amount <= 0 and minimum == 0:
        return 0
    rounded = int(round(amount / 5.0) * 5)
    return max(minimum, rounded)


def lat_lng_to_canvas(lat: float, lng: float) -> tuple[float, float]:
    """True North-Up geographic projection of Greater Hyderabad onto 1000x700 canvas."""
    x = ((lng - LNG_MIN) / (LNG_MAX - LNG_MIN)) * 1000.0
    y = 700.0 - ((lat - LAT_MIN) / (LAT_MAX - LAT_MIN)) * 700.0
    return round(x, 1), round(y, 1)


def canvas_to_lat_lng(x: float, y: float) -> tuple[float, float]:
    lng = LNG_MIN + (x / 1000.0) * (LNG_MAX - LNG_MIN)
    lat = LAT_MIN + ((700.0 - y) / 700.0) * (LAT_MAX - LAT_MIN)
    return round(lat, 5), round(lng, 5)


def compute_bearing_label(lat1: float, lng1: float, lat2: float, lng2: float) -> str:
    dlng = math.radians(lng2 - lng1)
    y = math.sin(dlng) * math.cos(math.radians(lat2))
    x = math.cos(math.radians(lat1)) * math.sin(math.radians(lat2)) - (
        math.sin(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.cos(dlng)
    )
    deg = (math.degrees(math.atan2(y, x)) + 360.0) % 360.0
    dirs = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
    return dirs[int((deg + 22.5) // 45) % 8]


class DestinationPoint(BaseModel):
    id: str
    name: str
    area: str
    category: str
    lat: float
    lng: float
    x: float
    y: float
    distance_km: Optional[float] = None
    drive_mins: Optional[int] = None
    bearing: Optional[str] = None


class FacilitySummary(BaseModel):
    id: str
    name: str
    tag: str
    district: str
    address: str
    x: float
    y: float
    lat: float
    lng: float
    hourly_rate: int
    two_wheeler_rate: int
    ev_kwh_rate: int
    ev_2w_kwh_rate: int
    covered: bool
    open_24_7: bool
    has_fastag: bool
    has_valet: bool
    ev_zone_name: str
    ev_charger_types: List[str]
    demand_profile: str
    rating: float
    entrance_hint: str
    metro_station: Optional[str] = None
    metro_lines: List[str] = Field(default_factory=list)
    amenities: List[str] = Field(default_factory=list)
    total_bays: int
    available_bays: int
    regular_total: int
    regular_available: int
    two_wheeler_total: int
    two_wheeler_available: int
    ev_total: int
    ev_available: int
    ev_car_total: int
    ev_car_available: int
    ev_2w_total: int
    ev_2w_available: int
    effective_rate: int
    distance_km: float
    drive_mins: int
    bearing: str
    walk_to_dest_mins: Optional[int] = None
    watchlisted: bool = False


class BayOut(BaseModel):
    id: str
    facility_id: str
    level: str
    code: str
    row_idx: int
    col_idx: int
    zone: BayZone
    vehicle_category: BayVehicleCategory = "car"
    charger_standard: str
    charger_kw: int
    walk_meters: int
    status: BayStatus
    covered: bool


class FacilityDetail(FacilitySummary):
    bays: List[BayOut]
    levels: List[str]


class VehicleOut(BaseModel):
    id: str
    make_model: str
    plate: str
    car_type: VehicleKind
    is_ev: bool
    charger_pref: str = "None"


class VehicleCreate(BaseModel):
    make_model: str = Field(..., min_length=2, max_length=60)
    plate: str = Field(..., min_length=4, max_length=18)
    car_type: VehicleKind = "Hatchback"
    is_ev: bool = False
    charger_pref: str = "None"


class ReservationCreate(BaseModel):
    facility_id: str
    bay_id: str
    vehicle_id: str
    duration_mins: int = Field(..., ge=30, le=720)
    payment_method: str = Field(..., min_length=2, max_length=60)
    promo_code: Optional[str] = None
    include_parking: bool = True
    charge_ev: bool = False


class ReservationExtend(BaseModel):
    extra_mins: int = Field(..., ge=30, le=240)
    payment_method: str = "UPI Instant"


class ReservationOut(BaseModel):
    id: str
    facility_id: str
    facility_name: str
    bay_id: str
    bay_code: str
    level: str
    zone: str
    vehicle_id: str
    vehicle_name: str
    vehicle_plate: str
    start_time: str
    end_time: str
    entered_at: Optional[str] = None
    exited_at: Optional[str] = None
    duration_mins: int
    status: ReservationStatus
    include_parking: bool = True
    base_fee: int
    ev_fee: int
    discount_amount: int
    total_fee: int
    promo_code: Optional[str] = None
    payment_method: str
    transaction_id: str
    gate_pin: str
    qr_payload: str
    entrance_hint: str
    created_at: str


class NotificationOut(BaseModel):
    id: str
    type: NotificationType
    title: str
    message: str
    facility_id: Optional[str] = None
    reservation_id: Optional[str] = None
    action_label: Optional[str] = None
    is_read: bool
    created_at: str


class RouteStep(BaseModel):
    instruction: str
    distance_m: int
    maneuver: str
    x: float
    y: float


class NavigationResponse(BaseModel):
    facility_id: str
    facility_name: str
    distance_km: float
    eta_mins: int
    entrance_hint: str
    path_points: List[List[float]]
    steps: List[RouteStep]


class PriceQuoteRequest(BaseModel):
    facility_id: str
    duration_mins: int = Field(..., ge=30, le=720)
    zone: BayZone = "regular_deck"
    vehicle_category: BayVehicleCategory = "car"
    include_parking: bool = True
    charge_ev: bool = False
    promo_code: Optional[str] = None


class PriceQuoteResponse(BaseModel):
    hourly_rate: int
    base_fee: int
    ev_fee: int
    discount_amount: int
    total_fee: int
    promo_applied: Optional[str] = None
    promo_note: Optional[str] = None


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def compute_haversine_road_km(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lng2 - lng1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    straight = 2 * r * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return round(max(0.25, straight * 1.28), 2)


def calculate_price_quote(
    hourly_rate: int,
    two_wheeler_rate: int,
    ev_kwh_rate: int,
    ev_2w_kwh_rate: int,
    duration_mins: int,
    zone: BayZone,
    vehicle_category: BayVehicleCategory,
    charge_ev: bool,
    promo_code: Optional[str],
    include_parking: bool = True,
) -> PriceQuoteResponse:
    """Compute separate ₹5-multiple pricing where Parking and EV Charging are independent."""
    is_two_wheeler = zone == "two_wheeler_floor" or vehicle_category == "two_wheeler"
    applied_rate = two_wheeler_rate if is_two_wheeler else hourly_rate
    hours = duration_mins / 60.0
    ev_active = bool(charge_ev and zone == "ev_charging_zone")

    if not include_parking and not ev_active:
        raise ValueError("Select at least Parking Space or EV Charging")

    base_fee = round_inr_5(applied_rate * hours, minimum=10) if include_parking else 0

    ev_fee = 0
    if ev_active:
        if is_two_wheeler:
            # 2-Wheeler EV battery pack is ~3-4 kWh total; realistic ~1.5 kWh/hr capped at 3 kWh
            kwh_est = min(hours * 1.5, 3.0)
            ev_fee = round_inr_5(ev_2w_kwh_rate * kwh_est, minimum=15)
        else:
            # EV Car urban top-up at hub ~3 kWh/hr capped at 8 kWh per parking session
            kwh_est = min(hours * 3.0, 8.0)
            ev_fee = round_inr_5(ev_kwh_rate * kwh_est, minimum=25)
    subtotal = base_fee + ev_fee

    code = (promo_code or "").strip().upper()
    discount = 0
    promo_applied = None
    promo_note = None

    if code == "HYDEV20":
        if charge_ev and zone == "ev_charging_zone":
            discount = round_inr_5(subtotal * 0.20, minimum=10)
            promo_applied = code
            promo_note = f"20% EV discount (-₹{discount}) applied"
        else:
            promo_note = "HYDEV20 applies to EV Charging slots"
    elif code == "FASTAG15":
        discount = round_inr_5(subtotal * 0.15, minimum=5)
        promo_applied = code
        promo_note = f"15% FASTag discount (-₹{discount}) applied"
    elif code == "NAMASTE50":
        discount = min(max(0, subtotal - 10), 50)
        promo_applied = code
        promo_note = f"₹{discount} welcome discount applied"
    elif code:
        promo_note = "Try HYDEV20, FASTAG15, or NAMASTE50"

    total_fee = round_inr_5(max(10, subtotal - discount), minimum=10)
    return PriceQuoteResponse(
        hourly_rate=applied_rate if include_parking else 0,
        base_fee=base_fee,
        ev_fee=ev_fee,
        discount_amount=discount,
        total_fee=total_fee,
        promo_applied=promo_applied,
        promo_note=promo_note,
    )
