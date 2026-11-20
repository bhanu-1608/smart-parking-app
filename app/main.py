from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Query, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from . import database as db
from .models import (
    DestinationPoint,
    FacilityDetail,
    FacilitySummary,
    NavigationResponse,
    NotificationOut,
    PriceQuoteRequest,
    PriceQuoteResponse,
    ReservationCreate,
    ReservationExtend,
    ReservationOut,
    VehicleCreate,
    VehicleOut,
    calculate_price_quote,
)
from .simulator import build_turn_by_turn_route, hub, sensor_heartbeat_loop

STATIC_DIR = Path(__file__).resolve().parents[1] / "static"


@asynccontextmanager
async def lifespan(_: FastAPI):
    db.init_db()
    task = asyncio.create_task(sensor_heartbeat_loop(4.5))
    try:
        yield
    finally:
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass


app = FastAPI(title="DeccanPulse Hyderabad Smart Parking API", version="4.0.0", lifespan=lifespan)

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/")
async def serve_index():
    index_path = STATIC_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="Frontend bundle not found")
    return FileResponse(str(index_path))


@app.get("/api/health")
def health_check():
    return {"status": "ok", "service": "deccanpulse", "city": "Hyderabad"}


@app.get("/api/destinations", response_model=List[DestinationPoint])
def get_destinations(
    search: str = Query(""),
    user_x: Optional[float] = Query(None),
    user_y: Optional[float] = Query(None),
):
    return db.search_destinations(search, user_x=user_x, user_y=user_y)


@app.get("/api/facilities", response_model=List[FacilitySummary])
def get_facilities(
    user_x: float = Query(585.0),
    user_y: float = Query(217.0),
    search: str = Query(""),
    only_ev: bool = Query(False),
    only_two_wheeler: bool = Query(False),
    only_covered: bool = Query(False),
    max_distance_km: Optional[float] = Query(None),
    dest_x: Optional[float] = Query(None),
    dest_y: Optional[float] = Query(None),
):
    return db.list_facilities(
        user_x=user_x,
        user_y=user_y,
        search=search,
        only_ev=only_ev,
        only_two_wheeler=only_two_wheeler,
        only_covered=only_covered,
        max_distance_km=max_distance_km,
        dest_x=dest_x,
        dest_y=dest_y,
    )


@app.get("/api/facilities/{facility_id}", response_model=FacilityDetail)
def get_facility(
    facility_id: str,
    user_x: float = Query(585.0),
    user_y: float = Query(217.0),
):
    detail = db.get_facility_detail(facility_id=facility_id, user_x=user_x, user_y=user_y)
    if not detail:
        raise HTTPException(status_code=404, detail="Parking hub not found")
    return detail


@app.post("/api/facilities/{facility_id}/watchlist")
def toggle_facility_watchlist(facility_id: str):
    try:
        watchlisted = db.toggle_watchlist(facility_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"facility_id": facility_id, "watchlisted": watchlisted}


@app.get("/api/navigation/{facility_id}", response_model=NavigationResponse)
def get_navigation(
    facility_id: str,
    user_x: float = Query(585.0),
    user_y: float = Query(217.0),
):
    detail = db.get_facility_detail(facility_id=facility_id, user_x=user_x, user_y=user_y)
    if not detail:
        raise HTTPException(status_code=404, detail="Parking hub not found")
    return build_turn_by_turn_route(
        user_x=user_x,
        user_y=user_y,
        fac_id=detail.id,
        fac_name=detail.name,
        fac_x=detail.x,
        fac_y=detail.y,
        fac_lat=detail.lat,
        fac_lng=detail.lng,
        entrance_hint=detail.entrance_hint,
    )


@app.get("/api/vehicles", response_model=List[VehicleOut])
def get_vehicles():
    return db.list_vehicles()


@app.post("/api/vehicles", response_model=VehicleOut)
def add_vehicle(payload: VehicleCreate):
    return db.create_vehicle(payload)


@app.delete("/api/vehicles/{vehicle_id}")
def remove_vehicle(vehicle_id: str):
    db.delete_vehicle(vehicle_id)
    return {"deleted": vehicle_id}


@app.post("/api/quote", response_model=PriceQuoteResponse)
def get_quote(payload: PriceQuoteRequest):
    detail = db.get_facility_detail(payload.facility_id)
    if not detail:
        raise HTTPException(status_code=404, detail="Parking area not found")
    try:
        return calculate_price_quote(
            hourly_rate=detail.hourly_rate,
            two_wheeler_rate=detail.two_wheeler_rate,
            ev_kwh_rate=detail.ev_kwh_rate,
            ev_2w_kwh_rate=detail.ev_2w_kwh_rate,
            duration_mins=payload.duration_mins,
            zone=payload.zone,
            vehicle_category=payload.vehicle_category,
            charge_ev=payload.charge_ev,
            promo_code=payload.promo_code,
            include_parking=payload.include_parking,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/reservations", response_model=List[ReservationOut])
def get_reservations(status: Optional[str] = Query(None)):
    return db.list_reservations(status_filter=status)


@app.post("/api/reservations", response_model=ReservationOut)
async def create_reservation(payload: ReservationCreate):
    try:
        created = db.create_reservation(payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    await hub.broadcast({"event": "reservation_created", "facility_id": created.facility_id, "bay_id": created.bay_id})
    return created


@app.post("/api/reservations/{res_id}/enter", response_model=ReservationOut)
async def enter_gate(res_id: str):
    try:
        updated = db.enter_reservation(res_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    await hub.broadcast({"event": "gate_entered", "facility_id": updated.facility_id, "bay_id": updated.bay_id})
    return updated


@app.post("/api/reservations/{res_id}/exit", response_model=ReservationOut)
async def exit_gate(res_id: str):
    try:
        updated = db.exit_reservation(res_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    await hub.broadcast({"event": "gate_exited", "facility_id": updated.facility_id, "bay_id": updated.bay_id})
    return updated


@app.post("/api/reservations/{res_id}/extend", response_model=ReservationOut)
async def extend_reservation(res_id: str, payload: ReservationExtend):
    try:
        updated = db.extend_reservation(res_id, payload.extra_mins, payload.payment_method)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    await hub.broadcast({"event": "reservation_updated", "reservation_id": res_id})
    return updated


@app.post("/api/reservations/{res_id}/cancel", response_model=ReservationOut)
async def cancel_reservation(res_id: str):
    try:
        updated = db.cancel_reservation(res_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    await hub.broadcast({"event": "reservation_updated", "facility_id": updated.facility_id})
    return updated


@app.get("/api/notifications", response_model=List[NotificationOut])
def get_notifications():
    return db.list_notifications()


@app.post("/api/notifications/read")
def read_notifications(notif_id: Optional[str] = Query(None)):
    db.mark_notifications_read(notif_id)
    return {"ok": True}


@app.post("/api/simulate/pulse")
async def trigger_sensor_pulse():
    pulse = db.simulate_sensor_pulse()
    if pulse:
        await hub.broadcast(pulse)
    return pulse or {"event": "noop"}


@app.post("/api/simulate/expiry-alert")
async def trigger_expiry_alert(reservation_id: Optional[str] = Query(None)):
    notif = db.create_notification(
        notif_type="expiry",
        title="Parking Expires in 10 Mins",
        message="Extend your parking time now to avoid extra charges at exit.",
        reservation_id=reservation_id,
        action_label="Extend +30m",
    )
    await hub.broadcast({"event": "notification_created", "notification": notif.model_dump()})
    return notif


@app.websocket("/ws/live")
async def websocket_live(ws: WebSocket):
    await hub.connect(ws)
    try:
        while True:
            await ws.receive_text()
    except WebSocketDisconnect:
        hub.disconnect(ws)
