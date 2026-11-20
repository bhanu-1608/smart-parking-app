from __future__ import annotations

import asyncio
import json
from typing import List, Set
from fastapi import WebSocket

from .database import simulate_sensor_pulse
from .models import NavigationResponse, RouteStep, canvas_to_lat_lng, compute_haversine_road_km


class SensorHub:
    """Manages connected WebSocket clients and broadcasts live bay telemetry."""

    def __init__(self) -> None:
        self.clients: Set[WebSocket] = set()

    async def connect(self, ws: WebSocket) -> None:
        await ws.accept()
        self.clients.add(ws)

    def disconnect(self, ws: WebSocket) -> None:
        self.clients.discard(ws)

    async def broadcast(self, payload: dict) -> None:
        if not self.clients:
            return
        message = json.dumps(payload)
        stale = []
        for client in list(self.clients):
            try:
                await client.send_text(message)
            except Exception:
                stale.append(client)
        for client in stale:
            self.clients.discard(client)


hub = SensorHub()


async def sensor_heartbeat_loop(interval_seconds: float = 4.5) -> None:
    """Periodically simulate realistic booking/entry/exit activity and push via WebSocket."""
    while True:
        await asyncio.sleep(interval_seconds)
        pulse = simulate_sensor_pulse()
        if pulse:
            await hub.broadcast(pulse)


def build_turn_by_turn_route(
    user_x: float,
    user_y: float,
    fac_id: str,
    fac_name: str,
    fac_x: float,
    fac_y: float,
    fac_lat: float,
    fac_lng: float,
    entrance_hint: str,
) -> NavigationResponse:
    """Generate Hyderabad turn-by-turn road steps and waypoints from GPS to hub."""
    user_lat, user_lng = canvas_to_lat_lng(user_x, user_y)
    dist_km = compute_haversine_road_km(user_lat, user_lng, fac_lat, fac_lng)
    eta_mins = max(3, round(dist_km * 2.6 + 2))

    # Two-bend realistic arterial path across Hyderabad grid
    p1 = [round((user_x + fac_x) * 0.5, 1), round(user_y, 1)]
    p2 = [round((user_x + fac_x) * 0.5, 1), round(fac_y, 1)]
    path_points = [[round(user_x, 1), round(user_y, 1)], p1, p2, [round(fac_x, 1), round(fac_y, 1)]]

    horiz = "East" if fac_x >= user_x else "West"
    vert = "South" if fac_y >= user_y else "North"
    d1_m = max(180, round(dist_km * 480))
    d2_m = max(150, round(dist_km * 420))

    arterial_1 = "HITEC Main Rd / Jubilee Rd" if user_x < 450 else "Necklace Rd / Tank Bund Corridor"
    arterial_2 = "Inner Ring Rd / Road No. 12" if fac_y < 420 else "PVNR Expy / Charminar Link"

    turn_dir = "right" if (
        (horiz == "East" and vert == "South") or (horiz == "West" and vert == "North")
    ) else "left"

    steps: List[RouteStep] = [
        RouteStep(
            instruction=f"Drive {horiz} along {arterial_1}",
            distance_m=d1_m,
            maneuver=f"straight-{horiz.lower()}",
            x=p1[0],
            y=p1[1],
        ),
        RouteStep(
            instruction=f"Turn {turn_dir} onto {arterial_2} ({vert})",
            distance_m=d2_m,
            maneuver=f"turn-{turn_dir}",
            x=p2[0],
            y=p2[1],
        ),
        RouteStep(
            instruction=entrance_hint,
            distance_m=80,
            maneuver="arrive-gate",
            x=round(fac_x, 1),
            y=round(fac_y, 1),
        ),
    ]

    return NavigationResponse(
        facility_id=fac_id,
        facility_name=fac_name,
        distance_km=dist_km,
        eta_mins=eta_mins,
        entrance_hint=entrance_hint,
        path_points=path_points,
        steps=steps,
    )
