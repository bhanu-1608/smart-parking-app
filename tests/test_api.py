from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

from fastapi import HTTPException

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))


class DeccanPulseTwoWheelerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.tmpdir = tempfile.TemporaryDirectory()
        cls.db_path = Path(cls.tmpdir.name) / "test_deccanpulse_v8.db"
        os.environ["BAYPULSE_DB_PATH"] = str(cls.db_path)

        from app import database as db  # noqa: E402
        from app import main  # noqa: E402
        from app import models  # noqa: E402

        db.init_db()
        cls.db = db
        cls.main = main
        cls.models = models

    @classmethod
    def tearDownClass(cls) -> None:
        os.environ.pop("BAYPULSE_DB_PATH", None)
        cls.tmpdir.cleanup()

    def test_01_hubs_include_car_scooty_bike_and_ev_floors(self) -> None:
        facs = self.main.get_facilities(
            user_x=585.0,
            user_y=217.0,
            search="",
            only_ev=False,
            only_two_wheeler=False,
            only_covered=False,
            max_distance_km=None,
            dest_x=None,
            dest_y=None,
        )
        self.assertEqual(len(facs), 21)
        fac_map = {f.id: f for f in facs}
        self.assertIn("fac-kompally", fac_map)
        self.assertIn("fac-ecil", fac_map)
        self.assertIn("fac-uppal", fac_map)
        self.assertIn("fac-miyapur", fac_map)
        self.assertIn("fac-narsingi", fac_map)
        self.assertEqual(fac_map["fac-kphb"].metro_lines, ["Red"])
        self.assertEqual(fac_map["fac-nagole"].metro_lines, ["Blue"])
        self.assertEqual(fac_map["fac-secbad"].metro_lines, ["Blue", "Green"])
        self.assertEqual(fac_map["fac-mgbs"].metro_lines, ["Red", "Green"])
        for f in facs:
            self.assertEqual(f.hourly_rate % 5, 0)
            self.assertEqual(f.two_wheeler_rate % 5, 0)
            self.assertEqual(f.ev_kwh_rate % 5, 0)
            self.assertEqual(f.ev_2w_kwh_rate % 5, 0)
            self.assertIn(f.hourly_rate, (30, 40, 50))
            self.assertIn(f.two_wheeler_rate, (10, 15))
            self.assertEqual(f.ev_kwh_rate, 15)
            self.assertEqual(f.ev_2w_kwh_rate, 10)
            self.assertGreaterEqual(f.regular_total, 20)
            self.assertGreaterEqual(f.two_wheeler_total, 16)
            self.assertGreaterEqual(f.ev_car_total, 3)
            self.assertGreaterEqual(f.ev_2w_total, 3)
            self.assertEqual(f.ev_total, f.ev_car_total + f.ev_2w_total)

        detail = self.main.get_facility(facility_id="fac-hitec", user_x=585.0, user_y=217.0)
        tw_bays = [b for b in detail.bays if b.zone == "two_wheeler_floor"]
        self.assertGreaterEqual(len(tw_bays), 20)
        self.assertTrue(any(b.code.startswith("2W-") for b in tw_bays))
        ev_car_bays = [b for b in detail.bays if b.zone == "ev_charging_zone" and b.vehicle_category == "car"]
        ev_2w_bays = [b for b in detail.bays if b.zone == "ev_charging_zone" and b.vehicle_category == "two_wheeler"]
        self.assertGreaterEqual(len(ev_car_bays), 3)
        self.assertGreaterEqual(len(ev_2w_bays), 3)
        self.assertTrue(all(b.code.startswith("EVC-") for b in ev_car_bays))
        self.assertTrue(all(b.code.startswith("EV2W-") for b in ev_2w_bays))

    def test_02_scooty_bike_booking_and_pricing_multiples_of_5(self) -> None:
        self.assertEqual(len(self.main.get_vehicles()), 0)

        scooty = self.main.add_vehicle(
            self.models.VehicleCreate(
                make_model="Honda Activa 6G",
                plate="TS 09 SC 2219",
                car_type="Scooty / Bike",
                is_ev=False,
                charger_pref="None",
            )
        )
        self.assertEqual(scooty.car_type, "Scooty / Bike")

        detail = self.main.get_facility(facility_id="fac-nagole", user_x=585.0, user_y=217.0)
        tw_slot = next(b for b in detail.bays if b.zone == "two_wheeler_floor" and b.status == "available")

        quote = self.main.get_quote(
            self.models.PriceQuoteRequest(
                facility_id="fac-nagole",
                duration_mins=120,
                zone="two_wheeler_floor",
                charge_ev=False,
                promo_code=None,
            )
        )
        self.assertEqual(quote.hourly_rate, detail.two_wheeler_rate)
        self.assertEqual(quote.total_fee % 5, 0)

        booked = self.db.create_reservation(
            self.models.ReservationCreate(
                facility_id="fac-nagole",
                bay_id=tw_slot.id,
                vehicle_id=scooty.id,
                duration_mins=120,
                payment_method="UPI Instant (GPay / PhonePe)",
                charge_ev=False,
            )
        )
        self.assertEqual(booked.zone, "two_wheeler_floor")
        self.assertEqual(booked.total_fee % 5, 0)

        entered = self.db.enter_reservation(booked.id)
        self.assertEqual(entered.status, "entered")

        exited = self.db.exit_reservation(booked.id)
        self.assertEqual(exited.status, "completed")

    def test_03_destination_search_and_pulse(self) -> None:
        dests = self.main.get_destinations(search="Nagole", user_x=585.0, user_y=217.0)
        self.assertEqual(len(dests), 1)
        self.assertEqual(dests[0].bearing, "SE")

        pulse = self.db.simulate_sensor_pulse()
        self.assertIsNotNone(pulse)
        self.assertIn(pulse["status"], ("available", "reserved", "occupied"))

        with self.assertRaises(HTTPException):
            self.main.get_facility(facility_id="missing", user_x=585.0, user_y=217.0)

    def test_04_ev_wing_car_and_two_wheeler_realistic_rates(self) -> None:
        ev_scooty = self.main.add_vehicle(
            self.models.VehicleCreate(
                make_model="Ather 450X",
                plate="TS 09 EV 1122",
                car_type="EV Scooty",
                is_ev=True,
                charger_pref="3.3kW AC",
            )
        )
        detail = self.main.get_facility(facility_id="fac-hitec", user_x=585.0, user_y=217.0)
        ev_2w_bay = next(
            b for b in detail.bays
            if b.zone == "ev_charging_zone" and b.vehicle_category == "two_wheeler" and b.status == "available"
        )
        ev_2w_quote = self.main.get_quote(
            self.models.PriceQuoteRequest(
                facility_id="fac-hitec",
                duration_mins=120,
                zone="ev_charging_zone",
                vehicle_category="two_wheeler",
                charge_ev=True,
                promo_code=None,
            )
        )
        self.assertEqual(ev_2w_quote.hourly_rate, detail.two_wheeler_rate)
        self.assertEqual(ev_2w_quote.base_fee, 30)
        self.assertEqual(ev_2w_quote.ev_fee, 30)
        self.assertEqual(ev_2w_quote.total_fee, 60)

        ev_car_quote = self.main.get_quote(
            self.models.PriceQuoteRequest(
                facility_id="fac-hitec",
                duration_mins=120,
                zone="ev_charging_zone",
                vehicle_category="car",
                charge_ev=True,
                promo_code=None,
            )
        )
        self.assertEqual(ev_car_quote.hourly_rate, detail.hourly_rate)
        self.assertEqual(ev_car_quote.base_fee, 80)
        self.assertEqual(ev_car_quote.ev_fee, 90)
        self.assertEqual(ev_car_quote.total_fee, 170)

        booked_ev2w = self.db.create_reservation(
            self.models.ReservationCreate(
                facility_id="fac-hitec",
                bay_id=ev_2w_bay.id,
                vehicle_id=ev_scooty.id,
                duration_mins=120,
                payment_method="UPI Instant (GPay / PhonePe)",
                include_parking=True,
                charge_ev=True,
            )
        )
        self.assertEqual(booked_ev2w.total_fee, 60)
        self.assertEqual(booked_ev2w.bay_code, ev_2w_bay.code)

    def test_05_separate_individual_charge_only_vs_parking_only(self) -> None:
        ev_car = self.main.add_vehicle(
            self.models.VehicleCreate(
                make_model="Tata Nexon.ev",
                plate="TS 09 EV 5544",
                car_type="EV Car",
                is_ev=True,
                charger_pref="CCS2 DC",
            )
        )
        detail = self.main.get_facility(facility_id="fac-uppal", user_x=585.0, user_y=217.0)
        ev_car_bay = next(
            b for b in detail.bays
            if b.zone == "ev_charging_zone" and b.vehicle_category == "car" and b.status == "available"
        )

        # EV Charge Only: do NOT include parking fees
        charge_only = self.main.get_quote(
            self.models.PriceQuoteRequest(
                facility_id="fac-uppal",
                duration_mins=120,
                zone="ev_charging_zone",
                vehicle_category="car",
                include_parking=False,
                charge_ev=True,
                promo_code=None,
            )
        )
        self.assertEqual(charge_only.hourly_rate, 0)
        self.assertEqual(charge_only.base_fee, 0)
        self.assertEqual(charge_only.ev_fee, 90)
        self.assertEqual(charge_only.total_fee, 90)

        booked_charge_only = self.db.create_reservation(
            self.models.ReservationCreate(
                facility_id="fac-uppal",
                bay_id=ev_car_bay.id,
                vehicle_id=ev_car.id,
                duration_mins=120,
                payment_method="UPI Instant (GPay / PhonePe)",
                include_parking=False,
                charge_ev=True,
            )
        )
        self.assertFalse(booked_charge_only.include_parking)
        self.assertEqual(booked_charge_only.base_fee, 0)
        self.assertEqual(booked_charge_only.ev_fee, 90)
        self.assertEqual(booked_charge_only.total_fee, 90)

        # Neither parking nor charging selected must raise HTTP 400
        with self.assertRaises(HTTPException):
            self.main.get_quote(
                self.models.PriceQuoteRequest(
                    facility_id="fac-uppal",
                    duration_mins=60,
                    zone="ev_charging_zone",
                    vehicle_category="car",
                    include_parking=False,
                    charge_ev=False,
                )
            )


if __name__ == "__main__":
    unittest.main()
