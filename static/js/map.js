/**
 * True North-Up Greater Hyderabad Geographic Map Renderer.
 * Coordinate Bounds: Longitude 78.24..78.64 E -> X 0..1000; Latitude 17.21..17.52 N -> Y 700..0.
 */

function el(tag, attrs = {}) {
  const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
  return node;
}

export function projLatLng(lat, lng) {
  const x = ((lng - 78.24) / (78.64 - 78.24)) * 1000;
  const y = 700 - ((lat - 17.21) / (17.52 - 17.21)) * 700;
  return [Number(x.toFixed(1)), Number(y.toFixed(1))];
}

function polyPts(coords) {
  return coords.map(([la, ln]) => projLatLng(la, ln).join(",")).join(" ");
}

export function renderHyderabadMap(svg, state, handlers) {
  const { zoom, panX, panY } = state.mapView;
  const w = 1000 / zoom;
  const h = 700 / zoom;
  const vx = Math.max(0, Math.min(1000 - w, 500 - w / 2 + panX));
  const vy = Math.max(0, Math.min(700 - h, 350 - h / 2 + panY));
  svg.setAttribute("viewBox", `${vx.toFixed(1)} ${vy.toFixed(1)} ${w.toFixed(1)} ${h.toFixed(1)}`);
  svg.innerHTML = "";

  const base = el("g");
  base.appendChild(el("rect", { x: 0, y: 0, width: 1000, height: 700, fill: "var(--map-city)" }));

  // 1) Nehru Outer Ring Road (ORR) wrapping Greater Hyderabad (true coordinates)
  const orrCoords = [
    [17.508, 78.305], // Patancheru NW
    [17.515, 78.445], // Medchal / Kompally North
    [17.505, 78.555], // Shamirpet NE
    [17.445, 78.625], // Ghatkesar East
    [17.345, 78.615], // Pedda Amberpet SE
    [17.255, 78.545], // Bongloor South-East
    [17.225, 78.425], // Shamshabad / RGIA South
    [17.285, 78.345], // Rajendranagar SW
    [17.395, 78.315], // Narsingi / Kokapet West
    [17.465, 78.295], // Kollur West
    [17.508, 78.305],
  ];
  base.appendChild(
    el("polyline", {
      points: polyPts(orrCoords),
      fill: "none",
      stroke: "var(--map-orr)",
      "stroke-width": 10,
      "stroke-linejoin": "round",
    })
  );

  // 2) Twin Reservoirs (Osman Sagar & Himayat Sagar) + Mir Alam Tank + Hussain Sagar + Durgam Cheruvu + Saroornagar
  const lakes = [
    // Osman Sagar (Gandipet West)
    [
      [17.392, 78.285], [17.396, 78.315], [17.372, 78.322], [17.358, 78.300], [17.374, 78.278],
    ],
    // Himayat Sagar (SW)
    [
      [17.338, 78.345], [17.344, 78.374], [17.316, 78.378], [17.304, 78.352], [17.320, 78.336],
    ],
    // Hussain Sagar Lake (Center — iconic N-S heart/teardrop between Hyd & Secbad)
    [
      [17.436, 78.468], [17.434, 78.482], [17.420, 78.485], [17.409, 78.474], [17.416, 78.463],
    ],
    // Durgam Cheruvu (HITEC / Jubilee)
    [
      [17.436, 78.387], [17.437, 78.395], [17.428, 78.396], [17.427, 78.388],
    ],
    // Mir Alam Tank (South of Musi near Zoo)
    [
      [17.354, 78.436], [17.355, 78.454], [17.344, 78.453], [17.343, 78.437],
    ],
    // Saroornagar Lake (SE near LB Nagar)
    [
      [17.359, 78.526], [17.358, 78.538], [17.348, 78.537], [17.349, 78.525],
    ],
  ];
  for (const poly of lakes) {
    base.appendChild(
      el("polygon", {
        points: polyPts(poly),
        fill: "var(--map-water)",
        stroke: "rgba(0, 212, 170, 0.35)",
        "stroke-width": 1.6,
      })
    );
  }

  // KBR National Park (Jubilee / Banjara Hills)
  base.appendChild(
    el("polygon", {
      points: polyPts([[17.426, 78.414], [17.428, 78.428], [17.414, 78.430], [17.412, 78.416]]),
      fill: "var(--map-park)",
      stroke: "rgba(0, 212, 170, 0.25)",
      "stroke-width": 1.5,
    })
  );

  // 3) Real Musi River flowing West (Osman/Himayat Sagar) -> East (Afzalgunj -> Amberpet -> Nagole -> East)
  const musiNorth = [
    [17.376, 78.320], [17.368, 78.385], [17.373, 78.445], [17.377, 78.485], [17.379, 78.525], [17.373, 78.565], [17.365, 78.625],
  ];
  base.appendChild(
    el("polyline", {
      points: polyPts(musiNorth),
      fill: "none",
      stroke: "var(--map-water)",
      "stroke-width": 11,
      "stroke-linecap": "round",
      "stroke-linejoin": "round",
    })
  );

  // 4) Major Highways & Expressways (PVNR Expy, NH-65, NH-44, Inner Ring Rd, HITEC-Gachibowli)
  const roads = [
    // NH-65 (Miyapur -> KPHB -> Ameerpet -> Panjagutta -> Nampally -> Dilsukhnagar -> LB Nagar -> East)
    [[17.505, 78.345], [17.486, 78.391], [17.437, 78.448], [17.406, 78.469], [17.378, 78.481], [17.345, 78.552], [17.330, 78.610]],
    // NH-44 North-South Spine (Kompally -> Secunderabad -> Tank Bund -> Charminar -> Shamshabad RGIA)
    [[17.515, 78.485], [17.445, 78.498], [17.420, 78.482], [17.378, 78.481], [17.350, 78.472], [17.245, 78.430]],
    // PVNR Expressway (Mehdipatnam -> Attapur -> Aramghar -> RGIA Airport)
    [[17.395, 78.438], [17.365, 78.425], [17.320, 78.425], [17.245, 78.430]],
    // Inner Ring Road East/South (Tarnaka -> Habsiguda -> Uppal -> Nagole -> LB Nagar -> Aramghar)
    [[17.435, 78.501], [17.410, 78.550], [17.372, 78.559], [17.345, 78.552], [17.328, 78.475], [17.320, 78.425]],
    // West Tech Corridor (Kondapur -> HITEC -> Mindspace -> Gachibowli -> Financial Dist)
    [[17.465, 78.355], [17.450, 78.381], [17.434, 78.385], [17.418, 78.344], [17.410, 78.320]],
    // Banjara/Jubilee Spine (HITEC -> Jubilee Checkpost -> Banjara Rd 12 -> Panjagutta)
    [[17.442, 78.377], [17.429, 78.409], [17.413, 78.444], [17.431, 78.451]],
  ];
  for (const rCoords of roads) {
    base.appendChild(
      el("polyline", {
        points: polyPts(rCoords),
        fill: "none",
        stroke: "var(--map-road)",
        "stroke-width": 5.5,
        "stroke-linecap": "round",
        "stroke-linejoin": "round",
      })
    );
  }

  // 5) Real Hyderabad Metro Rail Lines (Red, Blue, Green) + Station Nodes & Interchanges
  const metros = [
    // Red Line Corridor I: Miyapur -> KPHB -> Ameerpet -> Nampally -> MGBS -> Dilsukhnagar -> LB Nagar
    [
      [
        [17.496, 78.362], [17.493, 78.378], [17.484, 78.389], [17.466, 78.406], [17.445, 78.436],
        [17.437, 78.448], [17.431, 78.451], [17.412, 78.461], [17.405, 78.463], [17.391, 78.469],
        [17.378, 78.481], [17.369, 78.524], [17.346, 78.552],
      ],
      "#e11d48",
    ],
    // Blue Line Corridor III: Raidurg -> HITEC -> Jubilee -> Ameerpet -> Begumpet -> Secbad -> Uppal -> Nagole
    [
      [
        [17.442, 78.377], [17.450, 78.381], [17.442, 78.393], [17.429, 78.409], [17.435, 78.432],
        [17.437, 78.448], [17.444, 78.467], [17.443, 78.487], [17.435, 78.501], [17.427, 78.532],
        [17.419, 78.540], [17.402, 78.559], [17.373, 78.559],
      ],
      "#0284c7",
    ],
    // Green Line Corridor II: JBS Parade Ground -> Secbad -> RTC X Roads -> Sultan Bazar -> MGBS
    [
      [
        [17.444, 78.498], [17.435, 78.501], [17.407, 78.497], [17.395, 78.491], [17.386, 78.485], [17.378, 78.481],
      ],
      "#059669",
    ],
  ];

  for (const [mCoords, color] of metros) {
    const pts = polyPts(mCoords);
    base.appendChild(
      el("polyline", {
        points: pts,
        fill: "none",
        stroke: "rgba(255, 255, 255, 0.88)",
        "stroke-width": 5.2,
        "stroke-linecap": "round",
        "stroke-linejoin": "round",
      })
    );
    base.appendChild(
      el("polyline", {
        points: pts,
        fill: "none",
        stroke: color,
        "stroke-width": 3.0,
        "stroke-linecap": "round",
        "stroke-linejoin": "round",
      })
    );
    for (const [sla, sln] of mCoords) {
      const [sx, sy] = projLatLng(sla, sln);
      base.appendChild(el("circle", { cx: sx, cy: sy, r: 3.1, fill: "#ffffff", stroke: color, "stroke-width": 1.6 }));
    }
  }

  // Major Metro Interchanges (Ameerpet, MGBS, Secunderabad)
  const interchanges = [
    [17.437, 78.448, "Ameerpet (Red·Blue)"],
    [17.378, 78.481, "MGBS (Red·Green)"],
    [17.435, 78.501, "Secbad Metro"],
  ];
  for (const [ila, iln, ilabel] of interchanges) {
    const [ix, iy] = projLatLng(ila, iln);
    base.appendChild(el("circle", { cx: ix, cy: iy, r: 5.5, fill: "#ffffff", stroke: "#18202c", "stroke-width": 2.0 }));
    const it = el("text", { x: ix + 8, y: iy - 7, fill: "#18202c", "font-family": "var(--font-mono)", "font-size": 11.5, "font-weight": 800 });
    it.textContent = ilabel;
    base.appendChild(it);
  }

  // 6) Clean Geographic Landmarks (uncluttered, spacious canvas)
  const textPlaces = [
    [17.422, 78.474, "HUSSAIN SAGAR"],
    [17.374, 78.300, "OSMAN SAGAR"],
    [17.320, 78.360, "HIMAYAT SAGAR"],
    [17.374, 78.515, "MUSI RIVER"],
    [17.420, 78.421, "KBR PARK"],
    [17.510, 78.415, "NEHRU ORR"],
  ];
  for (const [la, ln, label] of textPlaces) {
    const [tx, ty] = projLatLng(la, ln);
    const t = el("text", {
      x: tx,
      y: ty,
      "text-anchor": "middle",
      fill: "var(--text-dim)",
      opacity: "0.75",
      "font-family": "var(--font-phase-map)",
      "font-size": 11.5,
      "font-weight": 700,
    });
    t.textContent = label;
    base.appendChild(t);
  }

  svg.appendChild(base);

  // Route & Destination Overlay
  const overlay = el("g");
  if (state.activeRoute?.path_points?.length) {
    const pts = state.activeRoute.path_points.map((p) => `${p[0]},${p[1]}`).join(" ");
    overlay.appendChild(
      el("polyline", { points: pts, fill: "none", stroke: "rgba(184, 74, 34, 0.28)", "stroke-width": 9, "stroke-linecap": "round" })
    );
    overlay.appendChild(
      el("polyline", { points: pts, fill: "none", stroke: "#b84a22", "stroke-width": 4, "stroke-dasharray": "9 7", "stroke-linecap": "round" })
    );
  }

  if (state.activeDestination) {
    const d = state.activeDestination;
    const dg = el("g", { transform: `translate(${d.x}, ${d.y})` });
    dg.appendChild(el("circle", { cx: 0, cy: 0, r: 22, fill: "rgba(154, 101, 21, 0.22)" }));
    dg.appendChild(el("circle", { cx: 0, cy: 0, r: 8.5, fill: "#9a6515", stroke: "#ffffff", "stroke-width": 2.2 }));
    const dt = el("text", {
      x: 0,
      y: -15,
      "text-anchor": "middle",
      fill: "var(--brand-indigo)",
      "paint-order": "stroke",
      stroke: "var(--map-city)",
      "stroke-width": 3,
      "font-family": "var(--font-phase-map)",
      "font-size": 12.5,
      "font-weight": 800,
    });
    dt.textContent = d.name;
    dg.appendChild(dt);
    overlay.appendChild(dg);
  }
  svg.appendChild(overlay);

  // Clean Uncluttered Parking Symbols (P badge + Area Name only — all details in side panel)
  const pins = el("g");
  for (const fac of state.facilities) {
    const sel = fac.id === state.selectedFacilityId;
    const open = fac.available_bays;
    const col = open >= 5 ? "#0d8278" : open >= 2 ? "#9a6515" : "#be123c";

    const g = el("g", { transform: `translate(${fac.x}, ${fac.y})`, style: "cursor:pointer;" });
    if (sel) g.appendChild(el("circle", { cx: 0, cy: 0, r: 20, fill: "rgba(184, 74, 34, 0.16)", stroke: "#b84a22", "stroke-width": 2 }));

    g.appendChild(
      el("rect", {
        x: -12,
        y: -12,
        width: 24,
        height: 24,
        rx: 7,
        fill: sel ? "#b84a22" : "var(--bg-card)",
        stroke: sel ? "#ffffff" : col,
        "stroke-width": 2.2,
      })
    );
    const sym = el("text", {
      x: 0,
      y: 4.5,
      "text-anchor": "middle",
      fill: sel ? "#ffffff" : col,
      "font-family": "var(--font-brand-sans)",
      "font-size": 13.5,
      "font-weight": 800,
    });
    sym.textContent = "P";
    g.appendChild(sym);

    if (fac.metro_lines && fac.metro_lines.length > 0) {
      const mColor = fac.metro_lines[0] === "Red" ? "#e11d48" : fac.metro_lines[0] === "Blue" ? "#0284c7" : "#059669";
      g.appendChild(el("circle", { cx: 10, cy: -10, r: 4.2, fill: mColor, stroke: "#ffffff", "stroke-width": 1.3 }));
    }

    const lbl = el("text", {
      x: 0,
      y: 23,
      "text-anchor": "middle",
      fill: sel ? "#b84a22" : "var(--brand-indigo)",
      "paint-order": "stroke",
      stroke: "var(--map-city)",
      "stroke-width": 3.2,
      "font-family": "var(--font-phase-directory)",
      "font-size": 12,
      "font-weight": 700,
    });
    lbl.textContent = fac.name;
    g.appendChild(lbl);

    g.addEventListener("click", (e) => {
      e.stopPropagation();
      handlers.onSelectFacility(fac.id);
    });
    pins.appendChild(g);
  }
  svg.appendChild(pins);

  // Show Driver GPS Marker only when Location is Enabled OR Navigating
  if (state.locationEnabled || state.navSimPos) {
    const pos = state.navSimPos || state.userPos;
    const u = el("g", { transform: `translate(${pos.x}, ${pos.y})` });
    u.appendChild(el("circle", { cx: 0, cy: 0, r: 15, fill: "rgba(184, 74, 34, 0.24)" }));
    u.appendChild(el("circle", { cx: 0, cy: 0, r: 8, fill: "#b84a22", stroke: "#ffffff", "stroke-width": 2.5 }));
    const ut = el("text", { x: 0, y: -15, "text-anchor": "middle", fill: "var(--text-main)", "font-family": "var(--font-mono)", "font-size": 12, "font-weight": 800 });
    ut.textContent = state.navSimPos ? "Driving" : "You";
    u.appendChild(ut);
    svg.appendChild(u);
  }
}
