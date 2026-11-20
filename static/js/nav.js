/**
 * Turn-by-turn floating navigation HUD and deterministic QR Gate Pass SVG generator.
 */

export function generateQrSvg(payload, size = 82) {
  const grid = 15;
  let hash = 2166136261;
  for (let i = 0; i < payload.length; i += 1) {
    hash ^= payload.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }

  const isFinder = (r, c) =>
    (r < 4 && c < 4) || (r < 4 && c >= grid - 4) || (r >= grid - 4 && c < 4);

  const cell = size / grid;
  const rects = [];

  for (let r = 0; r < grid; r += 1) {
    for (let c = 0; c < grid; c += 1) {
      let dark = false;
      if (isFinder(r, c)) {
        const lr = r < 4 ? r : r - (grid - 4);
        const lc = c < 4 ? c : c - (grid - 4);
        dark = lr === 0 || lr === 3 || lc === 0 || lc === 3 || (lr >= 1 && lr <= 2 && lc >= 1 && lc <= 2);
      } else {
        const bit = (hash ^ (r * 31 + c * 17)) & 3;
        dark = bit === 0 || bit === 2;
      }
      if (dark) {
        rects.push(`<rect x="${(c * cell).toFixed(1)}" y="${(r * cell).toFixed(1)}" width="${cell.toFixed(1)}" height="${cell.toFixed(1)}" fill="#14110f"/>`);
      }
    }
  }

  return `<svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}" role="img" aria-label="Gate QR Code">${rects.join("")}</svg>`;
}

export function renderNavFloat(container, route, stepIndex, isSimulating, handlers) {
  if (!route) {
    container.classList.add("hidden");
    container.innerHTML = "";
    return;
  }

  container.classList.remove("hidden");
  const step = route.steps[Math.min(stepIndex, route.steps.length - 1)];
  const arrow =
    step.maneuver.includes("right")
      ? "↱"
      : step.maneuver.includes("left")
        ? "↰"
        : step.maneuver === "arrive-gate"
          ? "🏁"
          : "↑";

  container.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;">
      <div style="display:flex;align-items:center;gap:10px;min-width:0;">
        <div style="width:36px;height:36px;border-radius:10px;background:rgba(255,122,0,0.18);color:var(--accent-saffron);display:grid;place-items:center;font-size:18px;font-weight:800;flex-shrink:0;">
          ${arrow}
        </div>
        <div style="min-width:0;">
          <div style="font-size:13px;font-weight:700;">${step.instruction}</div>
          <div class="sub-txt">${step.distance_m}m • ${route.eta_mins} min (${route.distance_km} km)</div>
        </div>
      </div>
      <div style="display:flex;gap:6px;flex-shrink:0;">
        <button id="nav-sim-btn" class="btn btn-primary btn-sm">${isSimulating ? "Driving..." : "▶ Drive"}</button>
        <button id="nav-close-btn" class="btn btn-soft btn-sm">✕</button>
      </div>
    </div>
  `;

  container.querySelector("#nav-sim-btn").addEventListener("click", handlers.onSimulateDrive);
  container.querySelector("#nav-close-btn").addEventListener("click", handlers.onStopNav);
}
