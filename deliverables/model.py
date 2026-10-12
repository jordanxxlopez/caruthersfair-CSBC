"""Monte Carlo planning model for the Caruthers Fair Band Showcase.

All inputs are planning assumptions (no published attendance or parking data was found).
Outputs: results.json, figures/*.png, and the numbers used in the report.
"""
import json, os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, "figures")
os.makedirs(FIG, exist_ok=True)
RNG = np.random.default_rng(20270925)
N = 20000

# name: (low, mode, high, source/basis)
INPUTS = {
    "bands": (10, 12, 16, "Planning target for field-show bands (10 to 16)"),
    "guests_per_band": (80, 150, 250, "Assumed family and guest count per band"),
    "sat_fair_attendance": (15000, 22000, 30000, "Saturday share of 60k to 80k four-day estimate (2021 organizer estimate)"),
    "walkup_share": (0.02, 0.05, 0.10, "Share of Saturday fairgoers who watch the field show"),
    "stand_capacity": (800, 1500, 2500, "Bleacher capacity at the high school (unverified)"),
    "occupancy": (2.2, 2.8, 3.2, "People per car (assumed)"),
    "car_share": (0.70, 0.85, 0.95, "Share of spectators arriving by car (assumed)"),
    "fair_lot_share": (0.30, 0.50, 0.80, "Share of football-field parking used by the fair on Saturday (assumed)"),
    "field_spaces": (250, 350, 450, "Football-field parking spaces (unverified, to measure)"),
    "elem_spaces": (80, 120, 160, "Elementary lot spaces (unverified, to measure)"),
    "hs_spaces": (80, 120, 160, "High school lot spaces (unverified, to measure)"),
    "trailer_share": (0.20, 0.40, 0.60, "Share of bands bringing a trailer or bus to the lot"),
    "band_vehicles": (1, 2, 3, "Vehicles per band that need a space"),
    "donation_per_spec": (0.10, 0.25, 0.60, "Donations and program sales per spectator (assumed)"),
    "sponsor": (0, 1500, 5000, "Sponsor or fair-board contribution (unknown)"),
    "band_fee": (0, 100, 250, "Participation fee per field-show band (policy choice)"),
}

BUDGET = {  # line: (low, high) from the report budget table; mode is midpoint
    "Field and facility use": (500, 1500),
    "Liability insurance": (500, 1500),
    "Traffic and crossing control": (1000, 2500),
    "Portable restrooms": (600, 1200),
    "Sound and staging": (600, 1500),
    "Judges (3 to 4)": (600, 1600),
    "Clinicians (2)": (800, 1600),
    "Plaques, programs, signage": (500, 1200),
    "Fire and EMS standby (added)": (0, 1000),
    "Trailer staging, fencing and signage (added)": (300, 1200),
}
CONTINGENCY = 0.15


def tri(lo, mode, hi, n=N):
    return RNG.triangular(lo, mode, hi, n)


def draw():
    d = {k: tri(*v[:3]) for k, v in INPUTS.items()}
    d["bands"] = np.round(d["bands"])
    d["band_vehicles"] = np.round(d["band_vehicles"])
    return d


def run():
    d = draw()
    band_guests = d["bands"] * d["guests_per_band"]
    walkup = d["walkup_share"] * d["sat_fair_attendance"]
    demand_spec = band_guests + walkup
    spectators = np.minimum(demand_spec, d["stand_capacity"])  # bleachers cap who can sit; overflow stands on grass uncounted

    cars = spectators * d["car_share"] / d["occupancy"]
    band_veh = d["bands"] * d["trailer_share"] * 0 + d["bands"] * d["band_vehicles"]
    lot_demand = cars + band_veh
    supply = (d["field_spaces"] * (1 - d["fair_lot_share"])) + d["elem_spaces"] + d["hs_spaces"]
    shortfall = np.maximum(lot_demand - supply, 0)

    costs = np.zeros(N)
    for lo, hi in BUDGET.values():
        costs += RNG.triangular(lo, (lo + hi) / 2, hi, N)
    total_cost = costs * (1 + CONTINGENCY)

    revenue = spectators * d["donation_per_spec"] + d["sponsor"] + d["bands"] * d["band_fee"]
    net = revenue - total_cost

    return {
        "d": d, "spectators": spectators, "demand_spec": demand_spec, "cars": cars,
        "lot_demand": lot_demand, "supply": supply, "shortfall": shortfall,
        "cost": total_cost, "revenue": revenue, "net": net,
    }


def pct(a):
    return {p: float(np.percentile(a, p)) for p in (10, 50, 90)}


def tornado(out):
    """One-at-a-time sensitivity on net result: swap each input to its low/high, hold others at mode."""
    base = {k: v[1] for k, v in INPUTS.items()}
    def net_from(p):
        bands = round(p["bands"]); spec = min(bands * p["guests_per_band"] + p["walkup_share"] * p["sat_fair_attendance"], p["stand_capacity"])
        rev = spec * p["donation_per_spec"] + p["sponsor"] + bands * p["band_fee"]
        cost = sum((lo + hi) / 2 for lo, hi in BUDGET.values()) * (1 + CONTINGENCY)
        return rev - cost
    b = net_from(base)
    rows = []
    for k in ["bands", "guests_per_band", "walkup_share", "sat_fair_attendance", "stand_capacity", "donation_per_spec", "sponsor", "band_fee"]:
        lo, hi = dict(zip(["lo", "hi"], (INPUTS[k][0], INPUTS[k][2])))["lo"], INPUTS[k][2]
        p1 = dict(base); p1[k] = INPUTS[k][0]
        p2 = dict(base); p2[k] = INPUTS[k][2]
        rows.append((k, net_from(p1) - b, net_from(p2) - b))
    rows.sort(key=lambda r: abs(r[2] - r[1]), reverse=True)
    return b, rows


def charts(o, t):
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    teal, gold, grey = "#0e5a57", "#b97a10", "#8a948f"

    fig, ax = plt.subplots(figsize=(7, 3.6))
    ax.hist(o["spectators"], bins=60, color=teal, alpha=0.9)
    p = pct(o["spectators"])
    for q, c in [(10, grey), (50, gold), (90, grey)]:
        ax.axvline(p[q], color=c, ls="--", lw=1.2)
        ax.text(p[q], ax.get_ylim()[1] * 0.92, f"P{q}: {p[q]:,.0f}", rotation=90, va="top", ha="right", fontsize=8, color=c)
    ax.set_xlabel("Field-show spectators (seated, capped at stand capacity)")
    ax.set_ylabel("Simulated runs")
    ax.set_title("Predicted spectators, 20,000 runs")
    fig.tight_layout(); fig.savefig(os.path.join(FIG, "fig_spectators.png"), dpi=200); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 3.8))
    labels = [r[0].replace("_", " ") for r in t[1]]
    lows = [r[1] for r in t[1]]; highs = [r[2] for r in t[1]]
    y = np.arange(len(labels))
    ax.barh(y, lows, color=grey, label="Low input")
    ax.barh(y, highs, color=teal, label="High input")
    ax.axvline(0, color="black", lw=0.8)
    ax.set_yticks(y); ax.set_yticklabels(labels); ax.invert_yaxis()
    ax.set_xlabel(f"Change in net result vs. base case of {t[0]:,.0f} (dollars)")
    ax.set_title("Sensitivity of net result to each input (one at a time)")
    ax.legend(frameon=False, fontsize=8, loc="lower right")
    fig.tight_layout(); fig.savefig(os.path.join(FIG, "fig_tornado.png"), dpi=200); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 3.4))
    ax.hist(o["lot_demand"], bins=60, color=gold, alpha=0.5, label="Vehicles needing space")
    ax.hist(o["supply"], bins=60, color=teal, alpha=0.5, label="Usable spaces")
    ax.set_xlabel("Vehicles and spaces"); ax.set_ylabel("Simulated runs")
    ax.set_title("Parking demand against usable supply")
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(os.path.join(FIG, "fig_parking.png"), dpi=200); plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 3.4))
    ax.hist(o["net"], bins=60, color=teal, alpha=0.9)
    ax.axvline(0, color="black", lw=1)
    ax.set_xlabel("Net result (revenue minus cost, dollars)"); ax.set_ylabel("Simulated runs")
    ax.set_title(f"Net result: {np.mean(o['net'] >= 0) * 100:.0f}% of runs break even or better")
    fig.tight_layout(); fig.savefig(os.path.join(FIG, "fig_net.png"), dpi=200); plt.close(fig)


if __name__ == "__main__":
    o = run()
    t = tornado(o)
    summary = {
        "runs": N,
        "spectators": pct(o["spectators"]),
        "band_guests_mean": float(np.mean(o["d"]["bands"] * o["d"]["guests_per_band"])),
        "cars": pct(o["cars"]),
        "lot_demand": pct(o["lot_demand"]),
        "supply": pct(o["supply"]),
        "shortfall_prob": float(np.mean(o["shortfall"] > 0)),
        "shortfall_p50": float(np.percentile(o["shortfall"], 50)),
        "shortfall_p90": float(np.percentile(o["shortfall"], 90)),
        "cost": pct(o["cost"]),
        "revenue": pct(o["revenue"]),
        "net": pct(o["net"]),
        "prob_net_nonneg": float(np.mean(o["net"] >= 0)),
        "tornado_base": t[0],
        "tornado": [{"input": r[0], "low_change": r[1], "high_change": r[2]} for r in t[1]],
        "inputs": {k: {"low": v[0], "mode": v[1], "high": v[2], "basis": v[3]} for k, v in INPUTS.items()},
        "budget_lines": {k: {"low": v[0], "high": v[1]} for k, v in BUDGET.items()},
        "contingency": CONTINGENCY,
    }
    with open(os.path.join(HERE, "results.json"), "w") as f:
        json.dump(summary, f, indent=2)
    # sample for the workbook (first 2,000 runs, to keep the file small)
    np.savez(os.path.join(HERE, "samples.npz"), spectators=o["spectators"][:2000], cars=o["cars"][:2000],
             lot_demand=o["lot_demand"][:2000], supply=o["supply"][:2000], cost=o["cost"][:2000],
             revenue=o["revenue"][:2000], net=o["net"][:2000])
    charts(o, t)
    print(json.dumps({k: summary[k] for k in ["spectators", "cars", "lot_demand", "supply", "shortfall_prob", "shortfall_p50", "shortfall_p90", "cost", "revenue", "net", "prob_net_nonneg", "band_guests_mean"]}, indent=1))
    print("tornado base", round(t[0]), [(r[0], round(r[1]), round(r[2])) for r in t[1]])
