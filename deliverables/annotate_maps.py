"""Annotated site maps on the labeled aerial (figures/base_labeled.png).
Coordinates are pixel positions read from the aerial; they are approximate and must be checked on the ground."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, FancyBboxPatch
import matplotlib.image as mpimg
import numpy as np

BASE = mpimg.imread("figures/base_clean.png")
TEAL, GOLD, RED, BLUE, GREY = "#0e5a57", "#e0a526", "#b5452b", "#2b5fa8", "#333333"

def canvas(title):
    fig, ax = plt.subplots(figsize=(10, 5.2))
    ax.imshow(BASE)
    ax.set_xlim(0, 2000); ax.set_ylim(943, 0); ax.axis("off")
    ax.set_title(title, fontsize=11, loc="left", color=GREY)
    for (x, y, t) in [(430, 300, "Caruthers High School"), (1120, 640, "Caruthers Elementary"), (790, 560, "Caruthers District Fair")]:
        ax.text(x, y, t, fontsize=7.5, color="black", ha="center", va="center",
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=0.85))
    return fig, ax

def note(fig, text):
    fig.text(0.01, 0.01, text, fontsize=8, color=GREY, ha="left", va="bottom")

def zone(ax, pts, color, label, lx, ly, alpha=0.22, ls="-"):
    ax.add_patch(Polygon(pts, closed=True, facecolor=color, alpha=alpha, edgecolor=color, lw=2, ls=ls))
    ax.text(lx, ly, label, fontsize=8.5, ha="center", va="center", color="white",
            bbox=dict(boxstyle="round,pad=0.3", fc=color, ec="none", alpha=0.95))

def route(ax, pts, color, label, lx, ly, ls="-"):
    xs, ys = zip(*pts)
    ax.plot(xs, ys, color=color, lw=3.2, ls=ls, solid_capstyle="round")
    ax.annotate("", xy=pts[-1], xytext=pts[-2], arrowprops=dict(arrowstyle="-|>", color=color, lw=3))
    ax.text(lx, ly, label, fontsize=8.5, ha="center", va="center", color="white",
            bbox=dict(boxstyle="round,pad=0.3", fc=color, ec="none", alpha=0.95))

# 1. Overview of the sites
fig, ax = canvas("Figure 5. Sites: high school, elementary, fairgrounds, and the open grass lot (aerial, approximate)")
zone(ax, [(300, 190), (480, 190), (480, 430), (300, 430)], TEAL, "High school", 390, 160)
zone(ax, [(200, 530), (370, 530), (370, 690), (200, 690)], TEAL, "HS field", 285, 610, alpha=0.15)
zone(ax, [(1000, 690), (1300, 690), (1300, 880), (1000, 880)], BLUE, "Elementary school", 1150, 905)
zone(ax, [(1200, 330), (1500, 330), (1500, 560), (1200, 560)], BLUE, "Community field", 1350, 300, alpha=0.15)
zone(ax, [(700, 470), (900, 470), (900, 640), (700, 640)], GOLD, "Fairgrounds", 800, 440)
zone(ax, [(640, 600), (820, 600), (820, 790), (640, 790)], RED, "Open grass lot\n(fair parking)", 730, 690, alpha=0.2)
ax.text(1140, 860, "", fontsize=1)
note(fig, "Sites approximate, read from the aerial. Open grass lot is the fair's parking across from the high school, next to the fire station.")
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig("figures/map_overview.png", dpi=200); plt.close(fig)

# 2. Walking route for field-show bands, elementary to high school
fig, ax = canvas("Figure 6. Walking route for field-show bands: community field to elementary to high school (approximate)")
zone(ax, [(1200, 330), (1500, 330), (1500, 560), (1200, 560)], BLUE, "Warm-up and staging\n(community field)", 1350, 300, alpha=0.15)
zone(ax, [(1000, 690), (1300, 690), (1300, 880), (1000, 880)], BLUE, "Clinic (elementary)", 1150, 905)
zone(ax, [(200, 530), (370, 530), (370, 690), (200, 690)], TEAL, "Show slot", 285, 610, alpha=0.15)
zone(ax, [(300, 190), (480, 190), (480, 430), (300, 430)], TEAL, "High school", 390, 160, alpha=0.15)
route(ax, [(1190, 600), (1060, 610), (960, 600), (850, 540), (700, 470), (560, 420), (460, 410)], GOLD,
      "Walk (distance to be measured), flagged crossings", 820, 390)
note(fig, "Walk along W. Sandy Ave. and S. Raider Ave. (approximate). Flaggers at each crossing.")
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig("figures/map_walk.png", dpi=200); plt.close(fig)

# 3. Parking, trailer staging and gate off Clemenceau
fig, ax = canvas("Figure 7. Parking and trailer plan: visitors use the elementary lot, trailers stage off Clemenceau (approximate)")
zone(ax, [(640, 600), (820, 600), (820, 790), (640, 790)], RED, "Fair parking\n(open grass lot)", 730, 690, alpha=0.2)
zone(ax, [(1000, 630), (1150, 630), (1150, 700), (1000, 700)], BLUE, "Elementary lot:\nvisitors and shuttle", 1230, 790, alpha=0.2)
zone(ax, [(215, 600), (520, 600), (520, 860), (215, 860)], GOLD, "Candidate trailer\nstaging (grass)", 370, 730, alpha=0.2)
ax.text(560, 845, "Gate", fontsize=8.5, color="white", ha="center", va="center", weight="bold", bbox=dict(boxstyle="round,pad=0.3", fc=RED, ec="none"))
ax.annotate("", xy=(505, 900), xytext=(505, 862), arrowprops=dict(arrowstyle="-|>", color=RED, lw=2.5))
route(ax, [(560, 905), (560, 870), (520, 800), (470, 760)], GOLD, "Trailers in from Clemenceau", 700, 905, ls="--")
note(fig, "Gate location, grass ownership and fire station position to confirm. Fire lane must stay clear.")
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig("figures/map_parking.png", dpi=200); plt.close(fig)

# 4. Parade segment (approximate) on the labeled aerial
fig, ax = canvas("Figure 8. Parade segment, 10 a.m. on W. Tahoe Ave. from S. March Ave. to S. Raider Ave. (approximate, VERIFY)")
route(ax, [(780, 230), (700, 300), (620, 370), (520, 420)], RED, "Parade: W. Tahoe Ave., March to Raider", 650, 200, ls="--")
zone(ax, [(640, 600), (820, 600), (820, 790), (640, 790)], RED, "Fair parking", 730, 690, alpha=0.1)
zone(ax, [(300, 190), (480, 190), (480, 430), (300, 430)], TEAL, "High school", 390, 160, alpha=0.15)
note(fig, "Parade segment is approximate; the fair board must confirm the route before any band packet is issued.")
fig.tight_layout(rect=(0, 0.04, 1, 1)); fig.savefig("figures/map_parade.png", dpi=200); plt.close(fig)
print("maps saved")
