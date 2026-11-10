import sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle, Polygon, Rectangle

ROT, GRAU, DUNKEL = "#DC372D", "#C8C8C8", "#222222"
out = sys.argv[1]

# 1) Geordneter Rückbau
fig, ax = plt.subplots(figsize=(11.4, 4.75), dpi=200)
stufen = [("Normal Law", "volle Schutzfunktionen"),
          ("Alternate Law", "weniger Schutz"),
          ("Direct Law", "Knüppel direkt auf die Ruder"),
          ("Mechanische Rückfallebene", "Trimmung und Seitenruder")]
farben = [ROT, "#E56962", "#EE9B96", GRAU]
w, h = 2.6, 0.9
for i, ((t, s), f) in enumerate(zip(stufen, farben)):
    x, y = i * 2.85, 3 - i * 0.95
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                fc=f, ec="none"))
    ax.text(x + w / 2, y + 0.58, t, ha="center", va="center", fontsize=15, weight="bold",
            color="white" if i < 3 else DUNKEL)
    ax.text(x + w / 2, y + 0.25, s, ha="center", va="center", fontsize=11.5,
            color="white" if i < 3 else DUNKEL)
    if i < 3:
        ax.annotate("", xy=(x + w + 0.22, y - 0.05), xytext=(x + w - 0.05, y + 0.45),
                    arrowprops=dict(arrowstyle="-|>", color=DUNKEL, lw=1.6))
ax.text(0, -0.15, "Jede Stufe ist geplant, beschrieben und trainiert.", fontsize=13, color=DUNKEL, style="italic")
ax.set_xlim(-0.1, 11.3); ax.set_ylim(-0.4, 4.0); ax.axis("off")
fig.savefig(f"{out}/rueckbau_a320.png", bbox_inches="tight", transparent=False, facecolor="white")
plt.close(fig)

# 2) Formcodierung: Rad vs. Keil (schematisch, eigene Zeichnung)
fig, ax = plt.subplots(figsize=(11.4, 4.75), dpi=200)
for cx, label, sub in [(2.6, "Fahrwerk", "Knauf in Radform"), (8.4, "Landeklappen", "Knauf in Tragflächenform")]:
    ax.add_patch(Rectangle((cx - 0.12, 0.5), 0.24, 2.0, fc="#555555", ec="none"))
    ax.add_patch(Rectangle((cx - 1.4, 0.3), 2.8, 0.25, fc=GRAU, ec="none"))
    ax.text(cx, -0.25, label, ha="center", fontsize=18, weight="bold", color=DUNKEL)
    ax.text(cx, -0.75, sub, ha="center", fontsize=13, color=DUNKEL)
# Rad
ax.add_patch(Circle((2.6, 3.1), 0.75, fc="#333333", ec="none"))
ax.add_patch(Circle((2.6, 3.1), 0.42, fc="white", ec="none"))
ax.add_patch(Circle((2.6, 3.1), 0.12, fc="#333333", ec="none"))
# Tragflächenprofil
xs = [6.9, 7.3, 8.0, 9.0, 9.9, 9.0, 8.0, 7.3]
ys = [2.7, 3.05, 3.2, 3.1, 2.75, 2.55, 2.5, 2.52]
ax.add_patch(Polygon(list(zip(xs, ys)), closed=True, fc=ROT, ec="none"))
ax.text(5.5, 1.6, "≠", ha="center", va="center", fontsize=60, color=ROT)
ax.text(5.5, 4.35, "Schematisch nach Chapanis – heute 14 CFR § 25.781", ha="center", fontsize=12, color="#666666", style="italic")
ax.set_xlim(0, 11); ax.set_ylim(-1.0, 4.6); ax.set_aspect("equal"); ax.axis("off")
fig.savefig(f"{out}/hebel_formcodierung.png", bbox_inches="tight", facecolor="white")
plt.close(fig)
print("ok")
