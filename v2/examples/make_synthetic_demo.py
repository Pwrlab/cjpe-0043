"""Generate a SYNTHETIC demo file (not field data) with two sensor heights,
rain and VPD, to show the input format of the height-difference method.

Run:  python examples/make_synthetic_demo.py
"""
from pathlib import Path

import numpy as np
import pandas as pd

rng = np.random.default_rng(42)
t = pd.date_range("2021-07-01 00:00", "2021-07-21 23:50", freq="10min")
h = np.asarray(t.hour + t.minute / 60)
rad = np.asarray(np.clip(950 * np.sin((h - 5.5) / 13 * np.pi), 0, None))
rad[(t.day == 8) & (h > 9) & (h < 16)] *= 0.25            # one overcast day
rain = np.zeros(len(t))
rain[(t.day == 8) & (t.hour == 15)] = 1.2                  # rain on day 8
vpd = np.clip(0.4 + 1.6 * np.sin((h - 8) / 24 * 2 * np.pi), 0.05, None) + rng.normal(0, 0.03, len(t))

df = pd.DataFrame({"time": t, "solar_radiation": rad.round(2), "rain_mm": rain, "VPD_kPa": vpd.round(3)})
night = rad <= 5
k = np.zeros(len(t))
for i in range(1, len(t)):
    k[i] = k[i - 1] + 1 if night[i] else 0
for j, (tr, stem, canopy) in enumerate([(0.004, 0.012, 0.010), (0.003, 0.010, 0.008), (0.005, 0.015, 0.012)], 1):
    day = 0.06 * np.roll(rad, 3) / 950                     # ~30 min lag behind radiation
    upper = np.where(night, canopy * np.exp(-0.035 * k) + tr, day)
    lower = np.where(night, upper + stem * np.exp(-0.02 * k), day * 1.05)
    df[f"T{j}_1.3m"] = (lower + rng.normal(0, 0.0004, len(t))).clip(0).round(5)
    df[f"T{j}_7.0m"] = (upper + rng.normal(0, 0.0004, len(t))).clip(0).round(5)

out = Path(__file__).with_name("synthetic_demo_two_heights.csv")
df.to_csv(out, index=False)
print(f"written {out} ({len(df)} rows)")
