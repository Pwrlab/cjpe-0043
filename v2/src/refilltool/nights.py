"""Split a sap flow time series into nights using solar radiation."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass
class Night:
    night_id: int
    start: int          # first row index (inclusive)
    end: int            # last row index (exclusive)
    start_time: pd.Timestamp
    end_time: pd.Timestamp
    partial: bool = False
    gap: bool = False

    @property
    def n(self):
        return self.end - self.start


def split_nights_legacy(rad) -> list[tuple[int, int]]:
    """Exact port of v1 ``extract_data`` (slices are [start, end), end excluded).

    Keeps v1 quirks: the first segment starts at row 0 even when data begin at
    night, the last night point before sunrise is dropped, and the start test
    uses ``< 5`` while the night test uses ``<= 5``.
    """
    rad = np.asarray(rad, dtype=float)
    n = len(rad)
    out, start, i = [], 0, 0
    while i < n:
        if rad[i] <= 5:
            if i + 1 == n:
                out.append((start, i))
                break
            if rad[i + 1] > 5:
                out.append((start, i))
        if rad[i] > 5:
            if i + 1 == n:
                break
            if rad[i + 1] < 5:
                start = i + 1
        i += 1
    return out


def _runs(mask):
    """Return (value, start, end_exclusive) for runs of a boolean array."""
    mask = np.asarray(mask, dtype=bool)
    if len(mask) == 0:
        return []
    change = np.flatnonzero(np.diff(mask.astype(int))) + 1
    starts = np.r_[0, change]
    ends = np.r_[change, len(mask)]
    return [(bool(mask[s]), int(s), int(e)) for s, e in zip(starts, ends)]


def night_mask(rad, threshold=5.0, min_run_points=0):
    """Night = radiation <= threshold. Runs shorter than ``min_run_points``
    (radiation spikes at night or dark spells in daytime) are absorbed into the
    surrounding period."""
    # Missing radiation takes the state of the previous record, so a gap does not split a night
    rad = pd.Series(np.asarray(rad, dtype=float)).ffill().bfill().to_numpy()
    mask = np.where(np.isnan(rad), False, rad <= threshold)
    if min_run_points > 1:
        for _ in range(3):  # a few passes are enough to merge nested blips
            runs = _runs(mask)
            changed = False
            for k, (val, s, e) in enumerate(runs):
                interior = 0 < k < len(runs) - 1
                if interior and (e - s) < min_run_points:
                    mask[s:e] = not val
                    changed = True
            if not changed:
                break
    return mask


def split_nights(time, rad, threshold=5.0, min_run_points=0, step=None) -> list[Night]:
    """v2 segmentation: full night runs, flags partial nights at both ends of
    the record and nights containing time gaps."""
    time = pd.to_datetime(pd.Series(time)).reset_index(drop=True)
    rad_nan = np.isnan(np.asarray(rad, dtype=float))
    mask = night_mask(rad, threshold, min_run_points)
    nights = []
    n = len(mask)
    for val, s, e in _runs(mask):
        if not val:
            continue
        partial = s == 0 or e == n
        gap = bool(rad_nan[s:e].any())  # inserted (missing) records
        if step is not None and e - s > 1:
            d = time.iloc[s:e].diff().dropna()
            gap = gap or bool((d > step * 1.5).any())
        nights.append(Night(len(nights) + 1, s, e, time.iloc[s], time.iloc[e - 1], partial, gap))
    return nights


def legacy_nights(time, rad) -> list[Night]:
    time = pd.to_datetime(pd.Series(time)).reset_index(drop=True)
    out = []
    for k, (s, e) in enumerate(split_nights_legacy(rad)):
        st = time.iloc[s]
        et = time.iloc[max(e - 1, s)]
        out.append(Night(k + 1, s, e, st, et, partial=(s == 0)))
    return out
