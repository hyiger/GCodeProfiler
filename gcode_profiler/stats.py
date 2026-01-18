import math

def weighted_quantile(values, weights, q: float):
    """Compute a weighted quantile.

    - Uses weights as relative importance (typically time_s per move).
    - q in [0,1].
    """
    if not values:
        return None
    if weights is None:
        weights = [1.0] * len(values)
    if len(values) != len(weights):
        raise ValueError("values and weights must be same length")
    q = max(0.0, min(1.0, float(q)))

    pairs = [(float(v), float(w)) for v, w in zip(values, weights) if v is not None and w is not None and w > 0]
    if not pairs:
        return None
    pairs.sort(key=lambda x: x[0])
    total_w = sum(w for _, w in pairs)
    if total_w <= 0:
        return None
    cutoff = q * total_w
    acc = 0.0
    for v, w in pairs:
        acc += w
        if acc >= cutoff:
            return v
    return pairs[-1][0]


def make_bins(min_v: float, max_v: float, bins: int):
    if bins < 1:
        bins = 1
    if max_v < min_v:
        min_v, max_v = max_v, min_v
    if math.isclose(max_v, min_v):
        return [(min_v, max_v)]
    step = (max_v - min_v) / bins
    out = []
    lo = min_v
    for i in range(bins):
        hi = (min_v + (i + 1) * step) if i < bins - 1 else max_v
        out.append((lo, hi))
        lo = hi
    return out


def bin_counts(values, bins_spec):
    """Return counts per bin. bins_spec is list[(lo, hi)], inclusive lo, exclusive hi except last."""
    counts = [0] * len(bins_spec)
    if not bins_spec:
        return counts
    for v in values:
        if v is None:
            continue
        placed = False
        for i, (lo, hi) in enumerate(bins_spec):
            if i < len(bins_spec) - 1:
                if lo <= v < hi:
                    counts[i] += 1
                    placed = True
                    break
            else:
                if lo <= v <= hi:
                    counts[i] += 1
                    placed = True
                    break
        if not placed:
            if v < bins_spec[0][0]:
                counts[0] += 1
            else:
                counts[-1] += 1
    return counts

