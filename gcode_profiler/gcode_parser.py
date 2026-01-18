import math
import re

def filament_area_mm2(d_mm: float) -> float:
    r = d_mm / 2.0
    return math.pi * r * r


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


def parse_gcode(
    gcode_path: str,
    filament_diameter_mm: float,
    status_cb=None,
    status_every_lines: int = 0,
):
    area = filament_area_mm2(filament_diameter_mm)

    # Position state
    x = y = z = 0.0
    e = 0.0
    feed_mm_min = None
    e_relative = True  # honor M82/M83

    current_layer = 0
    current_type = "UNKNOWN"

    saw_layer_tag = False
    last_layer_z_comment = None
    # Temps / fan setpoints (latest known)
    hotend_set = None
    bed_set = None
    chamber_set = None
    fan_s_0_255 = None

    # Layer Z mapping from slicer comments
    layer_z_map = {}

    moves = []

    re_type = re.compile(r";\s*TYPE:(.+)\s*$")
    re_z = re.compile(r";\s*Z:([0-9.+-]+)")
    re_layer = re.compile(r";\s*LAYER:\s*([0-9]+)")

    re_g0g1 = re.compile(r"^(G0|G1)\s+(.*)$")
    re_param = re.compile(r"([XYZEFS])([0-9.+-]+)")

    with open(gcode_path, "r", encoding="utf-8", errors="replace") as f:
        for i, line in enumerate(f, start=1):
            if status_cb is not None and status_every_lines and (i % status_every_lines == 0):
                status_cb(f"Parsed {i:,} lines")
            line = line.rstrip("\n")

            # Feature type
            m = re_type.search(line)
            if m:
                current_type = m.group(1).strip()
                continue

            # Layer markers
            m = re_layer.search(line)
            if m:
                saw_layer_tag = True
                current_layer = int(m.group(1))
                continue


            m = re_z.search(line)
            if m:
                try:
                    zc = float(m.group(1))
                except ValueError:
                    continue

                # If slicer didn't emit ;LAYER:n, infer layer index from increasing ;Z values.
                if not saw_layer_tag:
                    if last_layer_z_comment is None:
                        current_layer = 0
                    elif zc > last_layer_z_comment + 1e-6:
                        current_layer += 1
                    last_layer_z_comment = zc

                layer_z_map[current_layer] = zc
                continue

            # Extrusion mode
            if line.startswith("M82"):
                e_relative = False
                continue
            if line.startswith("M83"):
                e_relative = True
                continue

            # Fan
            if line.startswith("M106"):
                ms = re.search(r"\bS(\d+)", line)
                if ms:
                    fan_s_0_255 = int(ms.group(1))
                continue
            if line.startswith("M107"):
                fan_s_0_255 = 0
                continue

            # Temperatures
            if line.startswith(("M104", "M109")):
                ms = re.search(r"\bS([0-9.+-]+)", line)
                if ms:
                    try:
                        hotend_set = float(ms.group(1))
                    except ValueError:
                        pass
                continue
            if line.startswith(("M140", "M190")):
                ms = re.search(r"\bS([0-9.+-]+)", line)
                if ms:
                    try:
                        bed_set = float(ms.group(1))
                    except ValueError:
                        pass
                continue
            if line.startswith("M141"):
                ms = re.search(r"\bS([0-9.+-]+)", line)
                if ms:
                    try:
                        chamber_set = float(ms.group(1))
                    except ValueError:
                        pass
                continue

            # Moves
            mg = re_g0g1.match(line)
            if not mg:
                continue

            cmd = mg.group(1)
            rest = mg.group(2)
            params = {k: float(v) for (k, v) in re_param.findall(rest)}

            nx = params.get("X", x)
            ny = params.get("Y", y)
            nz = params.get("Z", z)

            if "F" in params:
                feed_mm_min = params["F"]
            feed_mm_s = (feed_mm_min / 60.0) if (feed_mm_min and feed_mm_min > 0) else None

            if "E" in params:
                e_cmd = params["E"]
                if e_relative:
                    de = e_cmd
                    ne = e + de
                else:
                    de = e_cmd - e
                    ne = e_cmd
            else:
                de = 0.0
                ne = e

            dx = nx - x
            dy = ny - y
            dz = nz - z
            dist = math.sqrt(dx * dx + dy * dy + dz * dz)

            t_s = (dist / feed_mm_s) if (feed_mm_s and dist > 0) else 0.0
            speed = feed_mm_s if feed_mm_s else None

            if t_s > 0 and de > 0:
                vol_mm3 = de * area
                flow = vol_mm3 / t_s
            else:
                flow = 0.0

            fan_pct = (fan_s_0_255 / 255.0 * 100.0) if fan_s_0_255 is not None else None

            moves.append(
                {
                    "layer": current_layer,
                    "z": nz,
                    "type": current_type,
                    "cmd": cmd,
                    "x0": x,
                    "y0": y,
                    "z0": z,
                    "x1": nx,
                    "y1": ny,
                    "z1": nz,
                    "dist_mm": dist,
                    "de_mm": de,
                    "speed_mm_s": speed,
                    "time_s": t_s,
                    "flow_mm3_s": flow,
                    "fan_pct": fan_pct,
                    "hotend_C": hotend_set,
                    "bed_C": bed_set,
                    "chamber_C": chamber_set,
                }
            )

            x, y, z, e = nx, ny, nz, ne

    return moves, layer_z_map


