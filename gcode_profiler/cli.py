import argparse
import sys
import time
from pathlib import Path
import hashlib
import json
import re

from .config_ini import parse_config_ini, config_get_float
from .constants import DEFAULT_BINS, DEFAULT_FILAMENT_DIAMETER_MM, DEFAULT_TOP_N_SLOWEST
from .gcode_parser import parse_gcode
from .excel_writer import write_xlsx, write_csv_exports, build_json_summary

_T0 = time.time()


def status(msg: str, enabled: bool = True):
    """Emit a lightweight progress message to stderr."""
    if not enabled:
        return
    dt = time.time() - _T0
    print(f"[{dt:6.1f}s] {msg}", file=sys.stderr, flush=True)


def make_profile_label(cfg_info: dict | None, fallback: str) -> str:
    """Create a descriptive run label from config.ini fields.

    Format: <HF/nozzle>_<layer>_<TAG>

    Uses (when present): filament_settings_id, print_settings_id, nozzle_diameter, layer_height.
    """
    if not cfg_info:
        return fallback

    try:
        nozzle = cfg_info.get("nozzle_diameter")
        layer = cfg_info.get("layer_height")
        filament_id = (cfg_info.get("filament_settings_id") or "").strip(' "')
        print_id = (cfg_info.get("print_settings_id") or "").strip(' "')

        # Prefer an HF* token (e.g., "... HF0.6") from filament_settings_id
        hf = None
        m = re.search(r"(HF\s*\d+(?:\.\d+)?)", filament_id, re.IGNORECASE)
        if m:
            hf = m.group(1).replace(" ", "")
        if hf is None and isinstance(nozzle, (int, float)):
            hf = f"HF{nozzle:.1f}".rstrip("0").rstrip(".")

        layer_s = None
        if isinstance(layer, (int, float)):
            layer_s = f"{layer:.2f}"

        # Pick a tag from print_settings_id (first alphabetic token of length>=3 that's not a common noise word)
        tag = None
        if print_id:
            toks = re.split(r"[^A-Za-z]+", print_id.upper())
            toks = [t for t in toks if t and t not in ("MM", "COREONE", "PRUSA", "AT")]
            for t in toks:
                if t.isalpha() and len(t) >= 3:
                    tag = t
                    break

        parts = [p for p in (hf, layer_s, tag) if p]
        return "_".join(parts) if parts else fallback
    except Exception:
        return fallback


def _load_config_info(path: Path) -> dict:
    cfg = parse_config_ini(str(path))
    return {
        "nozzle_diameter": config_get_float(cfg, "nozzle_diameter"),
        "filament_diameter": config_get_float(cfg, "filament_diameter"),
        "filament_density": config_get_float(cfg, "filament_density"),
        "filament_max_volumetric_speed": config_get_float(cfg, "filament_max_volumetric_speed"),
        "max_print_speed": config_get_float(cfg, "max_print_speed"),
        "layer_height": config_get_float(cfg, "layer_height"),
        "first_layer_height": config_get_float(cfg, "first_layer_height"),
        "max_layer_height": config_get_float(cfg, "max_layer_height"),
        "min_layer_height": config_get_float(cfg, "min_layer_height"),
        "max_fan_speed": config_get_float(cfg, "max_fan_speed"),
        "print_settings_id": cfg.get("print_settings_id"),
        "filament_settings_id": cfg.get("filament_settings_id"),
        "printer_settings_id": cfg.get("printer_settings_id"),
        "_raw": cfg,
    }


def main():
    parser = argparse.ArgumentParser(
        description="GCodeProfiler: analyze PrusaSlicer ASCII .gcode and export an Excel performance dashboard (metrics, legends, charts)."
    )
    parser.add_argument("gcode", help="Input ASCII .gcode file exported from PrusaSlicer")
    parser.add_argument(
        "--compare",
        nargs="*",
        help="Optional one or more additional .gcode files to compare against. Produces Compare_* sheets and comparison summary/charts.",
    )
    parser.add_argument("--config", help="Optional PrusaSlicer config.ini (key=value) to improve chart scaling and defaults.")
    parser.add_argument(
        "--compare-config",
        nargs="*",
        help=(
            "Optional config.ini file(s) corresponding to --compare runs. "
            "Provide either ONE file (applies to first compare run) or the same count as --compare."
        ),
    )
    parser.add_argument("--output", help="Optional output .xlsx path (defaults to <input>.xlsx or <input>_vs_<compare>.xlsx)")
    parser.add_argument("--bins", type=int, default=DEFAULT_BINS, help=f"Histogram bins (default: {DEFAULT_BINS})")
    parser.add_argument("--no-legends", action="store_true", help="Skip legend-like Feature Type summary table")
    parser.add_argument("--per-layer-only", action="store_true", help="Skip per-move rows for smaller files")
    parser.add_argument(
        "--top-n-slowest",
        type=int,
        default=DEFAULT_TOP_N_SLOWEST,
        help=f"How many slowest layers to list/chart (default: {DEFAULT_TOP_N_SLOWEST})",
    )
    parser.add_argument(
        "--top-n-segments",
        type=int,
        default=200,
        help="How many top extrusion segments by volumetric flow to list (default: 200)",
    )
    parser.add_argument(
        "--min-peak-segment-time",
        type=float,
        default=0.05,
        help="Ignore extrusion segments shorter than this (seconds) when computing peak flow/speed (reduces spike noise). Default: 0.05",
    )
    parser.add_argument(
        "--layout",
        choices=["compact", "wide"],
        default="compact",
        help="Dashboard layout. compact packs charts tighter; wide uses larger charts/spacing.",
    )
    parser.add_argument("--quiet", action="store_true", help="Suppress progress output")
    parser.add_argument("--filament-diameter", type=float, help="Override filament diameter (mm)")
    parser.add_argument("--filament-density", type=float, help="Override filament density (g/cm^3)")
    parser.add_argument("--csv", action="store_true", help="Also write CSV exports next to the .xlsx")
    parser.add_argument("--json", action="store_true", help="Also write a small JSON summary next to the .xlsx")

    args = parser.parse_args()
    status_enabled = not bool(args.quiet)

    gcode_path = Path(args.gcode)
    if not gcode_path.exists():
        raise FileNotFoundError(f"G-code file not found: {gcode_path}")

    compare_paths = [Path(p) for p in (args.compare or [])]
    for p in compare_paths:
        if not p.exists():
            raise FileNotFoundError(f"Compare G-code file not found: {p}")

    if args.output:
        out_xlsx = Path(args.output)
    else:
        if not compare_paths:
            out_xlsx = gcode_path.with_suffix(".xlsx")
        else:
            base = gcode_path.with_suffix("")
            suffix = "_vs_" + "_".join([p.stem for p in compare_paths])
            out_xlsx = base.with_name(f"{base.name}{suffix}.xlsx")

    config_info = None
    if args.config:
        cfg_path = Path(args.config)
        if not cfg_path.exists():
            raise FileNotFoundError(f"Config file not found: {cfg_path}")
        config_info = _load_config_info(cfg_path)

    a_label = make_profile_label(config_info, "A")

    # Compare config infos (optional)
    compare_config_infos: list[dict | None] = []
    compare_cfg_paths = [Path(p) for p in (args.compare_config or [])]
    for p in compare_cfg_paths:
        if not p.exists():
            raise FileNotFoundError(f"Compare config file not found: {p}")

    if compare_cfg_paths:
        if len(compare_cfg_paths) == 1 and len(compare_paths) >= 1:
            compare_config_infos = [_load_config_info(compare_cfg_paths[0])] + [None] * (len(compare_paths) - 1)
        elif len(compare_cfg_paths) == len(compare_paths):
            compare_config_infos = [_load_config_info(p) for p in compare_cfg_paths]
        else:
            raise ValueError("--compare-config must provide either 1 file or the same count as --compare")

    filament_diam = (
        float(args.filament_diameter)
        if args.filament_diameter is not None
        else float((config_info or {}).get("filament_diameter") or DEFAULT_FILAMENT_DIAMETER_MM)
    )
    filament_density = (
        float(args.filament_density)
        if args.filament_density is not None
        else float((config_info or {}).get("filament_density") or 1.24)
    )

    status(f"Parsing G-code A ({gcode_path.name})", status_enabled)
    moves, layer_z_map = parse_gcode(
        str(gcode_path),
        filament_diam,
        status_cb=(lambda m: status(m, status_enabled)),
        status_every_lines=250_000,
    )

    compare_runs = []
    for idx, cp in enumerate(compare_paths, start=1):
        status(f"Parsing compare G-code {idx} ({cp.name})", status_enabled)
        cm, cz = parse_gcode(
            str(cp),
            filament_diam,
            status_cb=(lambda m: status(m, status_enabled)),
            status_every_lines=250_000,
        )
        cfg_i = compare_config_infos[idx - 1] if (compare_config_infos and (idx - 1) < len(compare_config_infos)) else None
        compare_runs.append(
            {
                "path": str(cp),
                "label": make_profile_label(cfg_i, cp.stem),
                "moves": cm,
                "layer_z_map": cz,
                "config_info": cfg_i,
            }
        )

    status("Building Excel workbook", status_enabled)

    write_xlsx(
        moves,
        layer_z_map,
        str(out_xlsx),
        bins=int(args.bins),
        include_legends=not args.no_legends,
        per_layer_only=bool(args.per_layer_only),
        top_n_slowest=int(args.top_n_slowest),
        filament_diameter_mm=filament_diam,
        filament_density_g_cm3=filament_density,
        config_info=config_info,
        layout=str(args.layout),
        run_label=a_label,
        min_peak_segment_time_s=float(args.min_peak_segment_time),
        compare_runs=compare_runs,
        top_n_segments=int(args.top_n_segments),
        status_cb=(lambda m: status(m, status_enabled)),
    )

    # Optional sidecar exports
    if args.csv:
        write_csv_exports(moves, layer_z_map, str(out_xlsx), config_info=config_info, top_n_segments=int(args.top_n_segments))
    if args.json:
        summary = build_json_summary(moves, layer_z_map, config_info=config_info)
        with open(Path(str(out_xlsx)).with_suffix(".summary.json"), "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2)

        # A lightweight run metadata file (helps compare/trace)
        meta = {
            "input_gcode": str(gcode_path),
            "compare_inputs": [str(p) for p in compare_paths],
            "output_xlsx": str(out_xlsx),
            "label_a": a_label,
            "labels_compare": [r["label"] for r in compare_runs],
            "filament_diameter_mm": filament_diam,
            "filament_density_g_cm3": filament_density,
        }
        meta_bytes = json.dumps(meta, sort_keys=True).encode("utf-8")
        meta["run_hash"] = hashlib.sha256(meta_bytes).hexdigest()
        with open(Path(str(out_xlsx)).with_suffix(".run.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)

    status(f"Done -> {out_xlsx}", status_enabled)
    print(f"Wrote {len(moves)} moves to {out_xlsx}")


if __name__ == "__main__":
    main()
