import argparse
import sys
import time
from pathlib import Path
import hashlib
import json

from .config_ini import parse_config_ini, config_get_float
from .constants import DEFAULT_BINS, DEFAULT_FILAMENT_DIAMETER_MM, DEFAULT_TOP_N_SLOWEST
from .gcode_parser import parse_gcode
from .excel_writer import write_xlsx
from .excel_writer import write_csv_exports, build_json_summary


_T0 = time.time()


def status(msg: str, enabled: bool = True):
    """Emit a lightweight progress message to stderr."""
    if not enabled:
        return
    dt = time.time() - _T0
    print(f"[{dt:6.1f}s] {msg}", file=sys.stderr, flush=True)

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
    parser.add_argument("--output", help="Override output path (default: input basename + .xlsx)")
    parser.add_argument("--bins", type=int, default=DEFAULT_BINS, help=f"Number of legend bins (default: {DEFAULT_BINS})")
    parser.add_argument("--filament-diameter", type=float, default=None, help=f"Filament diameter in mm (default: {DEFAULT_FILAMENT_DIAMETER_MM}; can be overridden by --config)")
    parser.add_argument("--filament-density", type=float, default=None, help="Filament density in g/cm^3 for Used filament (g) estimates (default: 1.24; can be overridden by --config)")
    parser.add_argument("--no-legends", action="store_true", help="Skip Legend_* sheets (charts will be limited)")
    parser.add_argument("--per-layer-only", action="store_true", help="Skip per-move rows for smaller files")
    parser.add_argument("--top-n-slowest", type=int, default=DEFAULT_TOP_N_SLOWEST, help=f"How many slowest layers to list/chart (default: {DEFAULT_TOP_N_SLOWEST})")
    parser.add_argument("--top-n-segments", type=int, default=200, help="How many top extrusion segments by volumetric flow to list (default: 200)")
    parser.add_argument("--layout", choices=["compact", "wide"], default="compact", help="Dashboard layout. compact packs charts tighter; wide uses larger charts/spacing.")
    parser.add_argument("--csv", action="store_true", help="Also export key tables as CSV next to the .xlsx")
    parser.add_argument("--json-summary", action="store_true", help="Also export a JSON summary (next to the .xlsx)")
    parser.add_argument("--no-manifest", action="store_true", help="Do not write a run manifest JSON (default: write one)")
    parser.add_argument("--quiet", action="store_true", help="Suppress progress output")

    args = parser.parse_args()
    status_enabled = not bool(args.quiet)

    gcode_path = Path(args.gcode)
    if not gcode_path.exists():
        raise FileNotFoundError(f"G-code file not found: {gcode_path}")

    if gcode_path.suffix.lower() != '.gcode':
        print('Warning: input file does not have a .gcode extension. This tool expects ASCII G-code.')

    compare_paths = [Path(p) for p in (args.compare or [])]
    for p in compare_paths:
        if not p.exists():
            raise FileNotFoundError(f"Compare G-code file not found: {p}")

    if args.output:
        out_xlsx = Path(args.output)
    else:
        if not compare_paths:
            out_xlsx = gcode_path.with_suffix('.xlsx')
        else:
            out_xlsx = gcode_path.with_suffix('')
            suffix = "_vs_" + "_".join([p.stem for p in compare_paths])
            out_xlsx = out_xlsx.with_name(f"{out_xlsx.name}{suffix}.xlsx")

    config_info = None
    if args.config:
        cfg_path = Path(args.config)
        if not cfg_path.exists():
            raise FileNotFoundError(f"Config file not found: {cfg_path}")
        cfg = parse_config_ini(str(cfg_path))
        config_info = {
            'nozzle_diameter': config_get_float(cfg, 'nozzle_diameter'),
            'filament_diameter': config_get_float(cfg, 'filament_diameter'),
            'filament_density': config_get_float(cfg, 'filament_density'),
            'filament_max_volumetric_speed': config_get_float(cfg, 'filament_max_volumetric_speed'),
            'max_print_speed': config_get_float(cfg, 'max_print_speed'),
            'layer_height': config_get_float(cfg, 'layer_height'),
            'first_layer_height': config_get_float(cfg, 'first_layer_height'),
            'max_layer_height': config_get_float(cfg, 'max_layer_height'),
            'min_layer_height': config_get_float(cfg, 'min_layer_height'),
            'max_fan_speed': config_get_float(cfg, 'max_fan_speed'),
        }

    filament_diam = float(args.filament_diameter) if args.filament_diameter is not None else float((config_info or {}).get('filament_diameter') or DEFAULT_FILAMENT_DIAMETER_MM)
    filament_density = float(args.filament_density) if args.filament_density is not None else float((config_info or {}).get('filament_density') or 1.24)

    status(f"Parsing G-code A ({gcode_path.name})", status_enabled)
    moves, layer_z_map = parse_gcode(str(gcode_path), filament_diam, status_cb=(lambda m: status(m, status_enabled)), status_every_lines=250_000)
    compare_runs = []
    for idx, cp in enumerate(compare_paths, start=1):
        status(f"Parsing compare G-code {idx} ({cp.name})", status_enabled)
        cm, cz = parse_gcode(str(cp), filament_diam, status_cb=(lambda m: status(m, status_enabled)), status_every_lines=250_000)
        compare_runs.append({"path": str(cp), "label": cp.stem, "moves": cm, "layer_z_map": cz})

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
        compare_runs=compare_runs,
        status_cb=(lambda m: status(m, status_enabled)),
        top_n_segments=int(args.top_n_segments),
    )

    # Optional sidecar exports
    if args.csv:
        status("Writing CSV exports", status_enabled)
        write_csv_exports(moves, layer_z_map, str(out_xlsx), config_info=config_info, top_n_segments=int(args.top_n_segments))
    if args.json_summary:
        status("Writing JSON summary", status_enabled)
        summary = build_json_summary(moves, layer_z_map, config_info=config_info)
        with open(Path(str(out_xlsx)).with_suffix('.summary.json'), 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2)

    # Run manifest
    if not args.no_manifest:
        status("Writing run manifest", status_enabled)
        def sha256_file(path: Path) -> str:
            h = hashlib.sha256()
            with open(path, 'rb') as fp:
                for chunk in iter(lambda: fp.read(1024 * 1024), b''):
                    h.update(chunk)
            return h.hexdigest()

        manifest = {
            "tool": "GCodeProfiler",
            "output_xlsx": str(out_xlsx),
            "inputs": {
                "gcode": {"path": str(gcode_path), "sha256": sha256_file(gcode_path)},
                "compare": [{"path": p, "sha256": sha256_file(Path(p))} for p in [r["path"] for r in compare_runs]],
                "config": None,
            },
            "args": vars(args),
        }
        if args.config:
            cfg_path = Path(args.config)
            manifest["inputs"]["config"] = {"path": str(cfg_path), "sha256": sha256_file(cfg_path)}
        with open(Path(str(out_xlsx)).with_suffix('.run.json'), 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2)

    status(f"Done -> {out_xlsx}", status_enabled)
    print(f"Wrote {len(moves)} moves to {out_xlsx}")

