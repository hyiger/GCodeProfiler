import argparse
import sys
import time
from pathlib import Path

from .config_ini import parse_config_ini, config_get_float
from .constants import DEFAULT_BINS, DEFAULT_FILAMENT_DIAMETER_MM, DEFAULT_TOP_N_SLOWEST
from .gcode_parser import parse_gcode
from .excel_writer import write_xlsx


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
    parser.add_argument("--compare", help="Optional second .gcode to compare against. Produces extra Compare_* sheets and overlay charts.")
    parser.add_argument("--config", help="Optional PrusaSlicer config.ini (key=value) to improve chart scaling and defaults.")
    parser.add_argument("--output", help="Override output path (default: input basename + .xlsx)")
    parser.add_argument("--bins", type=int, default=DEFAULT_BINS, help=f"Number of legend bins (default: {DEFAULT_BINS})")
    parser.add_argument("--filament-diameter", type=float, default=None, help=f"Filament diameter in mm (default: {DEFAULT_FILAMENT_DIAMETER_MM}; can be overridden by --config)")
    parser.add_argument("--filament-density", type=float, default=None, help="Filament density in g/cm^3 for Used filament (g) estimates (default: 1.24; can be overridden by --config)")
    parser.add_argument("--no-legends", action="store_true", help="Skip Legend_* sheets (charts will be limited)")
    parser.add_argument("--per-layer-only", action="store_true", help="Skip per-move rows for smaller files")
    parser.add_argument("--top-n-slowest", type=int, default=DEFAULT_TOP_N_SLOWEST, help=f"How many slowest layers to list/chart (default: {DEFAULT_TOP_N_SLOWEST})")
    parser.add_argument("--layout", choices=["compact", "wide"], default="compact", help="Dashboard layout. compact packs charts tighter; wide uses larger charts/spacing.")
    parser.add_argument("--quiet", action="store_true", help="Suppress progress output")

    args = parser.parse_args()
    status_enabled = not bool(args.quiet)

    gcode_path = Path(args.gcode)
    if not gcode_path.exists():
        raise FileNotFoundError(f"G-code file not found: {gcode_path}")

    if gcode_path.suffix.lower() != '.gcode':
        print('Warning: input file does not have a .gcode extension. This tool expects ASCII G-code.')

    compare_path = Path(args.compare) if args.compare else None
    if compare_path is not None and not compare_path.exists():
        raise FileNotFoundError(f"Compare G-code file not found: {compare_path}")

    if args.output:
        out_xlsx = Path(args.output)
    else:
        if compare_path is None:
            out_xlsx = gcode_path.with_suffix('.xlsx')
        else:
            out_xlsx = gcode_path.with_suffix('')
            out_xlsx = out_xlsx.with_name(f"{out_xlsx.name}_vs_{compare_path.stem}.xlsx")

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
    compare_moves = None
    compare_layer_z_map = None
    if compare_path is not None:
        status(f"Parsing G-code B ({compare_path.name})", status_enabled)
        compare_moves, compare_layer_z_map = parse_gcode(str(compare_path), filament_diam, status_cb=(lambda m: status(m, status_enabled)), status_every_lines=250_000)

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
        compare_moves=compare_moves,
        compare_layer_z_map=compare_layer_z_map,
        compare_label=(compare_path.stem if compare_path is not None else None),
        status_cb=(lambda m: status(m, status_enabled)),
    )

    status(f"Done -> {out_xlsx}", status_enabled)
    print(f"Wrote {len(moves)} moves to {out_xlsx}")

