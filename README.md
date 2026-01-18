# GCodeProfiler

Convert **PrusaSlicer ASCII `.gcode`** into an **Excel workbook (`.xlsx`)** with per-move and per-layer metrics, computed legends, and Excel-native charts.

## Outputs

### Data sheets

- **Moves** (optional; see `--per-layer-only`)
  - distance, estimated time, speed (mm/s)
  - extrusion (mm of filament), volumetric flow (mm³/s)
  - fan (%), hotend/bed/chamber setpoints
  - feature type (from slicer comments like `;TYPE:...`)

- **Layers**
  - layer Z and computed layer height
  - total time, distance, extrusion per layer
  - average speed, average volumetric flow, average fan per layer
  - **peak**, **P95**, and **P99 (time-weighted)** speed/flow per layer (useful for tuning)
  - percent of layer time spent over configured speed/flow limits (when `--config` is provided)
  - last-known hotend/bed/chamber setpoints per layer

- **Legend_Speed**, **Legend_Flow_mm3s**, **Legend_Fan_pct**, **Legend_Temp_C**, **Legend_Bed_C**, **Legend_LayerHeight_mm**
  - min/max, equal-width bins, and bin counts (useful for histograms)

- **Legend_FeatureType**
  - PrusaSlicer-style feature breakdown per feature type: time, percentage of total time, used filament (m) and estimated mass (g)

- **FeatureType_Flow**
  - tuning-focused per-feature summary:
    - peak / P95 / P99 speed and volumetric flow
    - % of time over configured speed/flow limits (when `--config` is provided)

- **Top_Flow_Segments**
  - top-N extrusion segments ranked by volumetric flow (spike hunting)

- **FeatureFlow_Hist**
  - per-feature **time-weighted** flow histogram (which feature is pushing flow limits?)

- **Top_Slowest_Layers**
  - top-N layers ranked by `time_s` (also charted)

### Dashboard (first sheet)

The **Dashboard** (the first sheet) includes:

- Time per layer (line)
- Average speed per layer (line)
- Peak / P95 / P99 speed per layer (line)
- Average volumetric flow per layer (line)
- Peak / P95 / P99 volumetric flow per layer (line)
- Extrusion per layer (line)
- Average fan per layer (line)
- Temperature setpoints per layer (line: hotend/bed/chamber)
- Layer height per layer (column)
- Speed histogram (column, from `Legend_Speed` bin counts)
- Flow histogram (column, from `Legend_Flow_mm3s` bin counts)
- Top-N slowest layers (column)
- Feature type time-share (pie)
- Feature type table (top 10): time, percentage, used filament (m) and estimated mass (g)

When `--compare` is used, the dashboard also includes overlay compare charts (A vs B) for layer time, peak flow, P95 flow, and peak speed.

#### Chart readability (especially on Excel for macOS)

- The dashboard uses a two-column grid with extra horizontal separation to avoid chart overlap.
- X-axis labels on layer-based charts are automatically **down-sampled** (skip most labels) and **rotated** for readability.

## Install

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Basic

```bash
python gcode_profiler.py path/to/print.gcode
```

Output:

- `path/to/print.xlsx`

### Options

```bash
python gcode_profiler.py print.gcode   --bins 12   --filament-diameter 1.75   --top-n-slowest 25
```

Flags:

- `--output out.xlsx` : override output path
- `--compare other.gcode [more.gcode ...]` : compare against one or more other files (adds `Compare_*` sheets and overlay charts for the first compare)
- `--config config.ini` : optional PrusaSlicer config.ini (key=value) to improve defaults and chart scaling
- `--bins N` : number of legend bins used for histograms
- `--filament-diameter D` : filament diameter (mm) used for volumetric flow
- `--filament-density R` : filament density (g/cm³) used for Used filament (g) estimates (default: 1.24)
- `--no-legends` : skip `Legend_*` sheets (histogram charts will be skipped)
- `--per-layer-only` : skip per-move rows for smaller files
- `--top-n-slowest N` : how many slowest layers to list and chart
- `--layout compact|wide` : dashboard layout (compact packs charts tighter; wide uses larger charts/spacing)
- `--top-n-segments N` : how many top flow segments to include in `Top_Flow_Segments` (default: 200)
- `--csv` : write sidecar CSV exports next to the xlsx (layers + top flow segments + feature flow histogram)
- `--json-summary` : write a small JSON summary next to the xlsx (useful for regression tests)
- `--no-manifest` : disable writing the `.run.json` manifest next to the xlsx
- `--status-interval N` : print a status line every N parsed lines (default: 250000)
- `--quiet` : suppress progress output

### Progress output

By default, the tool prints lightweight progress messages to the terminal (stderr) so you can see it moving through large files. Example:

```text
[  0.0s] Parsing G-code A (print.gcode)
[  1.2s] Parsed 250,000 lines
[  2.6s] Parsed 500,000 lines
[  3.1s] Building Excel workbook
[  3.4s] Saving workbook
[  3.5s] Done -> print.xlsx
```

Use `--quiet` to disable these messages.

### Run manifest

By default, GCodeProfiler also writes a small JSON manifest next to the workbook:

- `print.run.json`

It includes input paths, CLI arguments, a timestamp, and SHA256 hashes (handy for provenance and CI artifacts). Disable with `--no-manifest`.

## Testing

Run unit + integration tests locally:

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

## GitHub Actions

This repo includes a GitHub Actions workflow that runs `pytest` on pushes and pull requests.

### Compare two G-code files

This is handy when you change one setting (nozzle, speeds, cooling, etc) and want an A/B view.

```bash
python gcode_profiler.py A.gcode --compare B.gcode
```

Output (default):

- `A_vs_B.xlsx`

Additional sheets:

- `Compare_Layers`: A/B layer-aligned metrics (time, peak/p95 speed/flow)
- `Compare_Summary`: headline differences (includes all compare files)

### CSV / JSON sidecars

If you pass `--csv`, the tool will write:

- `*_layers.csv`
- `*_top_flow_segments.csv`
- `*_feature_flow_hist.csv`

If you pass `--json-summary`, the tool will write:

- `*.summary.json`

By default, the tool also writes a `*.run.json` manifest with SHA-256 hashes of inputs and outputs (disable with `--no-manifest`).
### Using PrusaSlicer `config.ini` (optional)

You can pass a PrusaSlicer-exported `config.ini` to make the workbook **more comparable to what you see in PrusaSlicer**:

- Uses `filament_diameter` / `filament_density` as defaults (unless you override via CLI)
- Uses `filament_max_volumetric_speed` and `max_print_speed` to:
  - set nicer chart Y-axis ranges
  - add constant **reference-line series** on the Speed and Flow charts
  - set histogram / legend ranges to match your configured limits
- Uses `min_layer_height` / `max_layer_height` to scale the Layer Height chart and bins
- Adds **conditional formatting** in the `Layers` sheet to highlight:
  - average flow above `filament_max_volumetric_speed`
  - average speed above `max_print_speed`
  - peak and P95 flow above `filament_max_volumetric_speed`
  - peak and P95 speed above `max_print_speed`
  - layer heights outside `min_layer_height` / `max_layer_height`

Example:

```bash
python gcode_profiler.py print.gcode --config config.ini
```

## Important: `.bgcode` (binary G-code)

This tool expects **ASCII `.gcode`**. If you have Prusa binary **`.bgcode`**, convert it first in PrusaSlicer:

- **File → Convert** → `.bgcode` → `.gcode`

Then run the script on the resulting `.gcode`.

## Notes / Limitations

- Move timing is estimated from feedrate and distance (acceleration is not modeled).
- Feature types rely on slicer comments like `;TYPE:...`.

## License

MIT

## Project layout

The implementation is split into a small package so it is easier to maintain:

- `gcode_profiler/cli.py` – argument parsing and orchestration
- `gcode_profiler/gcode_parser.py` – G-code parsing into per-move records
- `gcode_profiler/excel_writer.py` – workbook creation (sheets, tables, charts)
- `gcode_profiler/stats.py` – weighted percentiles + binning utilities
- `gcode_profiler/config_ini.py` – optional `config.ini` parsing helpers
- `gcode_profiler/constants.py` – defaults

The original `gcode_profiler.py` script remains as a thin wrapper for backwards compatibility.
