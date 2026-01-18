from gcode_profiler.config_ini import parse_config_ini, config_get_float


def test_parse_config_ini_basic(tmp_path):
    p = tmp_path / "config.ini"
    p.write_text(
        """
# comment
filament_diameter = 1.75
filament_density = 1.24
filament_max_volumetric_speed = 8
max_print_speed = 200
""".lstrip(),
        encoding="utf-8",
    )

    cfg = parse_config_ini(str(p))
    assert cfg["filament_diameter"] == "1.75"
    assert config_get_float(cfg, "filament_diameter") == 1.75
    assert config_get_float(cfg, "max_print_speed") == 200.0
