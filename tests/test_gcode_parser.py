from gcode_profiler.gcode_parser import parse_gcode


def test_parse_gcode_infers_layers_from_z_comments(tmp_path):
    gcode = tmp_path / "sample.gcode"
    gcode.write_text(
        """
M83
;Z:0.2
;TYPE:Perimeter
G1 X0 Y0 F6000
G1 X10 Y0 E1.0 F1200
;Z:0.4
;TYPE:Infill
G1 X0 Y0 F6000
G1 X20 Y0 E2.0 F2400
""".lstrip(),
        encoding="utf-8",
    )

    moves, layer_z_map = parse_gcode(str(gcode), 1.75)
    assert len(moves) >= 3
    # should have at least 2 distinct layers from Z comments
    layers = sorted(set(m["layer"] for m in moves))
    assert len(layers) >= 2
    assert 0 in layer_z_map
    assert 1 in layer_z_map
    # at least one extruding move with positive flow
    assert any((m.get("flow_mm3_s") or 0) > 0 for m in moves)
