def test_import_excel_writer_module():
    # Smoke-test that the Excel writer module imports cleanly.
    # This catches missing openpyxl symbols (e.g., Series) early.
    import gcode_profiler.excel_writer  # noqa: F401
