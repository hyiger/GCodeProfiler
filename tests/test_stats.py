from gcode_profiler.stats import weighted_quantile, make_bins, bin_counts


def test_weighted_quantile_basic():
    vals = [0.0, 10.0, 20.0, 30.0]
    w = [1.0, 1.0, 1.0, 1.0]
    assert weighted_quantile(vals, w, 0.0) == 0.0
    assert weighted_quantile(vals, w, 1.0) == 30.0

    q50 = weighted_quantile(vals, w, 0.5)
    assert 10.0 <= q50 <= 20.0


def test_weighted_quantile_weighted_bias():
    vals = [0.0, 100.0]
    w = [9.0, 1.0]
    # With 90% of weight at 0, the 0.5 quantile should be 0
    assert weighted_quantile(vals, w, 0.5) == 0.0


def test_bins_and_counts():
    bins_spec = make_bins(0.0, 10.0, 5)
    assert len(bins_spec) == 5
    values = [0.1, 0.2, 1.9, 2.1, 9.9]
    counts = bin_counts(values, bins_spec)
    assert sum(counts) == len(values)
