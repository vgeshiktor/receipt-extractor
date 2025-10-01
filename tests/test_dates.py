from app import compute_dates

def test_months_back_two_shapes_dates():
    after, before = compute_dates(None, None, 2)
    assert after is not None and before is not None
    assert after.count("/") == 2 and before.count("/") == 2
