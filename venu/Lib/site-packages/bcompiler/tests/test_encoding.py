import pytest

from ..compile import encode_win


@pytest.mark.skip("Until we can work out how to get cp1252 encoding in here")
def test_cp1252_encode():
    wind_string = "£30"
    assert encode_win(wind_string) == "£30"
