import datetime

import pytest

from ..core import FinancialYear, Quarter


def test_fy():
    fy1 = FinancialYear(2017)
    assert str(fy1) == "FY2017/18"
    fy1 = FinancialYear(2012)
    assert str(fy1) == "FY2012/13"
    fy1 = FinancialYear(2013)
    assert str(fy1) == "FY2013/14"
    fy1 = FinancialYear(1999)
    assert str(fy1) == "FY1999/00"


def test_errors():
    with pytest.raises(ValueError) as excinfo:
        FinancialYear("2017")
    assert "A year must be an integer between 1950 and 2100" in str(excinfo.value)

    with pytest.raises(ValueError) as excinfo:
        FinancialYear(202332)
    assert "A year must be an integer between 1950 and 2100" in str(excinfo.value)


def test_quarters_in_fy():
    fy = FinancialYear(2016)
    assert str(fy.q1) == "Q1 16/17"
    assert fy.q1.start_date == datetime.date(2016, 4, 1)
    assert fy.q1.end_date == datetime.date(2016, 6, 30)
    assert fy.q4.start_date == datetime.date(2017, 1, 1)
    assert fy.start_date == datetime.date(2016, 4, 1)
    assert fy.end_date == datetime.date(2017, 3, 31)

    fy = FinancialYear(1999)
    assert str(fy.q1) == "Q1 99/00"
    assert fy.q1.start_date == datetime.date(1999, 4, 1)
    assert fy.q1.end_date == datetime.date(1999, 6, 30)
    assert fy.q4.start_date == datetime.date(2000, 1, 1)
    assert fy.start_date == datetime.date(1999, 4, 1)
    assert fy.end_date == datetime.date(2000, 3, 31)


def test_forbid_setting_quarters_manually():
    fy = FinancialYear(2008)
    q = Quarter(1, 2009)
    with pytest.raises(AttributeError) as excinfo:
        fy.q1 = q
    assert "can't set attribute" in str(excinfo.value)
    fy = FinancialYear(2018)
    q = Quarter(1, 2009)
    with pytest.raises(AttributeError) as excinfo:
        fy.q1 = q
    assert "can't set attribute" in str(excinfo.value)
    fy = FinancialYear(2019)
    q = Quarter(1, 2019)
    with pytest.raises(AttributeError) as excinfo:
        fy.q1 = q
    assert "can't set attribute" in str(excinfo.value)


def test_fy_from_quarter():
    q1_2017 = Quarter(1, 2017)
    assert q1_2017.fy.start_date == datetime.date(2017, 4, 1)
