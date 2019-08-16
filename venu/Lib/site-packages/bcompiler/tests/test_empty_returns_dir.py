import pytest

from ..utils import directory_has_returns_check


@pytest.mark.skip("capsys not working in this case - don't know why")
def test_empty_returns_dir_throws_exception(capsys, tmpdir):
    d = tmpdir.mkdir("returns")
    directory_has_returns_check(d)
    out, err = capsys.readouterr()
    assert err == "CRITICAL Please copy populated return files to returns directory."
