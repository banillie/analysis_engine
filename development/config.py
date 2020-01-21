import platform
from pathlib import Path
from datamaps.api import project_data_from_master

def _platform_docs_dir() -> Path:
    if platform.system() == "Linux":
        return Path.home() / "Documents" / "analysis_engine" / "core_data"
    if platform.system() == "Darwin":
        return Path.home() / "Documents" / "analysis_engine" / "core_data"
    else:
        return Path.home() / "Documents" / "analysis_engine" / "core_data"

root_path = _platform_docs_dir()
q3_1920 = project_data_from_master(root_path/'master_3_2019.xlsx', 3, 2019)