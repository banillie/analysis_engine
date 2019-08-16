# digest.py

#
# Pull data from an Excel form, based on a datamap.
import os
import fnmatch
import re

from datetime import datetime
from typing import Dict

from concurrent import futures

from bcompiler.compile import parse_source_cells
from bcompiler.utils import DATAMAP_MASTER_TO_RETURN


class Digest:
    """
    A Digest object is a compilation of key/value pairs from a specific Excel
    file. By default, the Digest is serialized and written to a database.
    """
    def __init__(self, file_name, series, series_item):
        self.file_name = file_name
        self.series = series.__str__()
        self.table = self.tableize(series_item)
        self._data = self._digest_source_file(file_name)

    def tableize(self, item):
        return re.sub('\s', '-', item).lower()

    def flatten_project(self, project_data):
        """
        Get rid of the gmpp_key gmpp_key_value stuff pulled from a single
        spreadsheet. Must be given a future.
        """
        return {
            item['gmpp_key']: item['gmpp_key_value'] for item in project_data}

    def _digest_source_file(self, file_name):
        flat = self.flatten_project(
            parse_source_cells(file_name, DATAMAP_MASTER_TO_RETURN))
        return flat

    @property
    def data(self):
        return self._data


def flatten_project(future) -> Dict[str, str]:
    """
    Get rid of the gmpp_key gmpp_key_value stuff pulled from a single
    spreadsheet. Must be given a future.
    """
    p_data = future.result()
    p_data = {item['gmpp_key']: item['gmpp_key_value'] for item in p_data}
    return p_data


def digest_source_files(base_dir, db_connection) -> None:
    source_files = []
    future_data = []
    for f in os.listdir(base_dir):
        if fnmatch.fnmatch(f, '*.xlsx'):
            source_files.append(os.path.join(base_dir, f))
    with futures.ThreadPoolExecutor(max_workers=4) as executor:
        for f in source_files:
            future_data.append(executor.submit(
                parse_source_cells, f, DATAMAP_MASTER_TO_RETURN))
            print("Processing {}".format(f))
        for future in futures.as_completed(future_data):
            f = flatten_project(future)
            db.insert(f)


def main():
    digest_source_files()


if __name__ == "__main__":
    main()
