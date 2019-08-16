import csv
import sys
from typing import List

from .cell import Cell


class Datamap:
    """
    Purpose of the Datamap is to map key/value sets to the database and a
    FormTemplate class. A Datamap comprises a list of Cell objects.

    Newly initialised Datamap object contains a template and a reference to
    a SQLite database file, but it's cell_map is empty. To create a base cell
    map from the template, using the datamap table in the database, call
    Datamap.cell_map_from_database(). To create a base cell map from the
    template, call Datamap.cell_map_from_csv().
    """

    def __init__(self) -> None:
        self.cell_map: List[Cell] = []

    def add_cell(self, cell: Cell) -> Cell:
        self.cell_map.append(cell)
        return cell

    def delete_cell(self, cell: Cell) -> Cell:
        self.cell_map.remove(cell)
        return cell

    def cell_map_from_csv(self, source_file: str) -> None:
        """
        Read from a CSV source file. Returns a list of corresponding Cell
        objects.
        """
        if source_file[-4:] == '.csv':
            try:
                self._import_source_data(source_file)
            except UnicodeDecodeError:
                print("There is a problem with the CSV file. Please ensure "
                      "it is saved as UTF-8 encoding after you create or edit it"
                      " (i.e. in Notepad, Excel, LibreOffice, etc.\nCannot continue."
                      "Exiting.")
                sys.exit(1)

    def _open_with_encoding_and_extract_data(self, source_file, encoding):
        with open(source_file, 'r', encoding=encoding) as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                self.cell_map.append(Cell(cell_key=row['cell_key'], cell_value=None,  # have no need of a value in dm
                                          cell_reference=row['cell_reference'], template_sheet=row['template_sheet'],
                                          bg_colour=None, fg_colour=None, number_format=None, verification_list=None))

    def _import_source_data(self, source_file: str) -> None:
        """Internal implementation of csv importer."""
        try:
            self._open_with_encoding_and_extract_data(source_file, "ISO-8859-1")
        except UnicodeDecodeError:
            self._open_with_encoding_and_extract_data(source_file, "utf-8")
        except Exception:
            print(f"Cannot decode file {source_file}")
            sys.exit(1)
