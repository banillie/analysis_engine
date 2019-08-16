import logging

from openpyxl import load_workbook

from bcompiler.utils import index_returns_directory
from bcompiler.process import Cleanser

logger = logging.getLogger('bcompiler.process.simple_comparitor')


class BCCell:

    def __init__(self, value, row_num=None, col_num=None, cellref=None):
        self.value = value
        self.row_num = row_num
        self.col_num = col_num
        self.cellref = cellref


class ParsedMaster:

    def __init__(self, master_file):
        self.master_file = master_file
        self._projects = []
        self._project_count = None
        self._key_col = []
        self._wb = load_workbook(self.master_file)
        self._ws = self._wb.active
        self._project_header_index = {}
        self._parse()

    def _cleanse_key(self, key):
        c = Cleanser(key)
        return c.clean()

    def _parse(self):
        """
        Private method to set up the class.
        self._key_col is column 'A' in the masters format.
        """
        self._projects = [cell.value for cell in self._ws[1][1:]]
#       self._projects.sort()
        self._project_count = len(self.projects)
        self._key_col = [self._cleanse_key(cell.value) for cell in self._ws['A']]
        self._index_projects()

    @property
    def projects(self):
        """
        Returns a list of project titles in the master.
        """
        return self._projects

    def _create_single_project_tuple(self, column=None, col_index=None):
        """
        Private method to construct a tuple of key, values based on
        the particular project (identified by reference to the its column,
        and can be given as a letter ('H') or an integer.

        This method is internal and is called by self.get_project_data.
        """
        if col_index is None:
            col_data = self._ws[column]
            z = list(zip(self._key_col, col_data))
            return [((item[0]), (item[1].value)) for item in z]
        else:
            col_data = []
            for row in self._ws.iter_rows(
                min_row=1,
                max_col=col_index,
                min_col=col_index,
                max_row=len(self._key_col)
            ):
                count = 0
                for cell in row:
                    col_data.append(cell.value)
                    count += 1
            z = list(zip(self._key_col, col_data))
            return [((item[0]), (item[1])) for item in z]

    def _index_projects(self):
        self._project_header_index = {}
        for cell in self._ws[1]:
            if cell.value in self.projects:
                self._project_header_index[cell.value] = cell.col_idx

    def print_project_index(self):
        print('{:<68}{:>5}'.format("Project Title", "Column Index"))
        print('{:*^80}'.format(''))
        for k, v in self._project_header_index.items():
            print('{:<68}{:>5}'.format(k, v))

    def _create_dict_all_project_tuples(self):
        pass

    def __repr__(self):
        return "ParsedMaster for {}".format(
            self.master_file
        )

    def get_project_data(self, column=None, col_index=None):
        if column is None and col_index is None:
            raise TypeError('Please include at least one param')

        if column == 'A':
            raise TypeError("column must be 'B' or later in alphabet")

        if column:
            if isinstance(column, type('b')):
                data = self._create_single_project_tuple(column)
            else:
                raise TypeError('column must be a string')

        if col_index:
            if isinstance(col_index, type(1)):
                data = self._create_single_project_tuple(col_index=col_index)
            else:
                raise TypeError('col_index must be an integer')

        return data

    def _query_for_key(self, data, key):
        """
        Iterate through keys in output from get_project_data
        data list and return True if a key is found. Does not return
        anything if not found.
        """
        for item in data:
            if item[0] == key:
                self._query_result = item[1]
                return True

    def get_data_with_key(self, data, key):
        """
        Given a data list with project key/values in it (derived from
        a master spreadsheet, query a specific key to return a value.
        """
        # first query that the value exists
        if self._query_for_key(data, key):
            return self._query_result
        else:
            logger.warning("No key {} in comparing master. Check for double spaces in cell in master. Skipping".format(key))
            return None

    def index_target_files_with_previous_master(self):
        """
        A previous master has a column-order of projects. If we are going
        to compare this with a series of projects used in bcompiler compile,
        which traverses a target directory and compiles each in turn into
        a master spreadsheet, the order must match, otherwise comparing
        values will not work.

        This function first gets obtains the order of project names from the
        files in the 'returns' directory, the it obtains the order or projects
        from the column headers in the master file from this object.
        """
        target_project_names = index_returns_directory()
        master_title_names = [
            key for key, value in self._project_header_index.items()]
        return (target_project_names, master_title_names)


def populate_cells(worksheet, bc_cells=[]):
    """
    Populate a worksheet with bc_cell object data.
    """
    for item in bc_cells:
        if item.cellref:
            worksheet[item.cellref].value = item.value
        else:
            worksheet.cell(
                row=item.row_num, column=item.col_num, value=item.value)
    return worksheet


class FileComparitor:
    """
    Simple method of comparing data in two master spreadsheets.
    """

    def __init__(self, masters=[]):
        """
        We want to get a list of master spreadsheets. These are simple
        file-references. The latest master should be master[-1].
        """
        self._comp_type = None

        if len(masters) > 2:
            raise ValueError("You can only analyse two spreadsheets.")

        if len(masters) == 2:
            # we're comparing two files
            self._masters = masters
            self._comp_type = 'two'
            self._get_data()

        if len(masters) == 1:
            # we're comparing against one single master
            self._master = masters[0]
            self._comp_type = 'one'
            self._get_data()

    def _get_data(self):
        """
        Private method that creates two ParsedMaster objects in a tuple. First
        is the earlier master, the second is the current. These states are
        derived from the order that the file references are given to the
        constructor.
        """
        if self._comp_type == 'two':
            self._early_master = ParsedMaster(self._masters[0])
            self._current_master = ParsedMaster(self._masters[1])
            return (self._early_master, self._current_master)

        if self._comp_type == 'one':
            self._early_master = ParsedMaster(self._master)
            return self._early_master

    @property
    def data(self):
        return self._early_master

    def compare(self, proj_index, key):
        """
        Returns a tuple of two values, the first is the value of key in
        proj_index in the early master, the second the equivalent in the
        current master. proj_index should be an integer and can be derived
        from the import spreadsheet or by ParsedMaster.print_project_index.
        """
        if self._comp_type == 'two':
            project_data_early = self._early_master.get_project_data(
                col_index=proj_index)
            project_data_current = self._current_master.get_project_data(
                col_index=proj_index)
            return(
                self._early_master.get_data_with_key(
                    project_data_early, key),
                self._current_master.get_data_with_key(
                    project_data_current, key))

        if self._comp_type == 'one':
            project_data_early = self._early_master.get_project_data(
                col_index=proj_index)
            return(
                self._early_master.get_data_with_key(project_data_early, key))
