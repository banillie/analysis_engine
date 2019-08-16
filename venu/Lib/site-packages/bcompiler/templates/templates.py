from openpyxl import load_workbook

from bcompiler.utils import TEMPLATE


class CommissioningTemplate():
    """
    A blank xlsx file, returned as an openpyxl object, ready for populating.
    Eventually, it will end up as a PopulatedCommissingTemplate (which has
    project data in it.
    """

    def __init__(self):
        self.source_file = TEMPLATE
        self.openpyxl_obj = self._load_workbook()
        self.sheets = self.openpyxl_obj.sheetnames
        self.blank = True

    def _load_workbook(self):
        return load_workbook(filename=self.source_file, keep_vba=True)
