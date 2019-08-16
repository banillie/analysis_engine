import pytest

from ..process.bootstrap import AuxReport, AuxReportBlock, add_git_command


class TestAuxReportBlock(object):

    def test_base_AuxReportBlock_modified(self):
        rb = AuxReportBlock('modified')
        assert rb.output[0] == """********MODIFIED FILES********"""

    def test_base_AuxReportBlock_untracked(self):
        rb = AuxReportBlock('untracked')
        assert rb.output[0] == "*******UNTRACKED FILES********"

    def test_base_AuxReportBlock_log(self):
        rb = AuxReportBlock('log')
        assert rb.output[0] == "{:*^30}".format("LOG FILES")

    def test_manually_add_git_command(self):
        add_git_command('status', 'git status')
        rb = AuxReportBlock('status')
        assert rb.output[0] == "{:*^30}".format("STATUS FILES")


class TestAuxReportBase(object):

    def test_base_AuxReport(self):
        assert AuxReport.modified_files == []
        assert AuxReport.untracked_files == []


    def test_add_attribute_value_Auxreport(self):
        AuxReport.modified_files.append('test')
        assert AuxReport.modified_files[0] == 'test'

    @pytest.mark.skip("Only partial functionality implemented")
    def test_function_to_find_master_file(self):
        pass


    def test_add_AuxReport_instance(self):
        r = AuxReport()
        assert str(r) == "Report(['untracked', 'modified', 'log', 'add', 'commit', 'checkout', 'push'])"


    def test_change_instance_expect_attribute_change(self):
        AuxReport.modified_files.append('test')
        assert AuxReport.modified_files[0] == 'test'
        r = AuxReport()
        assert hasattr(r, 'modified_files')
        assert r.modified_files[0] == 'test'

    def test_for_non_existing_attribute(self):
        assert not hasattr(AuxReport, 'non-existant-attr')


    def test_dynamically_adding_attribute(self):
        AuxReport.add_check_component('log')
        r = AuxReport()
        assert r.log_files == []


    def test_wrong_component_type_added(self):
        with pytest.raises(TypeError) as excinfo:
            AuxReport.add_check_component(1)
        assert excinfo.value.args[0] == "component must be a string"


    def test_get_list_of_check_components_from_instance(self):
        r = AuxReport()
        r.add_check_component('log')
        assert r.check_components[-1] == 'log'
