from ..process.datamap import Datamap
from ..process.cell import Cell


def test_datamap_object():
    d = Datamap()
    assert d.__class__.__name__ == 'Datamap'


def test_initial_state():
    d = Datamap()
    assert d.cell_map == []


def test_datamap_from_csv(datamap):
    d = Datamap()
    d.cell_map_from_csv(datamap)
    assert isinstance(d.cell_map[0], Cell)
    assert d.cell_map[0].cell_key == 'Project/Programme Name'
    assert d.cell_map[0].cell_reference == 'B5'


