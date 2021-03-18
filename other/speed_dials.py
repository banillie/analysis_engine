"""
Outputs analysis_engine for speed_dials. Outputs are place in analysis_engine/output. They are:
- A word document titled dca_changes which specifies which project dca ratings have changed
- matplotlib speed dial?
"""

from typing import List
import numpy as np

from analysis_engine.data import (
    root_path,
    dca_changes_into_word,
    get_word_doc,
    DcaData,
    gauge,
    open_pickle_file,
    open_word_doc,
    put_matplotlib_fig_into_word,
    build_speedials
)


m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
dca = DcaData(m, quarter=["standard"], type=["SRO"])
landscape_doc = open_word_doc(root_path / "input/summary_temp_landscape.docx")
# word_doc = dca_changes_into_word(dca, get_word_doc())
# word_doc.save(root_path / "output/dca_changes.docx")
build_speedials(dca, landscape_doc)
landscape_doc.save(root_path / "output/speedial_graph.docx")