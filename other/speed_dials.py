"""
Outputs analysis_engine for speed_dials. Outputs are place in analysis_engine/output. They are:
- A word document titled dca_changes which specifies which project dca ratings have changed
- matplotlib speed dial?
"""

from analysis_engine.data import (
    root_path,
    dca_changes_into_word,
    get_word_doc,
    DcaData,
    gauge,
    open_pickle_file
)


# m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
# dca = DcaData(m, quarter=["standard"])
# word_doc = dca_changes_into_word(dca, get_word_doc())
# word_doc.save(root_path / "output/dca_changes.docx")

# labels = ["R \n1", "A/R\n2", "A\n3", "A/G\n7", "G\n89"]
colours = ["#c00000", "#e77200", "#ffba00", "#92a700", "#007d00"]
count = [1, 2, 3, 69, 24]
gauge(
    count,
    colours,
    # count,
    2.33,
    2.55,
    title="DCA OVERALL",
    fname=True,
)
