"""
Outputs analysis_engine for speed_dials. Outputs are place in analysis_engine/output. They are:
- A word document titled dca_changes which specifies which project dca ratings have changed
- matplotlib speed dial?
"""

from analysis_engine.data import (
    Master,
    root_path,
    dca_changes_into_word,
    get_project_information,
    get_master_data,
    get_word_doc,
    DcaData,
    gauge
)


def compile_speed_dials():
    m = Master(get_master_data(), get_project_information())
    dca = DcaData(m)
    latest_quarter = str(m.master_data[0].quarter)
    last_quarter = str(m.master_data[1].quarter)
    dca.get_changes(latest_quarter, last_quarter)
    word_doc = dca_changes_into_word(dca, get_word_doc())
    word_doc.save(root_path / "output/dca_changes.docx")


compile_speed_dials()


gauge(
    labels=["R", "A/R", "A", "A/G", "G"],
    colors=["#c00000", "#e77200", "#ffba00", "#92a700", "#007d00"],
    arrow=3,
    arrow_two=2,
    title="DCA OVERALL",
    fname=True,
)
