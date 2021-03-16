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


m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
dca = DcaData(m, quarter=["standard"])
# word_doc = dca_changes_into_word(dca, get_word_doc())
# word_doc.save(root_path / "output/dca_changes.docx")

for conf_type in dca.dca_count["Q3 20/21"]:
    count = []
    for colour in ["Green", "Amber/Green", "Amber", "Amber/Red", "Red"]:
        no = dca.dca_count["Q3 20/21"][conf_type][colour][0]
        count.append(no)
    total = dca.dca_count["Q3 20/21"][conf_type]["Total"][0]
    up = 0
    down = 0
    for p in dca.dca_changes[conf_type]:
        change = dca.dca_changes[conf_type][p]["Change"]
        if change == "Up":
            up += 1
        if change == "Down":
            down += 1

    gauge(
        count,
        str(total),
        2.5,
        3.5,
        str(up),
        str(down),
        title=conf_type,
    )
