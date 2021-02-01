from analysis_engine.data import open_pickle_file, root_path, DandelionData, dandelion_data_into_wb

m = open_pickle_file(str(root_path / "core_data/pickle/master.pickle"))
dand = DandelionData(m, group=["RDM"])
dandelion_data_into_wb(dand)


