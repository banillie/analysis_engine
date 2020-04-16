'''
quick bit of code to compare reported narratives for covid-19 reporting. output is word document which
flags changes to text in red

requirements:
1) covid master wbs to go into analysis_engine core data directory
2) output file goes into output directory
'''


from datamaps.api import project_data_from_master
from analysis.engine_functions import compare_text_newandold
from analysis.data import root_path
from docx import Document

master_17_4_20 = project_data_from_master(root_path/'core_data/master_170420.xlsx', 1, 2020)
master_2_4_20 = project_data_from_master(root_path/'core_data/master_020420.xlsx', 1, 2020)

def printing():

     doc = Document()

     for project_name in master_17_4_20.projects:
         heading = str(project_name)
         intro = doc.add_heading(str(heading), 0)
         intro.alignment = 1
         intro.bold = True

         new = doc.add_paragraph()
         new.add_run('')

         latest_txt = master_17_4_20.data[project_name]['Narrative']
         last_txt = master_2_4_20.data[project_name]['Narrative']

         try:
            compare_text_newandold(latest_txt, last_txt, doc)
         except AttributeError:
             pass

         new_2 = doc.add_paragraph()
         new_2.add_run('')

     return doc

run = printing()

run.save(root_path/'output/covid_compare_text.docx')