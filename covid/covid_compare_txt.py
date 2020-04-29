'''
quick bit of code to compare reported narratives for covid-19 reporting. output is word document which
flags changes to text in red

requirements:
1) covid master wbs to go into analysis_engine core data directory
2) output file goes into output directory

follow instructions below.
'''


from datamaps.api import project_data_from_master
from analysis.engine_functions import compare_text_newandold
from analysis.data import root_path
from docx import Document

def printing():

     doc = Document()

     for project_name in latest_master.projects:
         heading = str(project_name)
         intro = doc.add_heading(str(heading), 0)
         intro.alignment = 1
         intro.bold = True

         new = doc.add_paragraph()
         new.add_run('')


         latest_txt = latest_master.data[project_name]['Project Narrative']
         if latest_txt is None:
             latest_txt = latest_master.data[project_name]['Group Narrative']

         try:
             last_txt = last_master.data[project_name]['Project Narrative']
             if last_txt is None:
                 last_txt = last_master.data[project_name]['Group Narrative']

             try:
                compare_text_newandold(latest_txt, last_txt, doc)
             except AttributeError:
                 pass

             new_2 = doc.add_paragraph()
             new_2.add_run('')

         except KeyError:

             sentence = 'First time project is reporting'
             p = doc.add_paragraph()
             runner = p.add_run(sentence)
             runner.bold = True
             doc.add_paragraph(latest_txt)
             doc.add_paragraph()


     return doc

'''RUNNING PROGRAMME INSTRUCTIONS'''

'''Insert into the two file paths below the name of latest and last master, after core_data/ .
Example master file name is master_170420.xlsx. Note make sure .xlsx is at end of file name. 
The rest of the file path information can remain unchanged'''
latest_master = project_data_from_master(root_path/'core_data/master_170420.xlsx', 1, 2020)
last_master = project_data_from_master(root_path/'core_data/master_020420.xlsx', 1, 2020)

run = printing()

'''The name of the output document can be changed via changing the file path here, if desired. The standard
name output name is covid_compare_text. Note make sure .docx is at end'''
run.save(root_path/'output/covid_compare_text.docx')