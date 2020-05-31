import matplotlib.pyplot as plt
from analysis.data import latest_cost_profiles, last_cost_profiles, baseline_1_cost_profiles, \
    year_list, root_path, a66, a303, crossrail, thameslink
from docx import Document
from docx.shared import Inches

doc = Document()
amended_year_list = year_list[:-1]
#plt.style.use('fivethirtyeight') #maybe? or sort graph style?

def get_financial_data(project_name):
    latest = []
    last = []
    baseline = []
    for year in amended_year_list:
        baseline.append(baseline_1_cost_profiles[project_name][year + ' total'])
        last.append(last_cost_profiles[project_name][year + ' total'])
        latest.append(latest_cost_profiles[project_name][year + ' total'])

    return latest, last, baseline

def put_in_word(project_name_list):
    for project_name in project_name_list:
        y = doc.add_paragraph()
        heading = 'Financial Analysis - Cost Profile'
        y.add_run(str(heading)).bold = True
        y = doc.add_paragraph()

        '''financial chart'''

        profile_data = get_financial_data(project_name)

        year = ['19/20', '20/21', '21/22', '22/23', '23/24', '24/25', '25/26', '26/27', '27/28', '28/29']
        baseline_profile = profile_data[2]
        last_profile = profile_data[1]
        latest_profile = profile_data[0]


        fig, ax = plt.subplots(figsize=(20,10))
        #plt.style.use('fivethirtyeight')
        #plt.style.use('seaborn-muted')

        plt.subplots_adjust(left=0.05, bottom=0.15, right=0.92, top=0.92, wspace=0.2, hspace=0.2)

        ax.plot(year, baseline_profile, color='blue', label='baseline', linewidth=4.0, marker="o")
        ax.plot(year, last_profile, color='yellow', label='last quarter', linewidth=4.0, marker="o")
        ax.plot(year, latest_profile, color='green', label='latest', linewidth=4.0, marker="o")

        ax.set_xlabel('Financial Years')
        ax.set_ylabel('Cost (£m)')
        xlab = ax.xaxis.get_label()
        ylab = ax.yaxis.get_label()
        xlab.set_style('italic')
        xlab.set_size(20)
        ylab.set_style('italic')
        ylab.set_size(20)

        ax.set_title(str(project_name) + ' Cost Profile Changes')  # is this needed?

        ax.grid(color='grey', linestyle='-', linewidth=0.2)

        ax.legend(borderpad=2)

        ax.figure.savefig('cost_profile.png')

        doc.add_picture('cost_profile.png', width=Inches(5.8)) #to place nicely in doc

        doc.save(root_path/'output/graph.docx')

project_list = [a66]
                #a303, crossrail, thameslink]

put_in_word(project_list)