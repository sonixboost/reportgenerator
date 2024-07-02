import datetime
import pandas as pd
import os
from docx import Document

#################### SETTINGS ########################
EXCEL_SHEET = 'No File Selected'
required_cols = [1,2,7,27,29]
save_path = os.getcwd()

document = Document()
# Set font
style = document.styles['Normal']
style.font.name = 'Calibri'
########################################################

def top_producer(producers: pd.DataFrame):
    producers_dict = {}
    for i,j in producers.iterrows():
        if j["Consultant Name"] not in producers_dict:
            producers_dict[j["Consultant Name"]] = [j["Amount"], j["Branch"]]
        else:
            producers_dict[j["Consultant Name"]][0] += j["Amount"]
            # print("Added: ", j["Amount"], " to ", j["Consultant Name"])
            # assert producers_dict[j["Consultant Name"]][1] == j["Branch"]
    # print(producers_dict)
    sorted_producers = sorted(producers_dict.items(), key=lambda x: x[1][0], reverse=True)
    # print(sorted_producers)

    document.add_paragraph("Top Producers:")
    global total_amount
    total_amount = 0
    for producer in sorted_producers:
        total_amount += producer[1][0]
        currency = ('{:,}'.format(producer[1][0]))
        document.add_paragraph("{0}~ {1} - RM{2}".format(producer[1][1], producer[0], currency),\
                               style='List Number')
    document.add_paragraph()
    total_amount = ('{:,}\n'.format(total_amount))

def top_case(cases: pd.DataFrame) -> tuple:
    case_dict = {}
    for i,j in cases.iterrows():
        number_of_case = 2 if j["Amount"] == 3299 or j["Amount"] == 2804 else 1
        if j["Consultant Name"] in case_dict:
            case_dict[j["Consultant Name"]][1] += number_of_case
        else:
            case_dict[j["Consultant Name"]] = [j["Branch"], number_of_case]
    # print(case_dict)
    sorted_cases = sorted(case_dict.items(), key=lambda x: x[1][1], reverse=True)
    # print(sorted_cases)
    
    document.add_paragraph("Top Cases:")
    global total_number_of_cases
    total_number_of_cases = 0
    for cases in sorted_cases:
        total_number_of_cases += cases[1][1]
        if cases[1][1] <= 1:
            document.add_paragraph("{0}~ {1} - {2} CASE".format(cases[1][0], cases[0], cases[1][1]),\
                                   style='List Number 2')
        else:
            document.add_paragraph("{0}~ {1} - {2} CASES".format(cases[1][0], cases[0], cases[1][1]),\
                                   style='List Number 2')
    document.add_paragraph()
    
def branch(branches: pd.DataFrame) -> str:
    branch_dict = {}
    for i,j in branches.iterrows():
        number_of_case = 2 if j["Amount"] == 3299 or j["Amount"] == 2804 else 1
        if j["Branch"] not in branch_dict:
            branch_dict[j["Branch"]] = number_of_case
        else:
            branch_dict[j["Branch"]] += number_of_case
    # print(branch_dict)
    sorted_cases = sorted(branch_dict.items(), key=lambda x: x[1], reverse=True)
    # print(sorted_cases)

    document.add_paragraph("Top Branches:")
    for branch in sorted_cases:
        if branch[1] <= 1:
            document.add_paragraph("{0}~ {1} CASE".format(branch[0], branch[1]), style="List Number 3")
        else:
            document.add_paragraph("{0}~ {1} CASES".format(branch[0], branch[1]), style="List Number 3")
    document.add_paragraph()

def date(dates: pd.DataFrame) -> list:
    start_date = dates[0].replace('-', ' ')
    end_date = dates[len(dates)-1].replace('-', ' ')
    return [start_date, end_date]

def generate_report(custom_name: str):
    try:
        ########### TO WRITE ##############
        dt1 = datetime.datetime.now()

        data = pd.read_excel(EXCEL_SHEET, usecols= required_cols)
        for i,j in data.iterrows():
            data = data.replace(j["Branch"], j["Branch"].split(" ").pop())
            if j["Status"] == "REJECTED":
                data = data.drop(i)
        # print(data)
        # print(data[["Branch","Consultant Name","Amount"]])
        date_output = date(data["Submission Date"])

        global report
        if not custom_name:
            filename = "Production Report {0}".format(datetime.datetime.now().isoformat(timespec='seconds'))
            filename = filename.replace(":", '-')
        else:
            filename = "{0}".format(custom_name)

        filename = os.path.join(save_path, filename + ".docx")
        # processed = processed_data
        # print(filename)

        print("Generating report...")
        para_header = document.add_paragraph("Month-to-date Production:\n{0} - {1}\nRM "\
                                             .format(date_output[0], date_output[1]))
        
        top_producer(data[["Branch","Consultant Name", "Amount"]])
        para_header.add_run(total_amount)
        top_case(data[["Branch", "Consultant Name", "Amount"]])
        branch(data[["Branch", "Amount"]])
        document.add_paragraph("TOTAL CASES: "+ str(total_number_of_cases)+ " CASES\n")

        document.add_paragraph("Your dedication to protecting and preserving family wealth distribution is truly \
commendable. You are not just selling products; you are safeguarding the future and \
ensring the well-being of countless families. \n\nKeep up the fantastic work, and may \
your efforts continue to help secure the financial legacies of many more families.\
\n\nOnce again, congratulations on your well-deserved success!")
        
        document.save(filename)
    
    except Exception as e:
        report = "Something went wrong. Please try again later."
        print(report)
        print(e)
    
    else:
        dt2 = datetime.datetime.now()
        td = (dt2 - dt1).total_seconds() *10**3
        td = str(round(td, 2))
        report = ("Report done in {0} ms.\nThe file is saved as: {1}".format(td, filename))
        print(report)
        return report


if __name__ == "__main__": 
    generate_report("")