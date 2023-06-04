import departmentdata as college

import collections as coll
import openpyxl as op


def getsummaryname(survey_name: str) -> str:
    if (survey_name.strip() != ''):
        return f"Summary - {survey_name}.xlsx"
    return f"Summary.xlsx"

def Create_Summary(dirpath: str,  Raw_scores: list,survey_name: str=''):

    wb = op.Workbook()
    sheet = wb.active
    
    Total_rows = len(Raw_scores)
    for i in range(0,Total_rows):
        sheet.append(Raw_scores[i][0:])


    wb.save(f"{dirpath}/{getsummaryname(survey_name)}")


