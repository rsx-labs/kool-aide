import json
import pandas as pd
import numpy as nm
from datetime import datetime
import xlsxwriter
from typing import List

from kool_aide.library.app_setting import AppSetting
from kool_aide.library.custom_logger import CustomLogger
from kool_aide.library.constants import *
from kool_aide.library.utilities import get_cell_address, \
    get_cell_range_address, get_version

from kool_aide.model.cli_argument import CliArgument
from kool_aide.assets.resources.messages import *

class SkillsMatrixReport:

    def __init__(self, logger: CustomLogger, settings: AppSetting, 
                    data: pd.DataFrame, writer = None) -> None:
        
        self._data = data
        self._settings = settings,
        self._logger = logger

        self._writer = writer
        self._workbook = None
        self._main_header_format = None
        self._header_format_orange = None
        self._header_format_gray = None
        self._cell_wrap_noborder = None
        self._cell_total = None
        self._cell_sub_total = None
        self._footer_format = None
        self._wrap_content = None
        self._number_two_places = None

        self._month1_dict ={}
        self._month2_dict = {}
        self._employee_dict = {}

        self._log('creating component ...')

    def _log(self, message, level=3):
        self._logger.log(f"{message} [report.skills_matrix_report]", level)

    def generate(self, format: str) -> None:
        if format == OUTPUT_FORMAT[3]: # excel
            return self._generate_excel_report()
        else:
            return False, NOT_SUPPORTED

    def _generate_excel_report(self) -> None:
        try:
            self._workbook = self._writer.book
            
            self._main_header_format = self._workbook.add_format(SHEET_TOP_HEADER)
            self._sub_header_format = self._workbook.add_format(SHEET_SUB_HEADER)
            self._footer_format = self._workbook.add_format(SHEET_CELL_FOOTER)
            self._wrap_content = self._workbook.add_format(SHEET_CELL_WRAP)
            self._header_format_orange = self._workbook.add_format(SHEET_HEADER_ORANGE)
            self._header_format_gray = self._workbook.add_format(SHEET_HEADER_GRAY)
            self._report_title = self._workbook.add_format(SHEET_TITLE)
            self._header_format_lt_gray = self._workbook.add_format(SHEET_HEADER_LT_GRAY)

            self._has_training = self._workbook.add_format(SHEET_HEADER_LT_BLUE)
            self._with_support = self._workbook.add_format(SHEET_HEADER_LT_GREEN)
            self._unsupported = self._workbook.add_format(SHEET_HEADER_LT_ORANGE)
            self._sme = self._workbook.add_format(SHEET_HEADER_LT_RED)

            self._main_header_format.set_align('center')
            self._main_header_format.set_align('vcenter')

            
            self._create_skills_matrix()
            self._create_skills_pool()
               
            self._writer.save()
       
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

   
    def _create_skills_matrix(self):
        try:
            data_frame = self._data
            sheet_name = 'Skills_Matrix'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            worksheet.set_column(0,0,40,self._cell_wrap_noborder)
            worksheet.set_column(1,50,3,self._cell_wrap_noborder)
            worksheet.set_row(1, 160)  
             
            current_row += 1
            current_col = 0
            data_frame.sort_values(by=['DisplayOrder','Skill'], inplace= True)
            unique_skills = data_frame['Skill'].unique()
            
            format_vertical_text = self._sub_header_format
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            format_vertical_text.set_align('center')
            
            start_row = current_row
            skill_dict={}
            worksheet.write(current_row, current_col, "Employee", self._main_header_format)
            skill_idx = 0
            for skill in unique_skills:
                worksheet.write(
                    current_row, 
                    current_col + skill_idx +1, 
                    skill, 
                    format_vertical_text
                )

                skill_dict[skill]=skill_idx + 1
                skill_idx += 1

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(len(skill_dict),1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(0, 0, 'Skills Matrix', self._report_title)

            current_row +=1
            group_per_emp = data_frame.groupby(['Employee Name']) 
            employee_df = pd.DataFrame(group_per_emp.size().reset_index())
            employee_dict = {}
            for index, row in employee_df.iterrows():
                worksheet.write(
                    current_row,
                    0, 
                    row['Employee Name']
                ) 
                employee_dict[row['Employee Name']] = current_row
                current_row += 1 
            
            last_row= current_row
            
            # map
            for index, row in data_frame.iterrows():
                try:
                    proficiency_level = int(row['ProficiencyLevel'])
                    if proficiency_level == 1:
                        prof_format = self._has_training
                    elif proficiency_level == 2:
                        prof_format = self._with_support
                    elif proficiency_level == 3:
                        prof_format = self._unsupported
                    else:
                        prof_format = self._sme

                    worksheet.write(
                        employee_dict[row['Employee Name']],
                        skill_dict[row['Skill']],
                        row['ProficiencyLevel'],
                        prof_format
                    )

                except Exception as ex:
                    self._log(f'error creating matrix. {str(ex)}',2)

            end_row = last_row + 2
            
            worksheet.write(end_row, 0, 'LEGENDS',self._header_format_lt_gray)
            worksheet.write(end_row, 1, '1', self._has_training)
            worksheet.write(end_row, 2, 'Has Received Training', self._header_format_lt_gray)
            worksheet.write(end_row, 8, '',self._header_format_lt_gray)   
            worksheet.write(end_row, 9, '',self._header_format_lt_gray)    
            worksheet.write(end_row, 10, '2',self._with_support)
            worksheet.write(end_row, 11, 'Can Deliver Supported', self._header_format_lt_gray)
            worksheet.write(end_row, 17, '',self._header_format_lt_gray)   
            worksheet.write(end_row, 18, '',self._header_format_lt_gray)      
            worksheet.write(end_row, 19, '3',self._unsupported)
            worksheet.write(end_row, 20, 'Can Deliver Unsupported', self._header_format_lt_gray)
            worksheet.write(end_row, 27, '',self._header_format_lt_gray)
            worksheet.write(end_row, 28, '4',self._sme)
            worksheet.write(end_row, 29, 'Subject Matter Expert', self._header_format_lt_gray)
            worksheet.write(end_row, 35, '',self._header_format_lt_gray)
            end_row = end_row + 2
            worksheet.write(
                end_row, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                self._footer_format
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)

    def _create_skills_pool(self):
        try:
            data_frame = self._data
            sheet_name = 'Skill_Count'
            
            current_row = 0
            current_col = 0

            worksheet = self._writer.book.add_worksheet(sheet_name)

            worksheet.set_column(0,0,40,self._cell_wrap_noborder)
            worksheet.set_column(1,50,3,self._cell_wrap_noborder)
            worksheet.set_row(26, 160)  
            # worksheet.set_row(35, 70)    
             
            current_row += 26
            current_col = 0
            data_frame.sort_values(by=['DisplayOrder','Skill'], inplace= True)
            unique_skills = data_frame['Skill'].unique()
            # skill_df.sort_values(by=['Skill','DisplayOrder'], inplace= True)
            format_vertical_text = self._header_format_gray
            format_vertical_text.set_rotation(90)
            format_vertical_text.set_align('bottom')
            format_vertical_text.set_align('center')
            
            start_row = current_row
            skill_dict={}
            worksheet.write(current_row, current_col, "Proficiency Level", self._header_format_gray)
            skill_idx = 0
            for skill in unique_skills:
                worksheet.write(
                    current_row, 
                    current_col + skill_idx +1, 
                    skill, 
                    format_vertical_text
                )

                skill_dict[skill]=skill_idx + 1
                skill_idx += 1

            title_range = get_cell_range_address(
                get_cell_address(0,1),
                get_cell_address(len(skill_dict),1)
            )
            worksheet.merge_range(title_range,'','')

            worksheet.write(0, 0, 'Skills Count', self._report_title)

            current_row +=1
            group_per_prof = data_frame.groupby(['ProficiencyLevel','Proficiency']) 
            prof_df = pd.DataFrame(group_per_prof.size().reset_index())
            prof_dict = {}
            for index, row in prof_df.iterrows():
                worksheet.write(
                    current_row,
                    0, 
                    row['Proficiency']
                ) 
                prof_dict[row['Proficiency']] = current_row
                current_row += 1 
            
            
            
            #map
            group_by_proj_prof = data_frame.groupby(['Skill','Proficiency'])
            count_per_prof = group_by_proj_prof['Proficiency'].count()
            for key,values in count_per_prof.iteritems():
                worksheet.write(
                    prof_dict[key[1].strip()] ,
                    skill_dict[key[0].strip()],
                    values
                )  
            
            worksheet.write(current_row, 0, 'Total', self._header_format_lt_gray)

            group_by_skill = data_frame.groupby(['Skill'])
            count_per_skill = group_by_skill['Skill'].count()
            for key,values in count_per_skill.iteritems():
                worksheet.write(
                    current_row ,
                    skill_dict[key.strip()],
                    values
                )  
            current_row += 1

            last_row= current_row

            chart1 = self._writer.book.add_chart({'type':'column','subtype':'stacked'}) 
                
            categories_address = get_cell_range_address(
                get_cell_address(1,27),
                get_cell_address(len(skill_dict),27),
                sheet_name
            )

            for key,value in prof_dict.items():
                values_address = get_cell_range_address(
                    get_cell_address(1,value+1),
                    get_cell_address(len(skill_dict),value+1),
                    sheet=sheet_name
                )
                chart1.add_series({
                    'name': f'={sheet_name}!{get_cell_address(0,value+1)}',
                    'categories': f'={categories_address}',
                    'values':     f'={values_address}',
                    'data_labels' : {'value':True}
                })
            
       
            chart1.set_x_axis({'name': 'Skill'}) 
            chart1.set_y_axis({'name': 'Employee Count'}) 
            chart1.set_style(10)
            chart1.set_title({'name':'Employee Skill Count'})
            chart1.set_legend({'position': 'right'})
            worksheet.insert_chart(
                'A2',
                chart1,
                {'x_scale':1.95, 'y_scale': 1.65,'x_offset':10,'y_offset':10}
            )

            end_row = last_row + 2
            worksheet.write(
                end_row, 
                0, 
                f'Report generated : {datetime.now()} by {get_version()}', 
                self._footer_format
            )
        
        except Exception as ex:
            self._log(f'error = {str(ex)}', 2)
