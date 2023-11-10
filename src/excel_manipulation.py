import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import TableStyleInfo, Table
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, DEFAULT_FONT


class ExcelManipulator:
    def __init__(self):
                        
        self.column_lists = {
            
            'empty_inserted_columns': ['Attempt 1',
                                       'Attempt 2',
                                       'Current Week Client Comments',
                                       'Age',
                                       'Prior Week Client Comments',
                                       'RAI Reconciliation Comments',
                                       'Status'],
                        
            'left_columns'          : ['Client',
                                       'DOS',
                                       'Account #',
                                       'MRN', 
                                       'Patient Name',
                                       'Carrier',
                                       'Department'],
            
            'singe_uac'             : ['UAC Reason - Provider(DOS)',
                                       'UAC Reason'],
            
            'multiple_uac'          : ['UAC Reason 1 - Provider(DOS)',
                                       'UAC Reason 1',
                                       'UAC Reason 2 - Provider(DOS)',
                                       'UAC Reason 2'],
            
            'right_columns'         : ['Attempt 1',
                                       'Attempt 2',
                                       'Current Week Client Comments',
                                       'Age',
                                       'Pro Date Sent To Client',
                                       'Prior Week Client Comments',
                                       'RAI Reconciliation Comments',
                                       'Status'],                    
        }
        

    def pandas_column_clean(self, chartsearch_path, toggled_states):
        
        excel_chartsearch = pd.read_excel(chartsearch_path)
        toggled_attempt = toggled_states["toggle_attempt"]
        toggled_search_list = toggled_states["toggle_search_list"]
        
        beginning_col_names = excel_chartsearch.columns.tolist()
        
        if toggled_attempt is False:
            self.column_lists['empty_inserted_columns'].remove('Attempt 1')
            self.column_lists['empty_inserted_columns'].remove('Attempt 2')
            self.column_lists['right_columns'].remove('Attempt 1')
            self.column_lists['right_columns'].remove('Attempt 2')

        if toggled_search_list is False:
            self.column_lists['empty_inserted_columns'].remove('Status')
            self.column_lists['right_columns'].remove('Status')
               
        if 'Department' not in beginning_col_names:
            left_columns_chartsearch = self.column_lists['left_columns'].remove('Department')
        
        if 'Pro Date Sent To Client' not in beginning_col_names:
            right_columns_chartsearch = self.column_lists['right_columns'].remove('Pro Date Sent To Client') 

        if 'Client' not in beginning_col_names:
            left_columns_chartsearch = self.column_lists['left_columns'].remove('Client') 
                      
        if 'UAC Reason 1' and 'UAC Reason 2' in beginning_col_names:
            middle_columns = self.column_lists['multiple_uac']
        elif 'UAC Reason - Provider(DOS)' in beginning_col_names:
            middle_columns = self.column_lists['singe_uac']

        for col in self.column_lists['empty_inserted_columns']:
                excel_chartsearch[col] = ""
        
        
        # Assign Correct Order
        left_columns_chartsearch = excel_chartsearch[self.column_lists['left_columns']]
        middle_columns_chartsearch = excel_chartsearch[middle_columns]          
        right_columns_chartsearch = excel_chartsearch[self.column_lists['right_columns']]

        concatenated_excel_file = pd.concat([left_columns_chartsearch, middle_columns_chartsearch, right_columns_chartsearch], axis=1)
                        
        return concatenated_excel_file

    def openpyxl_format_workbook(self, concatenated_excel_file, alphabet): 

        wb = Workbook()
        worksheet = wb.active
        worksheet.title = "RAI Report"

        for row in dataframe_to_rows(concatenated_excel_file, index=False, header=True):
            worksheet.append(row)

        col_names = []

        for col in concatenated_excel_file.columns:
            col_names.append(col)


        col_num = len(concatenated_excel_file.axes[1])
        col_num_str = str(col_num)
        row_num = len(concatenated_excel_file.axes[0])
        row_num_str = str(row_num + 1)
        num_records = int(row_num_str) - 1

        col_to_letter = alphabet.get(col_num_str)

        table_dimension = str("A1:" + col_to_letter + row_num_str)

        excel_dimensions = {
                            'column_names'    : col_names,
                            'num_of_columns'  : col_num,
                            'num_of_rows'     : num_records,
                            'table_dimensions': table_dimension
                            }
        
        age_index = col_names.index('Age')
            
        age_index = str(age_index + 1)
        
            
        col_to_letter_age = alphabet.get(age_index)
         
        age_range = 2 
                    
        for row_num in range(age_range, int(row_num_str) + 1):
            worksheet[col_to_letter_age + '{}'.format(row_num)] = '=datedif(a{},today(),"D")'.format(str(age_range))  # noqa: E501
            age_range += 1
            
        col_to_letter = alphabet.get(str(col_num))
        
        table = Table(displayName = "table", ref = table_dimension)
        
            
        # Change table style to normal format
        style = TableStyleInfo(name = "TableStyleMedium2", showRowStripes = True)
            
        # Attatched the styles to table
        table.tableStyleInfo = style
        if 'Pro Date Sent To Client' in col_names:
            date_index = col_names.index('Pro Date Sent To Client')
            date_index = str(date_index + 1)
            col_to_letter_date = alphabet.get(date_index)

            for cell in worksheet[col_to_letter_date]:
                cell.alignment = Alignment(horizontal='center')  
                    
        for cell in worksheet[col_to_letter_age]:
            cell.alignment = Alignment(horizontal='center') 
                
        for cell in worksheet['A']:
            cell.alignment = Alignment(horizontal='center')  
            
        for cell in worksheet['B']:
            cell.alignment = Alignment(horizontal='center')  
            
        for cell in worksheet['C']:
            cell.alignment = Alignment(horizontal='left') 
                             
        # Attach table to worksheet
        worksheet.add_table(table)
            
        DEFAULT_FONT.size = 8
            
        return wb, excel_dimensions
