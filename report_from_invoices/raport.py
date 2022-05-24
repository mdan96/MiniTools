import os
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles import Alignment

def check_if_raport_exist():
    for file in os.listdir(f'{current_dir}\\output'):
        if file == 'raport.xlsx':
            os.remove(f'{current_dir}\\output\\raport.xlsx')
        else:
            pass

def generate_xlsx_files_dict():
    xlsx_files_dict = {}
    for file in os.listdir(f'{current_dir}\\facturi'):
        if file.endswith('.xlsx'):
            xlsx_files_dict[file] = f'{current_dir}\\facturi\\{file}'
    return xlsx_files_dict 

def collect_data_from_xlsx(xlsx_files_dict):
    data_from_xlsx = {}
    for key in xlsx_files_dict.keys():
        factura = openpyxl.load_workbook(xlsx_files_dict[key],data_only=True)
        ws_sheet = factura.active
        numar_factura = ws_sheet['E9'].value
        data_factura = ws_sheet['E10'].value
        cumparator = ws_sheet['G3'].value
        total_lei = ws_sheet['G40'].value
        data_from_xlsx[numar_factura] = [data_factura,cumparator,total_lei]
    return data_from_xlsx

def create_raport_file(data_from_xlsx):
       
    check_if_raport_exist()
    raport_excel = openpyxl.Workbook()
    raport_sheet = raport_excel.active
    
    # GENERATE TABLE HEADER
    ft_bold = Font(bold=True)
    raport_sheet['B2'] = 'NR. FACTURA'
    raport_sheet['C2'] = 'CUMPARATOR'
    raport_sheet['D2'] = 'DATA'
    raport_sheet['E2'] = 'TOTAL' 
    
    
    # TABLE HEADER STYLES
    raport_sheet['B2'].font = ft_bold
    raport_sheet['B2'].alignment = Alignment(horizontal='center')
    raport_sheet['C2'].font = ft_bold
    raport_sheet['C2'].alignment = Alignment(horizontal='center')
    raport_sheet['D2'].font = ft_bold
    raport_sheet['D2'].alignment = Alignment(horizontal='center')
    raport_sheet['E2'].font = ft_bold
    raport_sheet['E2'].alignment = Alignment(horizontal='center')

    # POPULATE ROWS
    row = 3
    to_add_string = 0
    for key in data_from_xlsx.keys():
        # assign variables
        nr_fac = key
        data_fac = data_from_xlsx[key][0]
        cumparator_fac = data_from_xlsx[key][1]
        total_lei = data_from_xlsx[key][2]
        # populate
        raport_sheet[f'B{row}'] = nr_fac
        raport_sheet[f'C{row}'] = cumparator_fac
        raport_sheet[f'D{row}'] = data_fac
        raport_sheet[f'E{row}'] = total_lei
        # styles
        raport_sheet[f'B{row}'].alignment = Alignment(horizontal='center')
        raport_sheet[f'C{row}'].alignment = Alignment(horizontal='center')
        raport_sheet[f'D{row}'].alignment = Alignment(horizontal='center')
        raport_sheet[f'E{row}'].alignment = Alignment(horizontal='center')
        to_add_string += int(raport_sheet[f'E{row}'].value.replace(' LEI','')) 
        row += 1
    
    raport_sheet['B9'] = 'TOTAL FACTURAT'
    raport_sheet['B9'].font = ft_bold
    raport_sheet['B9'].alignment = Alignment(horizontal='center')
    raport_sheet['C9'] = f'{to_add_string} LEI'
    raport_sheet['C9'].font = ft_bold
    raport_sheet['C9'].alignment = Alignment(horizontal='center')

    for col in raport_sheet.columns:
     max_length = 0
     column = col[0].column_letter # Get the column name
     for cell in col:
         try: # Necessary to avoid error on empty cells
             if len(str(cell.value)) > max_length:
                 max_length = len(str(cell.value))
         except:
             pass
     adjusted_width = (max_length + 2) * 1.2

     raport_sheet.column_dimensions[column].width = adjusted_width

    # Save excel
    raport_excel.save(f'{current_dir}\\output\\raport.xlsx')
           
# GLOBAL VARIABLES
current_dir = os.getcwd()    
# Generate dictionaries
xlsx_files_dict = generate_xlsx_files_dict()
data_from_xlsx = collect_data_from_xlsx(xlsx_files_dict)

# Create report
create_raport_file(data_from_xlsx)