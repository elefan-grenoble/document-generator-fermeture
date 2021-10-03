from update_sheet.helper.copy_helper import set_cell_style
from update_sheet.helper.field_helper import get_d, get_k, get_n
from update_sheet.helper.sheetname_helper import get_dates_for_month

def update_sheet_bilan(sheet_bilan, date_sheet_names):
    cell_a6 = sheet_bilan['A6']
    cell_b6 = sheet_bilan['B6']
    offset = 6

    for index in range(offset, len(date_sheet_names) + offset):
        sheet_bilan['A{}'.format(index)] = date_sheet_names[index - offset]
        sheet_bilan['B{}'.format(index)] = f'={get_d(date_sheet_names[index - offset], 18)}'

        set_cell_style(cell_a6, sheet_bilan['A{}'.format(index)])
        set_cell_style(cell_b6, sheet_bilan['B{}'.format(index)])

    value_b3 = '='
    for day in date_sheet_names:
        value_b3 += f'{get_d(day, 18)}+'
    value_b3 = value_b3[:-1]
    sheet_bilan['B3'] = value_b3
    sheet_bilan['B4'] = f'=IF(B3=SUM(B6:B{len(date_sheet_names) + offset}),"Valide","PB")'


def update_sheet_fin_de_mois(sheet, year, month):
    last_date = get_dates_for_month(year, month)[-1]
    for index in range(3, 17):
        k_field = f'={get_k(last_date, index)}+{get_d(last_date, index)}-{get_n(last_date, index)}'
        sheet[f'C{index}'] = k_field


def update_reserve(sheet, title_sheet_before):
    for index in range(3, 17):
        k_field = f'={get_k(title_sheet_before, index)}+{get_d(title_sheet_before, index)}-{get_n(title_sheet_before, index)}'
        sheet[f'K{index}'] = k_field
