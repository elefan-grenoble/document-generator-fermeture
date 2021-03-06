from update_sheet.helper.copy_helper import set_cell_style
from update_sheet.helper.field_helper import get_d, get_i, get_h, get_f
from update_sheet.helper.sheetname_helper import get_dates_for_month


def update_bilan(sheet_bilan, year, month):
    dates = get_dates_for_month(year, month)
    field_a4 = sheet_bilan['A4']
    field_b4 = sheet_bilan['B4']
    field_d6 = sheet_bilan['D6']
    field_e6 = sheet_bilan['E6']

    offset = 6
    total = '='
    for index, day in enumerate(dates):
        sheet_bilan[f'A{index+offset}'] = day
        set_cell_style(field_a4, sheet_bilan[f'A{index + offset}'])

        sheet_bilan[f'B{index + offset}'] = f'={get_i(day,11)}  +{get_h(day,11)}'
        set_cell_style(field_b4, sheet_bilan[f'B{index + offset}'])

        sheet_bilan[f'D{index + offset}'] = f'={get_d(day, 5)}+{get_f(day, 5)}+{get_d(day, 16)}+{get_f(day, 16)}'
        set_cell_style(field_d6, sheet_bilan[f'D{index + offset}'])

        sheet_bilan[f'E{index + offset}'] = f'=B{index + offset}-D{index + offset}'
        set_cell_style(field_e6, sheet_bilan[f'E{index + offset}'])

        total += f'{get_i(day,11)}+{get_h(day,11)}+'

    sheet_bilan['B3'] = total[:-1]
    sheet_bilan['B4'] = f'=IF(B3=SUM(B6:B{len(dates)+offset}),"Valide","PB")'

