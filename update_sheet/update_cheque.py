from update_sheet.helper.copy_helper import set_cell_style
from update_sheet.helper.sheetname_helper import get_dates_for_month
from update_sheet.helper.field_helper import get_i, get_b


def update_bilan(sheet_bilan, year, month):
    dates = get_dates_for_month(year, month)
    field_a4 = sheet_bilan['A4']
    field_b4 = sheet_bilan['B4']
    field_d4 = sheet_bilan['D4']
    field_e4 = sheet_bilan['E4']

    offset = 6
    total_adhesion = '='
    total_ventes = '='
    for index, day in enumerate(dates):
        sheet_bilan[f'A{index + offset}'] = day
        set_cell_style(field_a4, sheet_bilan[f'A{index + offset}'])

        sheet_bilan[f'B{index + offset}'] = f'={get_b(day, 19)}'
        set_cell_style(field_b4, sheet_bilan[f'B{index + offset}'])

        sheet_bilan[f'D{index + offset}'] = day
        set_cell_style(field_d4, sheet_bilan[f'D{index + offset}'])

        sheet_bilan[f'E{index + offset}'] = f'={get_i(day, 19)}'
        set_cell_style(field_e4, sheet_bilan[f'E{index + offset}'])

        total_adhesion += f'{get_b(day, 19)}+'
        total_ventes += f'{get_i(day, 19)}+'

    sheet_bilan['B3'] = total_adhesion[:-1]
    sheet_bilan['B4'] = f'=IF(B3=SUM(B6:B{len(dates) + offset}),"Valide","PB")'

    sheet_bilan['E3'] = total_ventes[:-1]
    sheet_bilan['E4'] = f'=IF(E3=SUM(E6:E{len(dates) + offset}),"Valide","PB")'
