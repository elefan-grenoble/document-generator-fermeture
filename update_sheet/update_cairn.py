from update_sheet.helper.copy_helper import set_cell_style
from update_sheet.helper.field_helper import get_c, get_f, get_j, get_n, get_r
from update_sheet.helper.sheetname_helper import get_dates_for_month


def update_bilan(sheet_bilan, year, month):
    dates = get_dates_for_month(year, month)

    field_b4 = sheet_bilan['B4']

    offset = 6
    total_adhesions = '='
    total_ventes = '='
    total_cash = '='
    total_e_cash = '='
    for index, day in enumerate(dates):
        sheet_bilan[f'A{index + offset}'] = day
        set_cell_style(sheet_bilan['A4'], sheet_bilan[f'A{index + offset}'])
        sheet_bilan[f'B{index + offset}'] = f'={get_f(day, 10)}'
        set_cell_style(field_b4, sheet_bilan[f'B{index + offset}'])

        sheet_bilan[f'D{index + offset}'] = day
        set_cell_style(sheet_bilan['D4'], sheet_bilan[f'D{index + offset}'])
        sheet_bilan[f'E{index + offset}'] = f'={get_j(day, 10)}'
        set_cell_style(field_b4, sheet_bilan[f'E{index + offset}'])

        sheet_bilan[f'G{index + offset}'] = day
        set_cell_style(sheet_bilan['G4'], sheet_bilan[f'G{index + offset}'])
        sheet_bilan[f'H{index + offset}'] = f'={get_n(day, 10)}'
        set_cell_style(field_b4, sheet_bilan[f'H{index + offset}'])

        sheet_bilan[f'J{index + offset}'] = day
        set_cell_style(sheet_bilan['J4'], sheet_bilan[f'J{index + offset}'])
        sheet_bilan[f'K{index + offset}'] = f'={get_r(day, 10)}'
        set_cell_style(field_b4, sheet_bilan[f'K{index + offset}'])

        total_adhesions += f'{get_f(day, 10)}+'
        total_ventes += f'{get_j(day, 10)}+'
        total_cash += f'{get_n(day, 10)}+'
        total_e_cash += f'{get_r(day, 10)}+'

    sheet_bilan['B3'] = total_adhesions[:-1]
    sheet_bilan['E3'] = total_ventes[:-1]
    sheet_bilan['H3'] = total_cash[:-1]
    sheet_bilan['K3'] = total_e_cash[:-1]
    sheet_bilan['B4'] = f'=IF(K3=SUM(B6:B{len(dates) + offset}),"Valide","PB")'
    sheet_bilan['E4'] = f'=IF(K3=SUM(E6:E{len(dates) + offset}),"Valide","PB")'
    sheet_bilan['H4'] = f'=IF(K3=SUM(H6:H{len(dates) + offset}),"Valide","PB")'
    sheet_bilan['K4'] = f'=IF(K3=SUM(K6:K{len(dates) + offset}),"Valide","PB")'


def update_fin_de_mois(sheet, year, month):
    dates = get_dates_for_month(year, month)
    title_last_sheet = dates[-1]
    for index in range(2, 8):
        sheet[f'C{index}'] = f'={get_c(title_last_sheet, index)}' \
                             f'+{get_f(title_last_sheet, index)}' \
                             f'+{get_j(title_last_sheet, index)}' \
                             f'-{get_n(title_last_sheet, index)}' \
                             f'-{get_r(title_last_sheet, index)}'
        adhesions = '='
        ventes = '='
        factures = '='
        for day in dates:
            adhesions += f'{get_f(day, index)}+'
            ventes += f'{get_j(day, index)}+'
            factures += f'{get_n(day, index)}+'
        sheet[f'F{index}'] = adhesions[:-1]
        sheet[f'J{index}'] = ventes[:-1]
        sheet[f'Q{index}'] = factures[:-1]


def update_reserve(sheet, title_sheet_before):
    for index in range(2, 8):
        sheet[f'C{index}'] = f'={get_c(title_sheet_before, index)}' \
                             f'+{get_f(title_sheet_before, index)}' \
                             f'+{get_j(title_sheet_before, index)}' \
                             f'-{get_n(title_sheet_before, index)}' \
                             f'-{get_r(title_sheet_before, index)}'
