from update_sheet.helper.copy_helper import set_cell_style
from update_sheet.helper.field_helper import get_p, get_u, get_aa, get_h, get_o, get_z, get_d, get_q, get_v
from update_sheet.helper.sheetname_helper import get_dates_for_month


def update_sheet_bilan(sheet_bilan, date_sheet_names):
    cell_a6 = sheet_bilan['A6']
    cell_b6 = sheet_bilan['B6']
    cell_d6 = sheet_bilan['D6']
    cell_g6 = sheet_bilan['G6']
    cell_h6 = sheet_bilan['H6']
    offset = 6

    for index in range(offset, len(date_sheet_names) + offset):
        sheet_name = date_sheet_names[index - offset]
        sheet_bilan['A{}'.format(index)] = date_sheet_names[index - offset]
        sheet_bilan['B{}'.format(index)] = f'={get_d(sheet_name, 18)}-{get_q(sheet_name, 18)}'
        sheet_bilan['D{}'.format(index)] = f'={get_q(sheet_name, 18)}'
        sheet_bilan['G{}'.format(index)] = f'={get_u(sheet_name, 18)}'
        sheet_bilan['H{}'.format(index)] = f'={get_v(sheet_name, 18)}'
        set_cell_style(cell_a6, sheet_bilan[f'A{index}'])
        set_cell_style(cell_b6, sheet_bilan[f'B{index}'])
        set_cell_style(cell_d6, sheet_bilan[f'D{index}'])
        set_cell_style(cell_g6, sheet_bilan[f'G{index}'])
        set_cell_style(cell_h6, sheet_bilan[f'H{index}'])

    value_b3 = '='
    for day in date_sheet_names:
        value_b3 += f'{get_d(day, 18)}-{get_q(day, 18)}+'
    value_b3 = value_b3[0:len(value_b3) - 1]
    sheet_bilan['B3'] = value_b3
    sheet_bilan['B4'] = f'=IF(B3=SUM(B6:B{len(date_sheet_names) + offset}),"Valide","PB")'

    first_day = date_sheet_names[0]
    value_k4 = '=\'FIN DE MOIS\'!R18+\'FIN DE MOIS\'!S18+\'FIN DE MOIS\'!G18-' \
               f'{get_o(first_day, 18)}-{get_p(first_day, 18)}-{get_q(first_day, 18)}'
    sheet_bilan['K4'] = value_k4


def update_sheet_fin_de_mois(sheet, year, month):
    title_sheet_before = get_dates_for_month(year, month)[-1]
    for index in range(3, 8):
        sheet[f'E{index}'] = f'={get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)}'
        sheet[f'F{index}'] = f'=IF({get_h(title_sheet_before, index)}' \
                             f'>0,' \
                             f'{get_p(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_aa(title_sheet_before, index)},' \
                             f'{get_p(title_sheet_before, index)}-{get_h(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_aa(title_sheet_before, index)})'
        sheet[f'G{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_q(title_sheet_before, index)},' \
                             f'IF(({get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)})' \
                             f'>={get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_o(title_sheet_before, index)})'

    for index in range(8, 17):
        sheet[f'E{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_o(title_sheet_before, index)},' \
                             f'IF({get_h(title_sheet_before, index)}>{get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)},' \
                             f'0,' \
                             f'{get_o(title_sheet_before, index)}-{get_h(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_z(title_sheet_before, index)}))'
        sheet[f'F{index}'] = f'={get_p(title_sheet_before, index)}+{get_aa(title_sheet_before, index)}'
        sheet[f'G{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_q(title_sheet_before, index)},' \
                             f'IF(({get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)})' \
                             f'>={get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_o(title_sheet_before, index)})'


def update_ouverture(sheet, title_sheet_before):
    for index in range(3, 8):
        sheet[f'O{index}'] = f'={get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)}'
        sheet[f'P{index}'] = f'=IF({get_h(title_sheet_before, index)}' \
                             f'>0,' \
                             f'{get_p(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_aa(title_sheet_before, index)},' \
                             f'{get_p(title_sheet_before, index)}-{get_h(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_aa(title_sheet_before, index)})'
        sheet[f'Q{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_q(title_sheet_before, index)},' \
                             f'IF(({get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)})' \
                             f'>={get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_o(title_sheet_before, index)}))'

    for index in range(8, 17):
        sheet[f'O{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_o(title_sheet_before, index)},' \
                             f'IF({get_h(title_sheet_before, index)}>{get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)},' \
                             f'0,' \
                             f'{get_o(title_sheet_before, index)}-{get_h(title_sheet_before, index)}-{get_u(title_sheet_before, index)}+{get_z(title_sheet_before, index)}))'
        sheet[f'P{index}'] = f'={get_p(title_sheet_before, index)}+{get_aa(title_sheet_before, index)}'
        sheet[f'Q{index}'] = f'=IF({get_h(title_sheet_before, index)}="",' \
                             f'{get_q(title_sheet_before, index)},' \
                             f'IF(({get_o(title_sheet_before, index)}+{get_z(title_sheet_before, index)})' \
                             f'>={get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_h(title_sheet_before, index)},' \
                             f'{get_d(title_sheet_before, index)}+{get_o(title_sheet_before, index)}))'
