def get_col(sheet, index, col):
    sheet = f'\'{sheet}\''
    field = f'{sheet}!' + '{}' + f'{index}'
    return field.format(col)


def get_z(sheet, index):
    return get_col(sheet, index, 'Z')


def get_c(sheet, index):
    return get_col(sheet, index, 'C')


def get_f(sheet, index):
    return get_col(sheet, index, 'F')


def get_j(sheet, index):
    return get_col(sheet, index, 'J')


def get_n(sheet, index):
    return get_col(sheet, index, 'N')


def get_g(sheet, index):
    return get_col(sheet, index, 'G')


def get_r(sheet, index):
    return get_col(sheet, index, 'R')


def get_k(sheet, index):
    return get_col(sheet, index, 'K')


def get_b(sheet, index):
    return get_col(sheet, index, 'B')

def get_i(sheet, index):
    return get_col(sheet, index, 'I')

def get_d(sheet, index):
    return get_col(sheet, index, 'D')


def get_q(sheet, index):
    return get_col(sheet, index, 'Q')


def get_o(sheet, index):
    return get_col(sheet, index, 'O')


def get_p(sheet, index):
    return get_col(sheet, index, 'P')


def get_u(sheet, index):
    return get_col(sheet, index, 'U')


def get_v(sheet, index):
    return get_col(sheet, index, 'V')


def get_aa(sheet, index):
    return get_col(sheet, index, 'AA')


def get_h(sheet, index):
    return get_col(sheet, index, 'H')
