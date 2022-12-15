import datetime
from dateutil.relativedelta import relativedelta
import os
from os import mkdir, path

from update_sheet.helper.sheetname_helper import months
from wb_creator.createCB import create_cb
from wb_creator.createCheque import create_cheque
from wb_creator.createFDC import create_fdc
from wb_creator.create_cairn import create_cairn
from wb_creator.create_cda import create_cda


next_month = datetime.datetime.now() + relativedelta(months=1)
year = next_month.year
month = next_month.month

if 'YEAR' not in os.environ or 'MONTH' not in os.environ:
    print(f'Variable YEAR or MONTH not set. Generating documents for next month [{months[month-1]} {year}]')
else:
    env_year = os.environ['YEAR']
    env_month = os.environ['MONTH']
    try:
        year = int(env_year)
        month = int(env_month)
    except ValueError:
        print(f'YEAR or MONTH environment variable can not be cast to int [{env_year}, {env_month}]')
        exit(1)

month_name = months[month-1]
print(f'Creating files for: {month_name} {year}')

output_folder = f'output'
if not path.exists(output_folder):
    print(f'creating output folder[{output_folder}]')
    mkdir(output_folder)
output_folder = f'{output_folder}/{year} - Fichiers de fermeture de caisse/'
if not path.exists(output_folder):
    print(f'creating output folder[{output_folder}]')
    mkdir(output_folder)

year_month = datetime.date(year, month, 1).strftime("%Y.%m")
output_folder = f'{output_folder}{year_month} - {month_name} {year}/'
if not path.exists(output_folder):
    print(f'creating output folder for month[{output_folder}]')
    mkdir(output_folder)


create_cda(year, month, output_folder)
create_fdc(year, month, output_folder)
create_cairn(year, month, output_folder)
create_cb(year, month, output_folder)
create_cheque(year, month, output_folder)
