import os

from wb_creator.createFDC import create_fdc
from wb_creator.create_cda import create_cda
from wb_creator.create_cairn import create_cairn
from wb_creator.createCB import create_cb
from wb_creator.createCheque import create_cheque
from os import mkdir, path
import locale
import datetime


if 'YEAR' not in  os.environ or 'MONTH' not in os.environ:
    raise Exception('YEAR and MONTH environment variables must be set')

env_year = os.environ['YEAR']
env_month = os.environ['MONTH']
try:
    year = int(env_year)
    month = int(env_month)
except ValueError:
    print(f'YEAR or MONTH environment variable can not be cast to int [{env_year}, {env_month}]')

locale.setlocale(category=locale.LC_ALL, locale='fr_FR.utf8')

month_name = datetime.date(year,month,1).strftime("%B").title()
print(f'Creating files for: {month_name} {year}')

output_folder = 'output/'
if not path.exists(output_folder):
    print(f'creating output folder[{output_folder}]')
    mkdir(output_folder)
year_month = datetime.date(year,month,1).strftime("%Y.%m")
output_folder = f'{output_folder}{year_month} - {month_name} {year}/'
if not path.exists(output_folder):
    print(f'creating output folder for month[{output_folder}]')
    mkdir(output_folder)


create_cda(year, month, output_folder)
create_fdc(year, month, output_folder)
create_cairn(year, month, output_folder)
create_cb(year, month, output_folder)
create_cheque(year, month, output_folder)