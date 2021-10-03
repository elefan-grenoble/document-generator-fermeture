from wb_creator.createFDC import create_fdc
from wb_creator.create_cda import create_cda
from wb_creator.create_cairn import create_cairn
from wb_creator.createCB import create_cb
from wb_creator.createCheque import create_cheque
from os import mkdir, path
from app_config import output_folder
# Create a workbook and add a worksheet.
year = 2021
month = 11

if not path.exists(output_folder):
    print('creating output folder')
    mkdir(output_folder)

create_cda(year, month)
create_fdc(year, month)
create_cairn(year, month)
create_cb(year, month)
create_cheque(year, month)