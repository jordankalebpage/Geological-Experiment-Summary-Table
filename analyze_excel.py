import glob
import logging
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

import numpy as np
import pandas as pd
import xlrd

# Hard-coded locations for convenience
# sum_tbl_loc = "C:\\Users\\jordan.page\\Downloads\\Summary Table.xls"
# experiment_folder_path = "C:\\Users\\jordan.page\\Documents\\Geo Stuff\\Second\\"


def analyze_excel() -> None:
    """
    * Ensure that the cells under the "Summary Table" column headers are completely empty.
    * I.e., make sure the Permeability (cm/sec @ 20C) column doesn't have those "-", and the
      LAB ID column doesn't have the "Type Lab ID here Cell A6" comment in it.
    """
    Tk().withdraw()
    sum_tbl_loc = askopenfilename(title="Select the Summary Table")
    experiment_folder_path = askdirectory(title="Select the Folder containing the Experiments") + '/'
    if sum_tbl_loc is None or experiment_folder_path == '/':
        logging.debug('Either the Summary Table or Experiment Folder was not selected.')
        exit(1)

    logging.debug(f'Using Summary Table at: {sum_tbl_loc}')
    logging.debug(f'Using Experiment Folder at: {experiment_folder_path}')

    project_num = pd.read_excel(sum_tbl_loc, nrows=3).values[1][2]
    logging.debug(f'Project Number: {project_num}')

    files = glob.glob(experiment_folder_path + project_num + '*')

    only_nums = [int(i.split('-')[-1].split(' ')[0]) for i in files]
    types = [' '.join(i.split('-')[-1].split(' ')[1:]).split('.')[0] for i in files]
    individual_nums = np.unique(only_nums)
    num_ids = len(individual_nums)
    logging.debug(f'Individual IDs: {individual_nums}')
    logging.debug(f'Number of individual IDs: {num_ids}')

    info_dict = {}
    logging.debug('Generating NA dict...')
    for i, v in enumerate(only_nums):
        info_dict[v] = info_dict.get(v, {'Data': {'Boring': 'NA', 'Depth': 'NA', 'Sample Number': 'NA',
                                                  'Water Content %': 'NA', 'Liquid Limit %': 'NA',
                                                  'Plastic Limit %': 'NA', 'Plasticity Index %': 'NA',
                                                  'USCS Symbol (Limits)': 'NA', 'Passing #200 %': 'NA',
                                                  'Passing 0.002 mm %': 'NA', 'USCS Symbol (Grain Size)': 'NA',
                                                  'USCS Classification': 'NA', 'Standard Proctor OMC %': 'NA',
                                                  'Standard Proctor MDD (pcf)': 'NA', 'Average Specific Gravity': 'NA',
                                                  'Permeability (cm/sec @ 20C)': 'NA'}})
        if info_dict.get(v).get('Experiments') is None:
            info_dict[v]['Experiments'] = [{'Experiment Type': types[i], 'File Location': files[i]}]
        else:
            info_dict[v]['Experiments'].append({'Experiment Type': types[i], 'File Location': files[i]})

    sum_tbl = pd.read_excel(sum_tbl_loc, skiprows=4, nrows=num_ids)

    logging.debug('Reading Excel files...')
    for lab_id in info_dict.keys():
        for experiment in info_dict.get(lab_id).get('Experiments'):
            experiment_type = experiment.get('Experiment Type')
            file_location = experiment.get('File Location')

            if 'Grain Sieve' in experiment_type:
                cell_values = {'Boring': (3, 7), 'Depth': (4, 7), 'Sample Number': (5, 7), 'Passing #200 %': (44, 5),
                               'Passing 0.002 mm %': (106, 1), 'USCS Symbol (Grain Size)': (49, 0),
                               'USCS Classification': (52, 0)}
                sheet = xlrd.open_workbook_xls(file_location).sheet_by_name("Sieve Hyd")

                info_dict[lab_id]['Data']['Boring'] = sheet.cell(*cell_values.get('Boring')).value
                info_dict[lab_id]['Data']['Depth'] = sheet.cell(*cell_values.get('Depth')).value
                info_dict[lab_id]['Data']['Sample Number'] = sheet.cell(*cell_values.get('Sample Number')).value
                info_dict[lab_id]['Data']['Passing #200 %'] = round(
                    float(sheet.cell(*cell_values.get('Passing #200 %')).value), 2)
                info_dict[lab_id]['Data']['Passing 0.002 mm %'] = round(
                    float(sheet.cell(*cell_values.get('Passing 0.002 mm %')).value), 2)
                info_dict[lab_id]['Data']['USCS Symbol (Grain Size)'] = sheet.cell(
                    *cell_values.get('USCS Symbol (Grain Size)')).value
                info_dict[lab_id]['Data']['USCS Classification'] = sheet.cell(
                    *cell_values.get('USCS Classification')).value
            elif 'Limit' in experiment_type or 'limit' in experiment_type:
                cell_values = {'Water Content %': (21, 3), 'Liquid Limit %': (26, 9), 'Plastic Limit %': (28, 9),
                               'Plasticity Index %': (30, 9),
                               'USCS Symbol (Limits)': (32, 9)}
                sheet = xlrd.open_workbook_xls(file_location).sheet_by_name("Sheet1")

                info_dict[lab_id]['Data']['Water Content %'] = round(
                    float(sheet.cell(*cell_values.get('Water Content %')).value), 1)
                info_dict[lab_id]['Data']['Liquid Limit %'] = int(sheet.cell(*cell_values.get('Liquid Limit %')).value)
                info_dict[lab_id]['Data']['Plastic Limit %'] = int(
                    sheet.cell(*cell_values.get('Plastic Limit %')).value)
                info_dict[lab_id]['Data']['Plasticity Index %'] = int(
                    sheet.cell(*cell_values.get('Plasticity Index %')).value)
                info_dict[lab_id]['Data']['USCS Symbol (Limits)'] = sheet.cell(
                    *cell_values.get('USCS Symbol (Limits)')).value
            elif 'Proctor Std' in experiment_type:
                cell_values = {'Standard Proctor OMC %': (13, 6), 'Standard Proctor MDD (pcf)': (14, 6)}
                sheet = xlrd.open_workbook_xls(file_location).sheet_by_name("Sheet1")

                info_dict[lab_id]['Data']['Standard Proctor OMC %'] = round(
                    float(sheet.cell(*cell_values.get('Standard Proctor OMC %')).value), 1)
                info_dict[lab_id]['Data']['Standard Proctor MDD (pcf)'] = round(
                    float(sheet.cell(*cell_values.get('Standard Proctor MDD (pcf)')).value), 1)
            elif 'Specific Gravity' in experiment_type:
                cell_values = {'Average Specific Gravity': (36, 8)}
                sheet = xlrd.open_workbook_xls(file_location).sheet_by_name("Sheet1")

                info_dict[lab_id]['Data']['Average Specific Gravity'] = round(
                    float(sheet.cell(*cell_values.get('Average Specific Gravity')).value), 2)
            elif 't Perm' in experiment_type:
                cell_values = {'Permeability (cm/sec @ 20C)': (12, 15)}
                sheet = xlrd.open_workbook_xls(file_location).sheet_by_name("Raw Data")

                info_dict[lab_id]['Data']['Permeability (cm/sec @ 20C)'] = sheet.cell(
                    *cell_values.get('Permeability (cm/sec @ 20C)')).value
            else:
                logging.debug(f'Unrecognized experiment type: {experiment_type}')

    logging.debug('Grabbing all Data lists...')
    all_lab_ids = [project_num + '-' + str(i).zfill(3) for i in individual_nums]
    all_borings = [v['Data']['Boring'] for k, v in info_dict.items()]
    all_depths = [v['Data']['Depth'] for k, v in info_dict.items()]
    all_sample_numbers = [v['Data']['Sample Number'] for k, v in info_dict.items()]
    all_water_contents = [v['Data']['Water Content %'] for k, v in info_dict.items()]
    all_liquid_limits = [v['Data']['Liquid Limit %'] for k, v in info_dict.items()]
    all_plastic_limits = [v['Data']['Plastic Limit %'] for k, v in info_dict.items()]
    all_plasticity_indices = [v['Data']['Plasticity Index %'] for k, v in info_dict.items()]
    all_uscs_symbol_limits = [v['Data']['USCS Symbol (Limits)'] for k, v in info_dict.items()]
    all_passing_200 = [v['Data']['Passing #200 %'] for k, v in info_dict.items()]
    all_passing_002 = [v['Data']['Passing 0.002 mm %'] for k, v in info_dict.items()]
    all_uscs_symbol_grain_sizes = [v['Data']['USCS Symbol (Grain Size)'] for k, v in info_dict.items()]
    all_uscs_classifications = [v['Data']['USCS Classification'] for k, v in info_dict.items()]
    all_standard_proctor_omc = [v['Data']['Standard Proctor OMC %'] for k, v in info_dict.items()]
    all_standard_proctor_mdd = [v['Data']['Standard Proctor MDD (pcf)'] for k, v in info_dict.items()]
    all_average_specific_gravities = [v['Data']['Average Specific Gravity'] for k, v in info_dict.items()]
    all_permeabilities = [v['Data']['Permeability (cm/sec @ 20C)'] for k, v in info_dict.items()]

    logging.debug('Vectorizing the Data lists into the Summary Table...')
    sum_tbl['LAB ID'] = all_lab_ids
    sum_tbl['Boring'] = all_borings
    sum_tbl['Depth'] = all_depths
    sum_tbl['Sample Number'] = all_sample_numbers
    sum_tbl['Water Content %'] = all_water_contents
    sum_tbl['Liquid Limit           %'] = all_liquid_limits
    sum_tbl['Plastic Limit               %'] = all_plastic_limits
    sum_tbl['Plasticity Index                    %'] = all_plasticity_indices
    sum_tbl['USCS Symbol (Limits)'] = all_uscs_symbol_limits
    sum_tbl['Passing #200              %'] = all_passing_200
    sum_tbl['Passing 0.002 mm %'] = all_passing_002
    sum_tbl['USCS Symbol (Grain Size)'] = all_uscs_symbol_grain_sizes
    sum_tbl['USCS Classification'] = all_uscs_classifications
    sum_tbl['Standard Proctor OMC      %'] = all_standard_proctor_omc
    sum_tbl['Standard Proctor MDD     (pcf)'] = all_standard_proctor_mdd
    sum_tbl['Average Specific Gravity'] = all_average_specific_gravities
    sum_tbl['Permeability  (cm/sec            @ 20oC)'] = all_permeabilities

    output_file_name = 'Summary Table ' + project_num + '.xlsx'

    try:
        logging.debug(f'Outputting Summary Table to file: {output_file_name}')
        sum_tbl.to_excel(output_file_name, sheet_name='Summary Table', index=False)
        logging.debug(f'Info dict: {info_dict}')
    except PermissionError:
        logging.debug(f'You need to close {output_file_name}, then run this utility again.')


if __name__ == '__main__':
    logging.basicConfig(filename='../Geo Stuff.log', encoding='utf-8', level=logging.DEBUG,
                        format='%(asctime)s | %(levelname)s | %(message)s',
                        datefmt='%m/%d/%Y %I:%M:%S %p')
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))

    analyze_excel()
