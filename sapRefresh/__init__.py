# -*- coding: utf-8 -*-
"""
Created on 3/14/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import pathlib

import pandas as pd

from Core.Time import timeit

CONFIG_PATH = r'\\branapv-sql01\DIGITAL_TRANSFORMATION\Python\Config\config.xlsx'


@timeit
def import_global_configurations(config_location):
    """import all the necessary information so that the script can work well"""
    df_global_configs = pd.read_excel(config_location, sheet_name='global_configs')
    path_log = pathlib.Path(df_global_configs.query('description=="path-log_directory"')['value'].values[0])
    return df_global_configs, path_log


# import global configurations
global_configs_df, LOG_PATH = import_global_configurations(CONFIG_PATH)
