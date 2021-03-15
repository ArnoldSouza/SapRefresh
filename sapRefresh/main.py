# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import os
import sys

import pandas as pd   # Library for data manipulation

current_path = os.path.abspath('.')
parent_path = os.path.dirname(current_path)
sys.path.append(parent_path)

from sapRefresh.Core import Connection as Conn
from sapRefresh.Core.Time import SpinnerCursor
from sapRefresh.Refresh import Engine as Xl
from sapRefresh.Refresh import Sap

from sapRefresh.Core.base_logger import get_logger
logger = get_logger(__name__)


def main():
    # kill every excel instance so that further erros cannot happen
    Xl.kill_excel_instances()

    # get configuration values
    config_values = Conn.get_config_values()

    # check if the application is running inside Hydro
    Conn.check_connection(config_values['app_server'], config_values['app_port'])

    # initiate Excel instance
    ExcelInstance = Xl.open_excel()

    # Configure the Excel Instance to optimize the execution
    Xl.optimize_instance(ExcelInstance, 'start')

    # open the target workbook
    WorkbookSAP = Xl.open_workbook(ExcelInstance, config_values['filepath'])

    # call method to activate SAP AfO AddIn
    Xl.ensure_addin(ExcelInstance)

    # capture the initial calculation state
    calc_state_init = Xl.calculation_state(ExcelInstance, 'start')

    # activate workbook
    Xl.ensure_wb_active(ExcelInstance, config_values['filename'])

    # initial calculation
    ExcelInstance.Application.Calculate()

    data_source = Xl.get_data_source(ExcelInstance)

    result_logon = Sap.sap_logon(ExcelInstance,
                                 data_source['DS'],
                                 config_values['bwclient'],
                                 config_values['bwuser'],
                                 config_values['bwpassword'])

    result_refresh = Sap.sap_refresh(ExcelInstance)

    data_source = Sap.sap_get_more_info(ExcelInstance, data_source)

    # ---------------------------------------------------------------------------
    #   Tests related to variables, filters and dimensions
    # ---------------------------------------------------------------------------

    list_variables = ExcelInstance.Application.Run("SAPListOfVariables", data_source['DS'], "INPUT_STRING", "PROMPTS")

    # need to create a iteration between variables and technical names
    for variable in list_variables:
        print(variable[0], '\t', variable[1], '\t',
              ExcelInstance.Application.Run("SAPGetVariable", data_source['DS'], variable[0], "TECHNICALNAME"))

    arrFilters = ExcelInstance.Application.Run("SAPListOfDynamicFilters", data_source['DS'], "INPUT_STRING")
    arrDimensions = ExcelInstance.Application.Run("SAPListOfDimensions", data_source['DS'])

    # test to check what this func returns
    test_arnold = ExcelInstance.Application.Run("SAPListOfEffectiveFilters", data_source['DS'], "INPUT_STRING")

    for value in arrFilters:
        if value != 'Measures':
            print(value)

    print('\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n')

    for value in arrDimensions:
        print(value)

    print('\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n')

    for value in test_arnold:
        print(value)

    # ---------------------------------------------------------------------------
    #   End of tests
    # ---------------------------------------------------------------------------

    result_refresh_data = Sap.sap_refresh_data(ExcelInstance, data_source['DS'])

    # put calculation back to original state
    Xl.calculation_state(ExcelInstance, 'stop', calc_state_init)

    # terminate the workbook instance
    Xl.close_workbook(WorkbookSAP)

    # Configure the Excel Instance to optimize the execution
    Xl.optimize_instance(ExcelInstance, 'stop')

    # Close Excel Instance
    ExcelInstance.Application.Quit()
    del ExcelInstance

    # ---------------------------------------------------------------------------
    #   End of refresh operation
    # ---------------------------------------------------------------------------

    # start waiting spinner
    print('\nStarting to import the result to pandas\n')
    spinner = SpinnerCursor(text='Loading', spinner='dots')
    spinner.start()

    # use Pandas Library to query Excel Data
    df = pd.read_excel(
        config_values['filepath'],
        sheet_name=data_source['Sheet'],
        engine='openpyxl',
        skiprows=0,
        nrows=15
    )

    # stop waiting spinner
    spinner.stop()

    print(df.head())

    print('\nApplication Finished\n')


if __name__ == '__main__':
    try:
        main()
        logger.info("The Workbook refresh was done successfully!")
    except Exception as e:
        logger.critical(f"Couldn't refresh the data. ({e.args[0]} | {e.args[1]})")


# change variable value
