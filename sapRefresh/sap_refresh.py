# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import os
import sys
import pathlib
from datetime import date

import pandas as pd
from halo import Halo
from tenacity import retry, wait_fixed, before_sleep_log, stop_after_attempt

MODULE_CURRENT_PATH = os.path.abspath('.')
MODULE_PARENT_PATH = os.path.dirname(MODULE_CURRENT_PATH)
sys.path.append(MODULE_PARENT_PATH)

from sapRefresh.Core import Connection as Conn
from sapRefresh.Core.Time import timeit, get_time_intelligence
from sapRefresh.Refresh import Engine as Xl
from sapRefresh.Refresh import Sap

# configure the log object
import logging
from sapRefresh.Core.base_logger import get_logger
from sapRefresh import LOG_PATH, global_configs_df, CONFIG_PATH  # import info from the __init__ file
logger, LOG_FILEPATH = get_logger(__name__, LOG_PATH)


@retry(reraise=True, wait=wait_fixed(10), before_sleep=before_sleep_log(logger, logging.DEBUG), stop=stop_after_attempt(3))
@timeit
def get_configurations(config_path):
    """get all necessary dataframes to serve the SapRefresh class"""
    global_configs = pd.read_excel(config_path, sheet_name='global_configs')
    data_sources = pd.read_excel(config_path, sheet_name='data_sources')
    variables_filters = pd.read_excel(config_path, sheet_name='variables_filters')
    print('Successfully loaded the configurations')
    return global_configs, data_sources, variables_filters


class SapRefresh:
    """Class that handles all the events related to the SapRefresh application"""
    def __init__(self, config_df=global_configs_df):
        # global configuration dataframe
        self.global_configs = config_df
        # file related parameter. Need to be set only when a workbook is loaded
        self.data_source = None  # dictionary with all infos about the data source
        self.source = None  # the name of the data source that was loaded (usually DS_1)
        self.filepath = None  # the filepath of the SAP AfO file
        # Classes of Excel API
        self.ExcelInstance = None
        self.WorkbookSAP = None
        # parameters set when workbook is opened
        self.calc_state_init = None
        # state parameters
        self.is_logged = None
        self.is_refreshed = None
        self.is_refreshed_data = None
        self.state_refresh_behavior = None
        self.state_variable_submit = None
        # List of all variables and filters in the data source
        self.variables_filters = None

    def open_report(self, filepath):
        """
        Initiate Excel instance. Configure it to optimize the execution. Then open the target workbook.
        Call method to activate SAP AfO AddIn. Capture the initial calculation state. Activate workbook.
        Finally do the initial calculation.
        """
        # ensure the application is in the correct network
        host = self.global_configs.query('description=="ping-host"')['value'].values[0]
        port = self.global_configs.query('description=="ping-port"')['value'].values[0]
        Conn.check_connection(host, port)
        # ensure the filepath is of the correct type
        if type(filepath) is not pathlib.WindowsPath:
            raise TypeError(f'The filepath variable ({filepath}) is not a pathlib.WindowsPath class')
        # kill every excel instance so that no further erros can happen
        Xl.kill_excel_instances()
        self.ExcelInstance = Xl.open_excel()
        Xl.optimize_instance(self.ExcelInstance, 'start')
        self.WorkbookSAP = Xl.open_workbook(self.ExcelInstance, filepath)
        Xl.ensure_addin(self.ExcelInstance)
        self.calc_state_init = Xl.calculation_state(self.ExcelInstance, 'start')
        Xl.ensure_wb_active(self.ExcelInstance, filepath.name)
        # assign the filepath of the SAP AfO file
        self.filepath = filepath
        print('The report is loaded')

    def calculate(self):
        """Calculate workbook to refresh values"""
        self.ExcelInstance.Application.Calculate()

    def get_data_source(self):
        """get data information about SAP AfO objects in a workbook"""
        self.data_source = Xl.get_data_source(self.ExcelInstance)
        self.source = self.data_source['DS']  # set source to a property for the easy access it
        print('\n', 'data source retrieved', '\n', self.data_source)

    def logon(self, source=None):
        """
        Logon into the SAP AfO System. The logon is file dependent. That's because you need to refer the
        data source to connect to SAP. It uses the source of the get_data_source method.
        """
        # assign variables
        client = self.global_configs.query('description=="logon-client"')['value'].values[0]
        user = self.global_configs.query('description=="logon-user"')['value'].values[0]
        password = self.global_configs.query('description=="logon-password"')['value'].values[0]
        if source is not None:
            self.source = source
        # execute the logon method
        self.is_logged = Sap.sap_logon(self.ExcelInstance, self.source, client, user, password)

    def refresh(self):
        """Do there initial refresh of data in the workbook."""
        self.is_refreshed = Sap.sap_refresh(self.ExcelInstance)

    def refresh_data(self):
        """Refresh the transaction data for all data sources in the workbook."""
        self.is_refreshed_data = Sap.sap_refresh_data(self.ExcelInstance, self.data_source)

    def additional_source_info(self):
        """
        Query more information and append it to the data source dictionary.
        Attention. This method is dependent of Data Source initiation.
        """
        self.data_source = Sap.sap_get_more_info(self.ExcelInstance, self.data_source)
        print('Additional data source information retrieved', '\n', self.data_source)

    @timeit
    def close(self):
        """
        Make all the necessary procedures to terminate the excel instance.
            - put the  calculation back to the original state
            - terminate the workbook instance
            - Configure the Excel Instance to the original state
            - Close the Excel Instance
        """
        Xl.calculation_state(self.ExcelInstance, 'stop', self.calc_state_init)
        Xl.close_workbook(self.WorkbookSAP)
        Xl.optimize_instance(self.ExcelInstance, 'stop')
        self.ExcelInstance.Application.Quit()
        self.WorkbookSAP = None
        self.ExcelInstance = None
        print('The application was Successfully closed')

    def get_variables_list(self):
        """Return a dictionary of the variables that exists in the data source"""
        variables_list = Sap.sap_get_variables(self.ExcelInstance, self.source)
        return variables_list

    def variables_filters_list(self):
        """Return a dataframe with all the variables and filters inside the datasource"""
        # get the list of variables
        variables_list = Sap.sap_get_variables(self.ExcelInstance, self.source)
        # get technical name of the variables and append to Restrictions list
        restrictions = []
        for variable in variables_list:
            restrictions.append(
                {
                    'command': 'SAPSetVariable',
                    'field': Sap.sap_get_technical_name(self.ExcelInstance, self.source, variable[0]),
                    'field_name': variable[0],
                    'value': variable[1]
                }
            )
        # get the list of filters (measures)
        filters_list = Sap.sap_get_filters(self.ExcelInstance, self.source)
        # get list of dimensions (fields)
        dimensions_list = Sap.sap_get_dimensions(self.ExcelInstance, self.source)
        # search in dimensions the technical name of each filter then append values to Restrictions list
        for filter_ in filters_list:
            if filter_[0] != 'Measures':
                values = dict()
                values['command'] = 'SAPSetFilter'
                for dimension in dimensions_list:
                    if dimension[1] == filter_[0]:
                        values['field'] = dimension[0]  # get the technical name
                values['field_name'] = filter_[0]
                values['value'] = filter_[1]
                restrictions.append(values)
        # create the dataframe with filters and variables
        # noinspection PyTypeChecker
        variables_filters = pd.DataFrame.from_dict(restrictions)
        variables_filters['data_source'] = self.source
        variables_filters['reference_type'] = 'value'
        variables_filters['data_source_name'] = self.data_source['DataSourceName']
        variables_filters['data_source_sheet'] = self.data_source['Sheet']
        # assign values to properties
        self.variables_filters = variables_filters
        return variables_filters

    def data_source_list(self):
        """Return a Df with data source information"""
        data_source = pd.DataFrame(list(self.data_source.items()), columns=['Key', 'Value'])
        return data_source

    @timeit
    def export_variables_filters(self):
        """export to an Excel file the data source information and variables and filters values"""
        # load information from class' properties
        filepath = self.filepath
        path_data_info = pathlib.Path(self.global_configs.query('description=="path-data_info"')['value'].values[0])
        # create the pathname
        name_file = filepath.name[:len(filepath.suffix) * -1]
        complement = '__information'
        file_extension = filepath.suffix
        new_name = pathlib.Path(name_file + complement + file_extension)
        new_filepath = path_data_info / new_name
        # assign dataframes
        data_source_info = self.data_source_list()
        variables_filters_info = self.variables_filters_list()
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(new_filepath, engine='xlsxwriter')
        # write the dataframes
        data_source_info.to_excel(writer, sheet_name='data_source_info')
        variables_filters_info.to_excel(writer, sheet_name='variables_filters_info')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

    def is_ds_active(self):
        """check whether a data source is active"""
        state_data_source = Sap.sap_is_ds_active(self.ExcelInstance, self.source)
        return state_data_source

    def is_connected(self):
        """check whether a data source is already connected"""
        state_connection = Sap.sap_is_connected(self.ExcelInstance, self.source)
        return state_connection

    def set_refresh_variables(self, variables_list):
        print('Starting to set the variables:')
        self.state_refresh_behavior = self.ExcelInstance.Application.Run("SAPSetRefreshBehaviour", "Off")
        self.state_variable_submit = self.ExcelInstance.Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "On")
        for index, variable in variables_list.iterrows():
            print('\t', f'Trying to set [{variable.field_name}] to [{variable.value}] ', end='')
            self.ExcelInstance.Application.Run(
                variable.command,
                variable.field,
                variable.value,
                "INPUT_STRING",
                variable.data_source)
            print('Ok!')
        # start waiting spinner
        spinner = Halo(text='Loading', spinner='dots')
        spinner.start()
        # refresh data with new variable values
        self.state_variable_submit = self.ExcelInstance.Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "Off")
        self.state_refresh_behavior = self.ExcelInstance.Application.Run("SAPSetRefreshBehaviour", "On")
        # stop waiting spinner
        spinner.succeed('End!')
        print('The variables were set properly')

    def set_refresh_filters(self, df_filters):
        print('Starting to set the filters:')
        self.state_refresh_behavior = self.ExcelInstance.Application.Run("SAPSetRefreshBehaviour", "Off")
        for index, filter_item in df_filters.iterrows():
            print('\t', f'Trying to set [{filter_item.field_name}] to [{filter_item.value}] ', end='')
            self.ExcelInstance.Application.Run(
                filter_item.command,
                filter_item.data_source,
                filter_item.field,
                filter_item.value,
                "INPUT_STRING"
            )
            print('Ok!')
        # start waiting spinner
        print('Start to refresh the data with the new restrictions')
        spinner = Halo(text='Loading', spinner='dots')
        spinner.start()
        # refresh data with new variable values
        self.state_refresh_behavior = self.ExcelInstance.Application.Run("SAPSetRefreshBehaviour", "On")
        # stop waiting spinner
        spinner.succeed('End!')
        print('The filters were set properly')


def get_report_information(filepath):
    """function to collect variables and filters of a report given a filepath"""
    # initiate workbook
    SapReport = SapRefresh()
    SapReport.open_report(filepath)
    SapReport.calculate()
    # logon and get information from data source
    SapReport.get_data_source()
    SapReport.logon()
    SapReport.refresh()
    SapReport.additional_source_info()
    # export variables and filters to an Excel file
    SapReport.export_variables_filters()
    SapReport.close()


def refresh_report(filename, data_sources, variables_filters):
    """execute the flow necessary to refresh de desired report"""
    # initiate workbook
    SapReport = SapRefresh()
    # configure path
    data_directory = pathlib.Path(global_configs_df.query('description=="path-data_directory"')['value'].values[0])
    file_target = data_directory / filename
    # open de SAP AfO report
    SapReport.open_report(file_target)
    SapReport.calculate()
    # collect data sources
    current_source = data_sources.query(f'filename=="{filename}"')['data_source'].values[0]
    # logging on
    SapReport.logon(current_source)
    # do initial data refresh
    SapReport.refresh_data()
    # replace dynamic values in the parameters
    dict_time_values = get_time_intelligence()
    # start to deal with filters
    df_filters = variables_filters.query(f'filename=="{filename}" and command=="SAPSetFilter"').replace({"value": dict_time_values})
    if not df_filters.empty:
        SapReport.set_refresh_filters(df_filters)
    # start to deal with variables
    df_variables = variables_filters.query(f'filename=="{filename}" and command=="SAPSetVariable"').replace({"value": dict_time_values})
    if not df_variables.empty:
        SapReport.set_refresh_variables(df_variables)
    # save and close the report
    SapReport.close()
    # send email
    mail_msg = f'Relatorio atualizado -> {str(file_target)}'
    Conn.send_email(mail_msg, 'SUCCESS', 'SAP Refresh - Refresh Reports')


def collect_information():
    """collect al data source information of the reports that are the the \Data_refresh directory"""
    data_directory = pathlib.Path(global_configs_df.query('description=="path-data_directory"')['value'].values[0])
    # search the directory for excel files to be refreshed
    list_files = Conn.search_directory(data_directory)
    for file in list_files:
        print(f'Starting to extract info from {file}')
        file_target = data_directory / file
        get_report_information(file_target)  # execute the data source extraction
        print(f'finished extraction from {file}', '\n')


def refresh_auto_reports():
    """function to automate the refresh of reports based on the parameters set in config file"""
    _, data_sources, variables_filters = get_configurations(CONFIG_PATH)
    # change the dynamic days in the sources
    today = date.today().day
    data_sources['refresh'] = data_sources['refresh'].apply(lambda x: 'Y' if x >= today or x == 99 else 'N')
    # start to refresh the reports
    for filename in data_sources.query('refresh=="Y"').filename:  # execute command only to valid rows
        print(f'Starting to refresh the reports: [{filename}]')
        refresh_report(filename, data_sources, variables_filters)
        print(f'Finished the refreshing of: [{filename}]')


if __name__ == '__main__':
    try:
        refresh_auto_reports()
        logger.info("The Workbook refresh was done successfully!")
    except Exception as e:
        # send error to the logger
        logger.critical(f"Couldn't refresh the data. ({e.args[0]} | {e.args[1]})")
        # send error by email
        msg_mail = f"Couldn't refresh the data. ({e.args[0]} | {e.args[1]})"
        Conn.send_email(msg_mail, 'ERROR', 'SAP Refresh')
