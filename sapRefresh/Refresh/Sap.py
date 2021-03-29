# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
from tenacity import retry, wait_fixed, before_sleep_log, stop_after_attempt

from sapRefresh.Core.Cripto import secret_decode
from sapRefresh.Core.Time import timeit

import logging
from sapRefresh.Core.base_logger import get_logger
from sapRefresh import LOG_PATH
logger, LOG_FILEPATH = get_logger(__name__, LOG_PATH)


@retry(reraise=True, wait=wait_fixed(10), before_sleep=before_sleep_log(logger, logging.DEBUG), stop=stop_after_attempt(3))
@timeit
def sap_logon(xl_Instance, source, client, user, password):
    """API method to trigger a logon to a system for a specified data source"""
    result = xl_Instance.Application.Run("SAPLogon", source, client, secret_decode(user), secret_decode(password))
    if result == 1:
        print('\nSuccessfully logged in SAP AfO')
    else:
        raise ConnectionError("Couldn't login in SAP AfO")
    return result


@timeit
def sap_refresh(xl_Instance):
    """
    Do there initial refresh of data in the workbook.
    All data sources and planning objects will be refreshed.
    If you execute this command for a data source which is already refreshed, all corresponding crosstabs are redrawn.
    """
    result = xl_Instance.Application.Run("SAPExecuteCommand", "Refresh")
    if result == 1:
        print('\nSuccessfully refreshed the workbook')
    else:
        raise ConnectionError("Couldn't refresh the SAP AfO")  # todo: need to handle errors differently to not stop the runtime
    return result


@timeit
def sap_refresh_data(xl_Instance, source):
    """
    Refresh the transaction data for all or defined data sources in the workbook.
    The corresponding transaction data is updated from the server and the crosstabs are redrawn.
    """
    result = xl_Instance.Application.Run("SAPExecuteCommand", "RefreshData", source)
    if result == 1:
        print(f'\nSuccessfully refreshed the source: {source}')
    else:
        raise ConnectionError(f"Couldn't refresh the the source: {source}")
    return result


def sap_get_more_info(xl_Instance, data_values):
    """get other information about SAP data source"""
    data_values['DataSourceName'] = xl_Instance.Application.Run("SapGetSourceInfo", data_values['DS'], "DataSourceName")
    data_values['Query'] = xl_Instance.Application.Run("SapGetSourceInfo", data_values['DS'], "QueryTechName")
    data_values['System'] = xl_Instance.Application.Run("SapGetSourceInfo", data_values['DS'], "System")
    return data_values


def sap_get_variables(xl_Instance, source):
    """Get the list of variables of a data source"""
    variables_list = xl_Instance.Application.Run("SAPListOfVariables", source, "INPUT_STRING", "ALL")
    return variables_list


def sap_get_technical_name(xl_Instance, source, variable_name):
    """returns the value of the technical name for a specific variable"""
    var_tech_name = xl_Instance.Application.Run("SAPGetVariable", source, variable_name, "TECHNICALNAME")
    return var_tech_name


def sap_get_filters(xl_Instance, source):
    """Get the list of filters (measures) of a data source"""
    filters_list = xl_Instance.Application.Run("SAPListOfDynamicFilters", source, "INPUT_STRING")
    return filters_list


def sap_get_dimensions(xl_Instance, source):
    """Get the list of dimensions (fields) of a data source"""
    dimensions_list = xl_Instance.Application.Run("SAPListOfDimensions", source)
    return dimensions_list


def sap_is_ds_active(xl_Instance, source):
    """check whether a data source is active"""
    state_data_source = xl_Instance.Application.Run("SAPGetProperty", "IsDataSourceActive", source)
    return state_data_source


def sap_is_connected(xl_Instance, source):
    """check whether a data source is already connected"""
    state_connection = xl_Instance.Application.Run("SAPGetProperty", "IsConnected", source)
    return state_connection
