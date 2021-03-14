# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
from sapRefresh.Core.Cripto import secret_decode
from sapRefresh.Core.Time import timeit


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
    To initially refresh the data in the workbook.
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
    data_values['DataSourceName'] = xl_Instance.Application.Run("SapGetSourceInfo",
                                                                data_values['DS'],
                                                                "DataSourceName")
    data_values['Query'] = xl_Instance.Application.Run("SapGetSourceInfo", data_values['DS'], "QueryTechName")
    data_values['System'] = xl_Instance.Application.Run("SapGetSourceInfo", data_values['DS'], "System")
    return data_values
