# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import win32com.client as win32
from win32com.client import constants as cst

from sapRefresh.Core.Time import timeit


def open_excel():
    """Start a instance of Excel application"""
    xl_Instance = win32.gencache.EnsureDispatch('Excel.Application')
    return xl_Instance


@timeit
def open_workbook(xl_Instance, path):
    """
    Open the file in the new Excel instance,
        The 1st false: don't update the links
        The 2nd false: and don't open it read-only
    """
    wb = xl_Instance.Workbooks.Open(path, False, False)
    return wb


@timeit
def ensure_addin(xl_Instance):
    """Force the plugin to be enabled in the instance of Excel"""
    for addin in xl_Instance.Application.COMAddIns:
        if addin.progID == 'SapExcelAddIn':
            if not addin.Connect:
                addin.Connect = True
            elif addin.Connect:
                addin.Connect = False
                addin.Connect = True
    print('\n', 'Is SapExcelAddIn Enabled?', xl_Instance.Application.COMAddIns['SapExcelAddIn'].Connect)


def ensure_wb_active(xl_Instance, filename):
    """Check if WorkBook is active, otherwise Logon fails if another wb is selected"""
    wb_name = xl_Instance.Application.ActiveWorkbook.Name
    print('Current workbook active is:', wb_name)
    if wb_name != filename:
        print(f'The desired wb ({filename}) is not active. Forcing it...', end='')
        xl_Instance.Application.Windows(filename).Activate()
        new_wb = xl_Instance.Application.ActiveWorkbook.Name
        print('Done! Workbook active is:', new_wb)


def optimize_instance(xl_Instance, action):
    """deals with excel calculation optimization"""
    if action == 'start':
        xl_Instance.Visible = True
        xl_Instance.DisplayAlerts = False
        xl_Instance.ScreenUpdating = False
        # xl_Instance.EnableEvents = False  # todo: check in reference code if this statement cause negative behavior in the script before uncomment it
    elif action == 'stop':
        xl_Instance.DisplayAlerts = True
        xl_Instance.ScreenUpdating = True
        # xl_Instance.EnableEvents = True
        xl_Instance.Application.Cursor = cst.xlDefault
        xl_Instance.Application.StatusBar = ''  # equivalent to vbNullString


def calculation_state(xl_Instance, action, state=None):
    """set the Calculation State to Manual in Excel"""
    if action == 'start':
        state = xl_Instance.Application.Calculation
        xl_Instance.Application.Calculation = cst.xlCalculationManual
    elif action == 'stop' and state is not None:
        xl_Instance.Application.Calculation = state
    return state


def get_data_source(xl_Instance):
    """get data information about SAP AfO objects"""
    values = dict()
    (
        values['CrossTabSource'],
        values['CrossTabName'],
        values['DS']
    ) = xl_Instance.Application.Run("SAPListOf", "CROSSTABS")

    values['Sheet'] = xl_Instance.ActiveWorkbook.Names("SAP" + values['CrossTabSource']).RefersToRange.Parent.Name
    values['Crosstab'] = "SAP" + values['CrossTabSource']
    return values


def close_workbook(wb_Instance):
    """Save the file and close it"""
    wb_Instance.Save()
    wb_Instance.Close()
