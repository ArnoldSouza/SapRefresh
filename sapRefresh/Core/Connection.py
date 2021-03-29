# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import socket
from os import listdir
from os.path import isfile, join
from pathlib import Path
import configparser

from tenacity import retry, wait_fixed, before_sleep_log, stop_after_attempt

import logging

from sapRefresh.Core.base_logger import get_logger
from sapRefresh import LOG_PATH
logger = get_logger(__name__, LOG_PATH)


def get_config_values(config_file='app_config.ini'):
    """Get configurations to the application from a configuration.ini file"""
    config = configparser.ConfigParser()
    config.read(config_file)  # get values from INI File

    configuration = dict()
    configuration['filename'] = config['VARIABLES']['filenames']
    configuration['bwclient'] = config['CONNECTION']['bwclient']
    configuration['bwuser'] = config['CONNECTION']['bwuser']
    configuration['bwpassword'] = config['CONNECTION']['bwpassword']
    configuration['app_server'] = config['APPLICATION_SERVER']['host']
    configuration['app_port'] = config['APPLICATION_SERVER']['port']
    # get the path to the file
    configuration['filepath'] = _get_wb_path(configuration['filename'])
    return configuration


def _get_wb_path(filename):
    """Path of excel file to import"""
    workbook_filepath = Path.cwd().joinpath('Workbooks/'+filename)
    return workbook_filepath


@retry(reraise=True, wait=wait_fixed(10), before_sleep=before_sleep_log(logger, logging.DEBUG),
       stop=stop_after_attempt(3))
def check_connection(host, port, timeout=3):
    """Check if a server is alive or not"""
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, int(port)))
        print(f"The server {host}:{port} is reachable")
    except socket.error as ex:
        raise RuntimeError(f"Error! Couldn't connect to {host}:{port}. Exception:", ex)


def search_directory(data_directory):
    """search the directory for excel files to be refreshed"""
    onlyfiles = [f for f in listdir(data_directory) if isfile(join(data_directory, f))]
    list_files = []
    for filename in onlyfiles:
        if filename[-4:].lower() == 'xlsx':
            list_files.append(filename)
    return list_files
