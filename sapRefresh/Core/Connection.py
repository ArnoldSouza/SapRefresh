# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import socket
import functools
import time
import sys
from pathlib import Path
import configparser


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


def check_connection(host, port, timeout=3):
    """Check if a server is alive or not"""
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, int(port)))
        print(f"The server {host}:{port} is reachable")
        response = True
    except socket.error as ex:
        print(f"Error! Couldn't connect to {host}:{port}. Exception:", ex)  # todo: need to handle errors differently to not stop the runtime
        response = False
    return response


def retry(ExceptionToCheck, tries=4, delay=3, backoff=2, logger=None):
    def deco_retry(f):
        @functools.wraps(f)
        def f_retry(*args, **kwargs):
            mtries, mdelay = tries, delay
            while mtries > 1:
                try:
                    return f(*args, **kwargs)
                except ExceptionToCheck as e:
                    msg = "%s: %s, Retrying in %d seconds..." % (f.__name__, str(e), mdelay)
                    if logger:
                        logger.warning(msg)
                    else:
                        print(msg)
                    time.sleep(mdelay)
                    mtries -= 1
                    mdelay *= backoff
            if mtries == 1:
                msg = "Couldn't run the application"
                if logger:
                    logger.error(msg)
                else:
                    print(msg)
                sys.exit()
            return f(*args, **kwargs)
        return f_retry  # true decorator
    return deco_retry
