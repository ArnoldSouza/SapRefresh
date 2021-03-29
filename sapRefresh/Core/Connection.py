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
import smtplib
import ssl
from email.mime.text import MIMEText
from datetime import datetime

from tenacity import retry, wait_fixed, before_sleep_log, stop_after_attempt

import logging
from Core.Cripto import secret_decode
from sapRefresh.Core.base_logger import get_logger
from sapRefresh import LOG_PATH, global_configs_df
logger, LOG_FILEPATH = get_logger(__name__, LOG_PATH)


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


def send_email(message_string, status_string, process_string):  #filepath_string, log_path_string, status_string, process_string):
    """Function to send communications via email"""
    # calculated fields
    log_path_string = str(LOG_FILEPATH)
    datetime_string = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
    subject_string = f'PYTHON AUTOMATE ({status_string}) - [{process_string}] - {datetime_string}'
    # email configurations
    smtp_server = global_configs_df.query('description=="mail-server"')['value'].values[0]
    port = global_configs_df.query('description=="mail-port"')['value'].values[0]
    sender_email = global_configs_df.query('description=="mail-user_name"')['value'].values[0]
    password = secret_decode(global_configs_df.query('description=="mail-password"')['value'].values[0])
    receiver_email = (global_configs_df.query('description=="mail-to"')['value'].values[0]).split(',')
    # elaborate the email message
    msg_email = f"""\
    Ola,

    Segue o status do processo [{process_string}]:

        STATUS: {status_string}
        DATA / HORA: {datetime_string}
        MENSAGEM: {message_string}
        DIRETORIO LOG: {log_path_string}

    @2021 GBS Latam - Hydro
    """
    # construct email message
    message = MIMEText(msg_email)
    message['subject'] = subject_string
    message['from'] = sender_email
    message['to'] = global_configs_df.query('description=="mail-to"')['value'].values[0]
    # Create a secure SSL context
    context = ssl.create_default_context()
    # log in to server and send email
    server = smtplib.SMTP(smtp_server, port)
    server.ehlo()  # Can be omitted
    server.starttls(context=context)  # Secure the connection
    server.ehlo()  # Can be omitted
    server.login(sender_email, password)
    # Statement to send email
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()


if __name__ == '__main__':
    filepath_str = r'C:\teste\alguma_pasta\algum_arquivo.xlsx'
    status_str = 'SUCCESS'
    process_str = 'SAP Refresh'
    send_email(filepath_str, status_str, process_str)
