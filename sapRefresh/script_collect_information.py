# -*- coding: utf-8 -*-
"""
Created on 3/29/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
from sap_refresh import collect_information
from sapRefresh.Core.base_logger import get_logger
from sapRefresh import LOG_PATH
logger, LOG_FILEPATH = get_logger(__name__, LOG_PATH)


try:
    collect_information()
    logger.info("The information collection was done successfully!")
except Exception as e:
    # send error to the logger
    logger.critical(f"Couldn't refresh the data. ({e.args[0]} | {e.args[1]})")
