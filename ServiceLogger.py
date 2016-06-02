# -*- coding: utf-8 -*-
"""
Created on Thu May 26 12:26:13 2016

@author: Dani
"""

from logging.handlers import RotatingFileHandler
from logging import Formatter, getLogger, DEBUG
import os


class ServiceLogger:

    def __init__(self, log, name, kind):
        handler = RotatingFileHandler(os.path.join("logs",log),maxBytes=1024*1024,backupCount=10)
        formatter = Formatter('[' + kind + "]\t%(asctime)s\t%(name)-15s\t%(levelname)-7.7s\t%(message)s","%d-%m-%Y %H:%M:%S")
        handler.setFormatter(formatter)
        self.logger = getLogger(name)
        self.logger.handlers = []
        self.logger.addHandler(handler)
        self.logger.setLevel(DEBUG)

    def __del__(self):
        del self.logger

    def log(self, message):
        self.logger.info(message)
