# -*- coding: utf-8 -*-
"""
Created on Thu May 26 12:26:13 2016

@author: Dani
"""

from AbstractServiceClass import AbstractServiceClass


class LogWriter(AbstractServiceClass):
    def __init__(self):
        AbstractServiceClass.__init__(self)

    def __doit__(self):

        """ Put your code in Here Plz :) """

        self.log('Too much work in here...')

        """ Not so fast buccaneer """