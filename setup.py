# -*- coding: utf-8 -*-
"""
Created on Fri May 27 13:51:40 2016

@author: Dani
"""
from distutils.core import setup
import py2exe

class Target:
    def __init__(self, **kw):
        self.__dict__.update(kw)
        # for the versioninfo resources
        self.version = "0.5.0"
        self.company_name = "No Company"
        self.copyright = "no copyright"
        self.name = "py2exe sample files"


myservice = Target(
    description = 'foo',
    modules = ['ServiceLauncher'],
    cmdline_style='pywin32'
)

setup(
    options = {"py2exe": {"compressed": 1, "bundle_files": 1} },    
    console=["ServiceLauncher.py"],
    zipfile = None,
    service=[myservice]
) 