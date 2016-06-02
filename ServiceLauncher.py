# -*- coding: utf-8 -*-
"""
Created on Thu May 26 10:27:40 2016

@author: Dani
"""

import win32service
import win32serviceutil
import win32api
import win32con
import win32event
import win32evtlogutil
import servicemanager
import inspect, os, glob, re, sys

from importlib import import_module, reload, invalidate_caches

from ServiceLogger import ServiceLogger

if hasattr(sys,"frozen") and sys.frozen in ("windows_exe", "console_exe"):
    cwd = os.path.dirname(os.path.abspath(sys.executable))
else:
    cwd = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
os.chdir(cwd)

svcDir = os.path.join(cwd,"services")
sys.path.append(svcDir)

from ServiceException import LoggerException, ScheduleException, ScriptException

class aservice(win32serviceutil.ServiceFramework):
   
    _svc_name_ = "PySvcLauncher"
    _svc_display_name_ = "Lanzador de servicios python"
    _svc_description_ = "Servicio de gestión de ejecución de scripts de Python"
         
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None,0,0,None)
        self.stop_requested = False
        self.logger = ServiceLogger("servicelauncher.log","servicelauncher",'ServiceLauncher')

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        #win32event.SetEvent(self.hWaitStop)
        win32event.SetEvent(self.stop_event)
        self.log('Stopping service ...')
        self.stop_requested = True
       
    def log(self, message):
        self.logger.log(message)
         
    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,servicemanager.PYS_SERVICE_STARTED,(self._svc_name_, ''))

        self.timeout = 60000     #60 seconds / 1 minute
        # This is how long the service will wait to run / refresh itself (see script below)

        instances = {}
        modules = {}
        while 1:
            # Wait for service stop signal, if I timeout, loop again
            rc = win32event.WaitForSingleObject(self.stop_event, self.timeout)
            # Check to see if self.hWaitStop happened
            if rc == win32event.WAIT_OBJECT_0:
                # Stop signal encountered
                servicemanager.LogInfoMsg(self._svc_name_ + " - STOPPED!")  # For Event Log
                self.log(self._svc_name_ + " - STOPPED!")
                break
            else:

                """ Ok, here's the real money shot right here.
                    [actual service code between rests]        """
                svcScripts = glob.glob(os.path.join(svcDir, "*Svc.py"))
                svcScripts = [re.split(r'[\\/]+', s)[-1] for s in svcScripts]

                for script in svcScripts:
                    try:
                        if script not in modules.keys():
                            modules.update({script: import_module(script[:-3])})
                        else:
                            modules[script] = reload(modules[script])
                    except Exception as x:
                        message = str("%s:%s - Error: %s at %d" % (
                            self._svc_name_, script, str(x), sys.exc_info()[-1].tb_lineno))
                        servicemanager.LogInfoMsg(message)  # For Event Log
                        self.log(message)
                        continue

                todel = []
                for script in [s for s in svcScripts if s not in instances.keys()]:
                    try:
                        aux_class = getattr(modules[script], script[:-6])
                        instances.update({script: aux_class()})
                    except Exception as x:
                        message = str("%s:%s - Error: %s at %d" % (
                            self._svc_name_, script, str(x), sys.exc_info()[-1].tb_lineno))
                        servicemanager.LogInfoMsg(message)  # For Event Log
                        self.log(message)
                        if script in instances.keys():
                            todel.append(script)  # Service will be forced to reload Script
                        continue
                    self.log('"%s" Added to Script list' % (script,))

                if len(todel) > 0:
                    for script in todel:
                        del instances[script]
                    del todel
                    invalidate_caches()


                for script in [s for s in instances.keys() if s not in svcScripts]:
                    try:
                        del instances[script]
                    except Exception as x:
                        message = str("%s:%s - Error: %s at %d" % (
                            self._svc_name_, script, str(x), sys.exc_info()[-1].tb_lineno))
                        servicemanager.LogInfoMsg(message)  # For Event Log
                        self.log(message)
                        continue
                    self.log('"%s" Removed from Script list' % (script,))

                todel = []
                for script in instances.keys():
                    try:
                        instances[script].run()
                    except(ScriptException, LoggerException, ScheduleException) as x:
                        servicemanager.LogInfoMsg(str(x))  # For Event Log
                        self.log(str(x))
                        todel.append(script)  # Service will be forced to reload Script
                        continue
                    except Exception as x:
                        message = str("%s:%s - Error: %s at %d" % (
                            self._svc_name_, script, str(x), sys.exc_info()[-1].tb_lineno))
                        servicemanager.LogInfoMsg(message)  # For Event Log
                        self.log(message)
                        todel.append(script)  # Service will be forced to reload Script
                        continue

                if len(todel) > 0:
                    for script in todel:
                        del instances[script]
                    del todel
                    invalidate_caches()
            """ [actual service code between rests] """


def ctrlHandler(ctrlType):
    return True
                  
if __name__ == '__main__':   
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)
    win32serviceutil.HandleCommandLine(aservice)

#if __name__ == '__main__':
#    if len(sys.argv) == 1:
#        servicemanager.Initialize()
#        servicemanager.PrepareToHostSingle(aservice)
#        servicemanager.StartServiceCtrlDispatcher()
#    else:
#        win32serviceutil.HandleCommandLine(aservice)
        
    #print(win32serviceutil.QueryServiceStatus("PySvcLauncher"))