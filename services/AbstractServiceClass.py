
import sys
import os
import inspect
import datetime
import time
from crontab import CronTab
from croniter import croniter
from ServiceException import LoggerException, ScheduleException, ScriptException
from abc import ABCMeta, abstractmethod

if hasattr(sys,"frozen") and sys.frozen in ("windows_exe", "console_exe"):
    cwd = os.path.dirname(os.path.abspath(sys.executable))
else:
    cwd = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
os.chdir(cwd + '\\..')

svcDir = os.path.join(cwd,"services")
sys.path.append(svcDir)

from ServiceLogger import ServiceLogger

# doit decorator for exception handling
def doitHandler(classname):
    def realHandler(func):
        def wrapper(*args, **kwds):
            try:
                func(*args, **kwds)
            except Exception as x:
                raise ScriptException('Service %s failed on Script handling. Message: %s' % (
                    classname, x), inspect.getinnerframes(sys.exc_info()[2])[-1][2])

        return wrapper
    return realHandler

class AbstractServiceClass(metaclass=ABCMeta):

        def __init__(self):
            self._classname_ = self.__class__.__name__
            self.__load__logger__()
            self.__load__schedule__()

        def __del__(self):
            del self._logger_

        @abstractmethod
        def __doit__(self): pass

        def __doit__handled__(self):
            @doitHandler(self._classname_)
            def handledScript():
                self.__doit__()
            handledScript()

        def __load__logger__(self):
            try:
                self._logger_ = ServiceLogger('services.log', (self._classname_[:12] + '...') if len(self._classname_) > 12 else self._classname_, 'Service')
            except Exception as x:
                raise LoggerException('Service %s failed on logger instantiation. Message: %s' % (
                    self._classname_, x))

        def __load__schedule__(self):
            try:
                fcron = CronTab(tabfile="services\\serviceCron.tab")
                job = [j for j in list(fcron.find_command(self._classname_)) if j.command == self._classname_]
                if len(job) <= 0:
                    raise Exception('Crontab command %s not found'%(self._classname_,))
                self._next_run_ = time.mktime(job[0].schedule().get_next().timetuple())
            except Exception as x:
                raise ScheduleException('Service %s failed on Crontab instantiation or handling. Message: %s' % (
                    self._classname_, x), sys.exc_info()[-1].tb_lineno)

            return True

        def __may__run__(self):

            if self._next_run_ <= time.mktime(datetime.datetime.now().timetuple()):
                return self.__load__schedule__()

            return False

        def log(self,message):
            self._logger_.log(message)

        def run(self):
            if self.__may__run__():
                self.__doit__handled__()
