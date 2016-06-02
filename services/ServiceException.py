class ServiceException(Exception):
    def __init__(self, message, line):
        Exception.__init__(self, message + str(" at %d"%(line,)))


class LoggerException(Exception):
    def __init__(self, message):
        Exception.__init__(self, message)


class ScheduleException(ServiceException):
    def __init__(self, message, line):
        ServiceException.__init__(self, message, line,)


class ScriptException(ServiceException):
    def __init__(self, message, line):
        ServiceException.__init__(self, message, line,)

