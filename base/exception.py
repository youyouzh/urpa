from base.log import logger


class MessageSendException(Exception):

    def __init__(self, message: str):
        logger.error('MessageSendException: ' + message)
        self.message = message


# 参数检查异常
class ParamInvalidException(Exception):

    def __init__(self, message: str, params=None):
        logger.warning('ParamInvalidException {}, params: {}'.format(message, params))
        self.message = message
        self.params = params


# ui自动化执行过程和Control交互异常
class ControlInvalidException(Exception):

    def __init__(self, message: str, selector=None):
        logger.error('ControlInvalidException {}, selector: {}'.format(message, selector))
        self.message = message
