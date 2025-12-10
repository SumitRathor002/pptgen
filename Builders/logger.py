import logging
import sys

class RefLogger:
    def __init__(self):
        self.logger = logging.getLogger("RefLogger")
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [RefID: %(reference_id)s] - %(message)s')
        handler.setFormatter(formatter)
        if not self.logger.handlers:
            self.logger.addHandler(handler)
        self.logger.setLevel(logging.INFO)

    def _log(self, level, message, reference_id=None):
        extra = {'reference_id': reference_id if reference_id is not None else 'N/A'}
        if level == 'info':
            self.logger.info(message, extra=extra)
        elif level == 'error':
            self.logger.error(message, extra=extra)
        elif level == 'warning':
            self.logger.warning(message, extra=extra)
        elif level == 'exception':
            self.logger.exception(message, extra=extra)

    def info(self, message, reference_id=None):
        self._log('info', message, reference_id)

    def error(self, message, reference_id=None):
        self._log('error', message, reference_id)

    def warning(self, message, reference_id=None):
        self._log('warning', message, reference_id)

    def exception(self, message, reference_id=None):
        self._log('exception', message, reference_id)

logger = RefLogger()
