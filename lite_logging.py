# -*- coding: utf-8 -*-
import os
from datetime import datetime


class Logger(object):
    '''Class for writing logs to a text file

    To write logs, you need to create a Logger object and assign it a path and file name for recording\n
    On this object, call the write_log method with text and set status.

    static statuses:\n
    DEBUG\n
    ERROR\n
    INFO\n
    WARNING

    constructor keyword arguments:\n
    parent_folders_path -- folders for log storage like this "Dir1\\Dir2" (default empty string)\n
    file_name           -- file name without time and extension (default 'temp')
    default_status      -- status for most reports (default INFO)
    '''

    DEBUG = '[ {:<7} ] '.format('DEBUG')
    ERROR = '[ {:<7} ] '.format('ERROR')
    INFO = '[ {:<7} ] '.format('INFO')
    WARNING = '[ {:<7} ] '.format('WARNING')

    def __init__(self, parent_folders_path='', file_name='temp', default_status=INFO):
        self._initialize_components(parent_folders_path, file_name)
        self._create_dir()
        self._default_status = default_status

    def _initialize_components(self, parent_folders_path, file_name):
        if parent_folders_path:
            self._log_dir_path = os.path.join(os.getenv('appdata'), parent_folders_path)
        else:
            self._log_dir_path = os.getenv('temp')
        self._log_file_path = os.path.join(self._log_dir_path, file_name + datetime.today().strftime('_%Y_%m_%d_%H_%M_%S') + '.txt')

    def _create_dir(self):
        if not os.path.exists(self._log_dir_path):
            os.makedirs(self._log_dir_path)

    def write_log(self, text, status=''):
        '''assign a status to this entry. Default status assign in constructor'''
        if not status:
            status = self._default_status
        if not isinstance(text, str):
            text = str(text)
        with open(self._log_file_path, 'a') as log_file:
            for i, row in enumerate(text.split('\n')):
                print(i, row)
                if i == 0:
                    log_file.write(status + row.encode('utf-8') + '\n')
                else:
                    log_file.write('{:<12}'.format('') + row.encode('utf-8') + '\n')

    def add_blank_line(self, separator=None):
        '''inserts a empty row for logical separation'''
        if separator is not None:
            if not isinstance(separator, str):
                separator = str(separator)
            separate_line = '\n' + separator * 150 + '\n\n'
        else:
            separate_line = '\n'
        with open(self._log_file_path, 'a') as log_file:
            log_file.write(separate_line)

    def get_log_file_path(self):
        if self._log_file_path:
            return self._log_file_path
        else:
            return None

    def get_log_dir_path(self):
        if self._log_dir_path:
            return self._log_dir_path
        else:
            return None


if __name__ == '__main__':
    logger = Logger(parent_folders_path=os.path.join('Synergy Systems', 'Logger'),
                    file_name='test_log',
                    default_status=Logger.WARNING)

    logger.write_log('test text for log')
    logger.add_blank_line()
    logger.write_log('second test text', Logger.ERROR)
    logger.write_log('third\n'
                     'test\n'
                     'text', Logger.WARNING)

    print(logger.get_log_dir_path())
    print(logger.get_log_file_path())
