import os
import configparser as cp

import util

class Singleton(object):
    _instance = None
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(Singleton, cls).__new__(cls, *args, **kwargs)
        return cls._instance

class ConfigLoader(Singleton):
    def __init__(self):
        self.config = cp.ConfigParser()
        self.path = util.get_modifiedpath(os.path.join('resources', 'config.ini'), False)
        self.config.read(self.path, encoding='utf-8')

    def get_in_dir(self):
        return self.config['params']['in_dir_path']

    def get_out_dir(self):
        return self.config['params']['out_dir_path']

    def get_out_file(self):
        return self.config['params']['out_file_name']

    def get_master_title(self):
        return self.config['master']['title']
    
    def get_handler_heading(self):
        return self.config['handler']['heading']
    
    def get_handler_combo_values(self):
        return self.config['handler']['combo_values'].split(',')
    
    def get_main_heading1(self):
        return self.config['main']['heading1']
    
    def get_main_heading2(self):
        return self.config['main']['heading2']

    def get_main_heading3(self):
        return self.config['main']['heading3']

    def get_columns(self):
        return self.config['core']['combo_values'].split(',')

class ConfigWriter(Singleton):
    def __init__(self):
        # to keep comments in config.ini, need these arguments
        self.config = cp.ConfigParser(comment_prefixes='/', allow_no_value=True)
        self.path = util.get_modifiedpath(os.path.join('resources', 'config.ini'), False)
        self.config.read(self.path, encoding='utf-8')

    def save(self, in_dir, out_dir, out_file):
        self.config['params'] = {
            'in_dir_path': in_dir,
            'out_dir_path': out_dir,
            'out_file_name': out_file
        }
        
        with open(self.path, 'w', encoding='utf-8') as file:
            self.config.write(file)
    
