import os
import shutil

import iniconfig

from shouyu.util.package import Package


class Config:
    FILE_NAME = 'kb.ini'
    ini = None

    @classmethod
    def _load(cls):
        if not os.path.exists(cls.FILE_NAME):
            source_path = Package.get_resource_path(cls.FILE_NAME)
            shutil.copy(source_path, cls.FILE_NAME)
        cls.ini = iniconfig.IniConfig(cls.FILE_NAME)
        return cls.ini

    @classmethod
    def get(cls, key, default, section='basic', convert=None):
        ini = cls._load()
        if convert:
            return ini.get(section, key, default, convert)
        else:
            return ini.get(section, key, default)

    @classmethod
    def excel_path(cls):
        return cls.get('excel_path', 'kb.xlsx')

    @classmethod
    def get_shortcut(cls, key, default=None, convert=None):
        return cls.get(key, default, 'shortcuts', convert)


if __name__ == '__main__':
    print(Config.get('excel_path', 'kb.xlsx'))
