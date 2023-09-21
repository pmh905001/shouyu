import os

import iniconfig

from shouyu.package import Package


class ConfigManager:
    ini = None

    @classmethod
    def _load(cls):
        if not os.path.exists('kb.ini'):
            os.system(f'copy {Package.get_resource_path("kb.ini")} kb.ini')
        cls.ini = iniconfig.IniConfig("kb.ini")
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
    def shortcut(cls, key, default=None, convert=None):
        return cls.get(key, default, 'shortcuts', convert)


if __name__ == '__main__':
    print(ConfigManager.get('excel_path', 'kb.xlsx'))
