import iniconfig


class ConfigManager:
    ini = iniconfig.IniConfig("kb.ini")

    @classmethod
    def get(cls, key, default, section='basic', convert=None):
        if convert:
            return cls.ini.get(section, key, default, convert)
        else:
            return cls.ini.get(section, key, default)

    @classmethod
    def excel_path(cls):
        return cls.get('excel_path', 'kb.xlsx')

    @classmethod
    def shortcut(cls, key, default=None, convert=None):
        return cls.get(key, default, 'shortcuts', convert)


if __name__ == '__main__':
    print(ConfigManager.get('excel_path', 'kb.xlsx'))
