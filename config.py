import iniconfig


class ConfigManager:
    ini = iniconfig.IniConfig("kb.ini")

    @classmethod
    def get(cls, key, default, section='basic'):
        return cls.ini.get(section, key, default)

    @classmethod
    def excel_path(cls):
        return cls.get('excel_path', 'kb.xlsx')

    @classmethod
    def shortcut(cls, key, default=None):
        return cls.get(key, default, 'shortcuts')


if __name__ == '__main__':
    print(ConfigManager.get('excel_path', 'kb.xlsx'))
