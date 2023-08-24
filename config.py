import iniconfig


class ConfigManager:
    ini = iniconfig.IniConfig("kb.ini")

    @classmethod
    def get(cls, key, default, section='basic'):
        return cls.ini.get(section, key, default)


if __name__ == '__main__':
    print(ConfigManager.get('excel_path','kb.xlsx'))
