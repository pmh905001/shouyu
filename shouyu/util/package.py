import os
import sys


class Package:
    @staticmethod
    def get_resource_path(relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath('.'), relative_path)

    @staticmethod
    def set_cwd():
        # if hasattr(sys, 'frozen') and sys.frozen:
        if hasattr(sys, '_MEIPASS'):
            os.chdir(os.path.dirname(sys.argv[0]))
