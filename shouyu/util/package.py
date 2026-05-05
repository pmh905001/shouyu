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
        # In a PyInstaller-frozen exe we want cwd == exe directory so that
        # relative reads (kb.ini) and writes (kb.log, pid.txt, backups…) all
        # land alongside the binary instead of wherever the user launched us
        # from (System32 if from Win+R, the user's home if from cmd, …).
        #
        # We use sys.executable here, NOT sys.argv[0]: when the user invokes
        # `shouyu hello` via PATH, argv[0] is just the bare name and
        # os.path.dirname returns "", which makes os.chdir raise
        # `OSError: [WinError 123] 文件名、目录名或卷标语法不正确。`.
        # sys.executable is documented to always be the full exe path under
        # PyInstaller.
        if hasattr(sys, '_MEIPASS'):
            target = os.path.dirname(os.path.abspath(sys.executable))
            if target:
                os.chdir(target)
