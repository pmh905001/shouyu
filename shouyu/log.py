import logging


class Log:
    FORMAT = '%(asctime)s %(levelname)s %(filename)-8s: %(lineno)s line -%(message)s'

    @classmethod
    def setup(cls):
        logging.basicConfig(
            level=logging.INFO,
            filename='kb.log',
            format=cls.FORMAT,
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        logger = logging.getLogger()
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter(cls.FORMAT))
        logger.addHandler(console_handler)
