import logging


def setup_log():
    logging.basicConfig(
        level=logging.INFO,
        filename='../kb.log',
        format="%(asctime)s %(levelname)s %(filename)-8s: %(lineno)s line -%(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    logger = logging.getLogger()
    logger.addHandler(logging.StreamHandler())
    logger.addHandler(logging.FileHandler('../kb.log'))

