import logging


def setup_logger(logfile):
    logger = logging.getLogger(logfile)
    # Check if handlers are already added to the logger
    if len(logger.handlers) == 0:
        logger.setLevel(logging.DEBUG)
        handler = logging.FileHandler(logfile)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger
