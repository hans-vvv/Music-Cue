import logging

logger = logging


def set_logger(filename: str) -> None:
    logger.basicConfig(
        filename=filename, filemode='a', format='%(asctime)s - %(levelname)s - %(message)s', level=logging.ERROR
    )


def get_logger():
    return logger



