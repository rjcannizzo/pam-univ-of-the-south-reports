import logging


def get_logger(log_file, logger_name, logger_level, file_handler_level, stream_handler=False):
    """
    Returns a logger with an optional Streamhandler.
    :param log_file: full path to the log file to write.
    :param logger_name: the name for this logger (usually '__name__')
    :param logger_level: The Lowest level. Set this or you won't get messages if they're lower than the default!
    :param file_handler_level: The level to set for the log file (e.g. logger.INFO)
    :param stream_handler: Bool. Activates a Streamhandler if True.
    :return: logger
    """
    logger = logging.getLogger(logger_name)
    logger.setLevel(logger_level)    
    f_handler = logging.FileHandler(log_file)
    f_handler.setLevel(file_handler_level)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')    
    f_handler.setFormatter(formatter)
    logger.addHandler(f_handler)    
    if stream_handler:
        s_handler = logging.StreamHandler()
        s_handler.setFormatter(formatter)
        logger.addHandler(s_handler)

    return logger
        