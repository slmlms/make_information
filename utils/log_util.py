import datetime
import os

from loguru import logger


@logger.catch
def to_log(config, type_name):
    level = config.get('logger', 'level')
    format = config.get('logger', 'format_' + level.lower())
    today = str(datetime.datetime.today().strftime('%Y-%m-%d'))
    log_dir = config.get('logger', 'logDir') + '/' + type_name + '/' + today
    if not os.path.dirname(log_dir): os.mkdir(log_dir)
    if level.lower() == 'debug':
        logger.add(log_dir + "/" + level + "_{time}.log", format=format, level=level, enqueue=False)
    elif level.lower() == 'success':
        logger.add(log_dir + "/" + level + "_{time}.log", format=format, level=level)
