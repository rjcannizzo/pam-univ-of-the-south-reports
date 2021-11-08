import logging
import math

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
s_handler = logging.StreamHandler()
f_handler = logging.FileHandler('logs/app.log')
f_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
s_handler.setFormatter(formatter)
f_handler.setFormatter(formatter)
logger.addHandler(s_handler)
logger.addHandler(f_handler)

def parse_int(value, default=None):
    try:
        return math.floor(float(value))
    except ValueError:
        logger.error(f"could not parse as int {value}")
        return value

def parse_float(value, default=None):
    try:
        return float(value)
    except ValueError:
        return value

def parse_string(value, default=None):
    if not value:
        return default
    try:
        return value.strip()
    except (ValueError, AttributeError):
        return value