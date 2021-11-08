def parse_int(value, default=None):
    try:
        return int(value)
    except ValueError:
        return default

def parse_float(value, default=None):
    try:
        return float(value)
    except ValueError:
        return default

def parse_string(value, default=None):
    if not value:
        return default
    try:
        return value.strip()
    except (ValueError, AttributeError):
        return default