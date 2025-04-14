from datetime import datetime

def format_date(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    elif isinstance(value, str):
        try:
            parsed_date = datetime.strptime(value, "%Y-%m-%d")
            return parsed_date.strftime("%Y-%m-%d")
        except ValueError:
            return value
    else:
        return value