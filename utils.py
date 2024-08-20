Student = tuple[str, str, str]
Team = list[Student]

def ensure_list(value):
    if not value:
        return []
    elif isinstance(value, list):
        return value
    else:
        return [value]
