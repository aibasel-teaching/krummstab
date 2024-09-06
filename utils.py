import json
from pathlib import Path


class Student:
    def __init__(self, firstname, lastname, email):
        self.firstname = firstname
        self.lastname = lastname
        self.email = email

    def __eq__(self, other: object) -> bool:
        return (self.lastname, self.firstname, self.email) == (other.lastname, other.firstname, other.email)

    def __lt__(self, other: object) -> bool:
        return (self.lastname, self.firstname, self.email) < (other.lastname, other.firstname, other.email)

    def __hash__(self) -> int:
        return hash((self.lastname, self.firstname, self.email))

    @property
    def full_name(self):
        return f"{self.firstname} {self.lastname}"


class Team:
    def __init__(self, students: list[Student]) -> None:
        self.students = tuple(sorted(students))

    def __eq__(self, other: object) -> bool:
        return self.students == other.students

    def __lt__(self, other: object) -> bool:
        return self.students < other.students

    def __hash__(self) -> int:
        return hash(self.students)

    def __str__(self) -> str:
        return self.name

    def __contains__(self, student) -> bool:
        return student in self.students

    def __getitem__(self, index) -> Student:
        return self.students[index]
    
    def __len__(self) -> int:
        return len(self.students)

    @property
    def name(self) -> str:
        """
        Concatenate the last names of students to get a pretty-ish string
        representation of teams.
        """
        return "_".join(sorted([student.lastname.replace(" ", "-") for student in self]))


class AdamScriptEncoder(json.JSONEncoder):
    """Custom JSON Encoder that can handle our classes Team and Student, as well
    as pathlib.Path objects."""
    def default(self, obj):
        if isinstance(obj, Path):
            return {"_type": "pathlib.Path", "value": str(obj)}
        elif isinstance(obj, Student):
            return {"_type": "Student", "value": [obj.firstname, obj.lastname, obj.email]}
        elif isinstance(obj, Team):
            return {"_type": "Team", "value": obj.students}
        return super().default(obj)

    def replace_path_keys(self, obj):
        # Path objects cannot be used as keys, so we flatten dicts using them as keys to lists.
        if isinstance(obj, dict):
            if any(isinstance(key, Path) for key in obj):
                return {"_type": "FlatDict", "value": [self.replace_path_keys(entry) for entry in obj.items()]}
            else:
                return {key: self.replace_path_keys(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [self.replace_path_keys(entry) for entry in obj]
        elif isinstance(obj, tuple):
            return tuple(self.replace_path_keys(entry) for entry in obj)
        else:
            return obj

    def encode(self, obj):
        return super().encode(self.replace_path_keys(obj))

class AdamScriptDecoder(json.JSONDecoder):
    """Custom JSON Decoder that can handle our classes Team and Student, as well
    as pathlib.Path objects."""
    def __init__(self, *args, **kwargs):
        super().__init__(object_hook=self.object_hook, *args, **kwargs)

    def object_hook(self, obj):
        if "_type" in obj:
            if obj["_type"] == "pathlib.Path":
                return Path(obj["value"])
            elif obj["_type"] == "Student":
                return Student(*obj["value"])
            elif obj["_type"] == "Team":
                return Team(obj["value"])
            elif obj["_type"] == "FlatDict":
                return dict(obj["value"])
        return obj


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"), cls=AdamScriptDecoder)


def write_json(path: Path, data):
    data_str = json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True, cls=AdamScriptEncoder)
    path.write_text(data_str, encoding="utf-8")


def ensure_list(value):
    if not value:
        return []
    elif isinstance(value, list):
        return value
    else:
        return [value]
