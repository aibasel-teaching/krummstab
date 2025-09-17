class Student:
    def __init__(self, first_name: str, last_name: str, email: str):
        self.first_name = first_name
        self.last_name = last_name
        self.email = email

    def __eq__(self, other) -> bool:
        return self.email == other.email

    def __lt__(self, other) -> bool:
        return self.email < other.email

    def __format__(self, spec) -> str:
        """
        Get a pretty printed string representation of a student:
        first_name last_name (email)
        """
        return f"{self.first_name} {self.last_name} ({self.email})"

    def to_tuple(self) -> tuple[str, str, str]:
        """
        Get a tuple of strings representation of a student.
        """
        return self.first_name, self.last_name, self.email
