from abc import ABC
from dataclasses import dataclass


@dataclass
class Attachment(ABC):
    name: str
    height_in_inches: int

    def __lt__(self, other):
        if isinstance(other, Attachment):
            return self.height_in_inches < other.height_in_inches

    def __le__(self, other):
        if isinstance(other, Attachment):
            return self.height_in_inches <= other.height_in_inches

    def __gt__(self, other):
        if isinstance(other, Attachment):
            return self.height_in_inches > other.height_in_inches

    def __ge__(self, other):
        if isinstance(other, Attachment):
            return self.height_in_inches >= other.height_in_inches


class Power(Attachment):
    pass


class Comm(Attachment):
    pass


class Streetlight(Attachment):
    pass
