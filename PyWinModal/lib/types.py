from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class Action:
    hotkey: str
    title: str
    callback: Callable
    repeatable: bool = False


@dataclass(frozen=True)
class Switch:
    hotkey: str
    title: str
