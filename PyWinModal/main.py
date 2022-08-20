from PyWinModal.lib.extensions import *
from PyWinModal.lib.types import Action, Switch
from PyWinModal.lib.ui import Modal, OptionsList

from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)

modal = Modal(
    [
        OptionsList(
            "Example1",
            [
                Action(
                    "tab",
                    "Go to previous window",
                    activate_last_window()
                ),
                Switch(
                    "s",
                    "Example2"
                ), 
            ],
        ),
        OptionsList(
            "Example2",
            [
                Action("t", "New tab", send_keys_to_active_window("ctrl+t"), repeatable=True),
                Action("s", "Search Google", search_web("Google", "https://www.google.com/search?q=%s")),
                Switch("b", "Example1"),
            ],
        ),
    ],
    "Example1",
    "F24",
)
modal.run()