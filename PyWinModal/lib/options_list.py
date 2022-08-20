from PyWinModal.lib.types import Action, Switch
import PySimpleGUI as sg


class OptionsList:
    def __init__(self, title: str, options: list[Action | Switch]) -> None:
        self.title = title
        self.actions: dict[str, Action] = {}
        self.switches: dict[str, Switch] = {}
        self.layout: list[list[sg.Text]] = []
        for option in options:
            if isinstance(option, Action):
                self.add_action(option)
            elif isinstance(option, Switch):
                self.add_switch(option)

    def add_action(self, action: Action) -> None:
        if self.actions and self.check_if_option_exists(action.hotkey):
            e = f"Hotkey {action.hotkey} already exists for {self.title}"
            raise ValueError(e)

        self.actions[action.hotkey] = action

        self.layout.append(
            [
                sg.Text(
                    f"{action.hotkey.title()} - {action.title}{' ⟳' if action.repeatable else ''}",
                    font=("Open Sans", 24, ""),
                    background_color="black",
                )
            ]
        )

    def add_switch(self, switch: Switch) -> None:
        if self.switches and self.check_if_option_exists(switch.hotkey):
            e = f"Hotkey {switch.hotkey} already exists for {self.title}"
            raise ValueError(e)

        self.switches[switch.hotkey] = switch

        self.layout.append(
            [
                sg.Text(
                    f"{switch.hotkey.title()} - {switch.title}...",
                    font=("Open Sans", 24, ""),
                    background_color="black",
                )
            ]
        )

    def check_if_option_exists(self, hotkey: str) -> bool:
        return hotkey in self.actions or hotkey in self.switches
