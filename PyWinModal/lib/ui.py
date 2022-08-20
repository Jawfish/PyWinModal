import time
import PySimpleGUI as sg
from psgtray import SystemTray
import keyboard
import os
from PIL import Image, ImageDraw
import win32gui
from pywinauto import Application
from PyWinModal.lib.options_list import OptionsList
import os


class Modal:
    def __init__(
        self,
        option_lists: list[OptionsList],
        default_list: str,
        leader_key: str = "F24",
    ) -> None:
        self.option_lists: list[OptionsList] = []
        self.option_layouts = [
            [sg.VPush(background_color="black")],  # for vertical alignment
            [],  # option list appends here
            [sg.VPush(background_color="black")],  # for vertical alignment
        ]
        self.showing: bool = False
        self.leader_key = leader_key
        self.current_layout: str = default_list
        self.app = Application(backend="uia").connect(process=os.getpid())
        self.current_window: int = 0  # active window before the modal was summoned
        keyboard.hook_key(self.leader_key, self.handle_leader, suppress=False)

        # add option list layouts
        for option_list in option_lists:
            self.option_lists.append(option_list)
            self.option_layouts[1].append(
                sg.Column(
                    option_list.layout,
                    key=option_list.title,
                    visible=False,
                    background_color="black",
                )
            )  # we append to the first column, hence the [0]

        # set up window
        self.window = sg.Window(
            "PyWinModal",
            self.option_layouts,
            no_titlebar=False,  # titlebar is necessary for focus grabbing
            keep_on_top=True,
            return_keyboard_events=True,
            use_default_focus=False,
            alpha_channel=0.8,
            background_color="black",
            finalize=True,
            element_justification="center",
        )
        self.window.hide()  # don't show until leader key is pressed

        # set up system tray icon
        tray_layout = [
            "",
            [
                "Exit",  # adds exit to the tray's right-click menu
            ],
        ]
        # add menu items to tray right-click menu
        for option_list in self.option_lists:
            tray_layout[1].insert(
                -1, option_list.title
            )  # -1 to keep the tray menu alphabetical with Exit at the end

        self.tray = SystemTray(
            tray_layout,
            single_click_events=False,
            window=self.window,
            tooltip="PyWinModal",
            icon=sg.DEFAULT_BASE64_ICON,
        )

        self.set_layout(default_list)

    def set_layout(self, layout_title: str) -> None:
        self.window[self.current_layout].update(visible=False)
        self.current_layout = layout_title
        self.window[self.current_layout].update(visible=True)

    def toggle_visibility(self) -> None:
        if self.showing:
            self.showing = False
            self.window.Hide()
            self.set_layout(self.option_lists[0].title)  # reset to default layout
        else:
            self.window.Maximize()
            handle = win32gui.FindWindow(None, "PythonModal")  # type: ignore
            self.showing = True
            self.window.UnHide()
            # doing this through win32gui is more consistent than PySimpleGUI
            try:
                win32gui.BringWindowToTop(handle)  # type: ignore
                win32gui.SetForegroundWindow(handle)  # type: ignore
                win32gui.SetActiveWindow(handle)  # type: ignore
            except win32gui.error:
                pass
            self.app.top_window().set_focus()

    def set_current_window(self):
        self.current_window = win32gui.GetForegroundWindow()

    def run(self):
        while True:
            event, values = self.window.read()
            if event == "	":
                event = "tab"

            if event == self.tray.key:
                event = values[event]
                if event == "__DOUBLE_CLICKED__":
                    self.toggle_visibility()

            # allow layout to be changed from tray icon menu
            for option_list in self.option_lists:
                if event == option_list.title:
                    self.set_layout(event)

            if event in (sg.WIN_CLOSED, "Exit"):
                break

            # evaluate key press events
            # this could be cleaned up a lot by consolidating things into a dict
            # performance isn't likely an issue though given the bottleneck is human input
            known_hotkey = False
            for option in self.option_lists:
                if option.title == self.current_layout:
                    # if the pressed key is associated with an action, execute the action's callback
                    for hotkey in option.actions:
                        if str(event).lower() == hotkey.lower():
                            known_hotkey = True
                            # if the action is not repeatable, hide the window
                            if not option.actions[hotkey].repeatable:
                                self.toggle_visibility()
                                option.actions[hotkey].callback()
                            # otherwise, active the window that was active before the modal was summoned
                            else:
                                modal = win32gui.GetForegroundWindow()
                                win32gui.SetForegroundWindow(self.current_window)  # type: ignore
                                win32gui.SetActiveWindow(self.current_window)  # type: ignore
                                option.actions[hotkey].callback()
                                # a delay is necessary to prevent the modal from stealing focus too soon
                                # needs to be relatively long to allow for key combos
                                # TODO: not sure why this is necessary, maybe keyboard.press_and_send is non-blocking
                                time.sleep(0.05)
                                # occasionally the modal won't be properly brought to the foreground, so run this in a loop
                                while win32gui.GetForegroundWindow() != modal:
                                    win32gui.SetForegroundWindow(modal)  # type: ignore
                            break
                    # if the hotkey wasn't associated with an action, check if it's a switch
                    if not known_hotkey:
                        for hotkey in option.switches:
                            if str(event).lower() == hotkey.lower():
                                known_hotkey = True
                                self.set_layout(option.switches[hotkey].title)
                                break
                        break
                    break

            # hide the modal if the key wasn't associated with an action or switch
            if not known_hotkey and not self.leader_key in event:
                self.toggle_visibility()

        self.close_program()

    def handle_leader(self, event: keyboard.KeyboardEvent) -> None:
        if event.event_type == "down":
            keyboard.block_key(self.leader_key)
            self.set_current_window()
            self.toggle_visibility()
            keyboard.unblock_key(self.leader_key)

    def close_program(self):
        self.tray.close()
        self.window.close()
        os._exit(0)

    def tray_notify(self, title: str, message: str):
        self.tray.show_message(title, message)

    # create tray icon
    def create_image(self, width, height, color1, color2):
        image = Image.new("RGB", (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
        dc.rectangle((0, height // 2, width // 2, height), fill=color2)
        return image
