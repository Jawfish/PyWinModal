from typing import Callable
from pywinauto import Application
import win32gui
import win32api
import win32con
import keyboard
import win32process
import psutil
import PySimpleGUI as sg
import os
import winreg
import win32clipboard
import ctypes


class Utils:
    @staticmethod
    def get_window_at_index(index: int) -> int:
        """Returns the window handle at the given EnumWindows index."""
        windows: list[int] = []

        def winEnumHandler(hwnd, ctx) -> None:
            nonlocal windows
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):  # type: ignore
                windows.append(hwnd)

        win32gui.EnumWindows(winEnumHandler, None)  # type: ignore
        return windows[index]

    @staticmethod
    def get_window_from_exe(
        exe: str,
    ) -> int | None:
        """Returns the window handle of the next window with the given exe name."""
        window: int | None = None
        exe = exe.split(".")[0].lower()

        def winEnumHandler(hwnd, ctx) -> None:
            if win32gui.GetWindowText(hwnd) and win32gui.IsWindowVisible(hwnd):  # type: ignore
                nonlocal window
                nonlocal exe
                tid, pid = win32process.GetWindowThreadProcessId(hwnd)  # type: ignore
                window_path = psutil.Process(pid).exe()
                process_exe = window_path.split("\\")[-1].split(".")[0].lower()
                if exe in process_exe:
                    window = hwnd

        win32gui.EnumWindows(winEnumHandler, None)  # type: ignore
        return window

    @staticmethod
    def get_window_from_partial_title(title: str) -> int | None:
        """Returns the window handle of the next window that contains the given string."""
        window: int | None = None

        def winEnumHandler(hwnd, ctx) -> None:
            nonlocal window
            if title.lower() in win32gui.GetWindowText(hwnd).lower() and not window:  # type: ignore
                window = hwnd

        win32gui.EnumWindows(winEnumHandler, None)  # type: ignore
        return window

    @staticmethod
    def get_active_window() -> int:
        """Returns the window handle of the window that was active before the modal was summoned."""
        return Utils.get_window_at_index(0)

    @staticmethod
    def get_text_input(search_entity: str = "") -> str:
        layout = [
            [
                sg.Text(
                    f"Search {search_entity}",
                )
            ],
            [
                sg.InputText(
                    focus=True,
                    key="input",
                )
            ],
            [sg.Submit(size=(0, 0))],
        ]
        window = sg.Window(
            "",
            layout,
            finalize=True,
            no_titlebar=True,
        )
        window.bring_to_front()
        window.move_to_center()
        window.force_focus()
        window.Element("input").SetFocus()

        event, values = window.read()
        window.close()

        text_input = values["input"]
        return text_input

    @staticmethod
    def move_foreground_window_to_center() -> Callable:
        """Moves the foreground window to the center of the screen."""

        def action():
            window = Utils.get_window_from_partial_title("PyWinModal")
            screen_width = win32api.GetSystemMetrics(0)  # type: ignore
            screen_height = win32api.GetSystemMetrics(1)  # type: ignore
            window_rect = win32gui.GetWindowRect(window)  # type: ignore
            window_width = window_rect[2] - window_rect[0]
            window_height = window_rect[3] - window_rect[1]
            win32gui.MoveWindow(
                window,  # type: ignore
                screen_width // 2 - window_width // 2,
                screen_height // 2 - window_height // 2,
                0,
                0,
                True,
            )

        return action

    @staticmethod
    def raise_window(hwnd: int) -> None:
        """Brings the window with the given handle to the foreground."""
        title = win32gui.GetWindowText(hwnd)  # type: ignore
        # if window is minimized, restore it
        if win32gui.IsIconic(hwnd):  # type: ignore
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # type: ignore
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)  # type: ignore
        win32gui.BringWindowToTop(hwnd)  # type: ignore
        win32gui.SetForegroundWindow(hwnd)  # type: ignore
        win32gui.SetActiveWindow(hwnd)  # type: ignore
        print(f"Raised window {title} with handle {hwnd}")


def close_active_window() -> Callable:
    """Closes the foreground window."""
    return lambda: win32gui.PostMessage(Utils.get_active_window(), win32con.WM_CLOSE, 0, 0)  # type: ignore


def activate_last_window() -> Callable:
    """Brings the last active window to the foreground."""
    return lambda: win32gui.SetForegroundWindow(Utils.get_window_at_index(1))  # type: ignore


def open_application(application: str) -> Callable:
    """Launches the given application."""
    return lambda: Application().start(application)


def summon_or_launch(application: str) -> Callable:
    """Brings the given application to the foreground or launches it if it is not running."""
    def action() -> None:
        # extract just the executable from the path
        executable: str = application.split("\\")[-1]
        # if the application is running, bring it to the foreground
        window = Utils.get_window_from_exe(executable)
        if window:
            print(f"Trying to bring {executable} to the foreground")
            win32gui.SetForegroundWindow(window)  # type: ignore
        else:
            # otherwise, launch the application
            print(f"Trying to launch {executable}")
            Application().start(application)

    return action


def go_to_window_by_title(title: str) -> Callable:
    """Brings the window with the given title to the foreground."""
    def action() -> None:
        try:
            window = Utils.get_window_from_partial_title(title)
            if window:
                Utils.raise_window(window)
        except:
            pass

    return action


def go_to_window_by_exe(exe: str) -> Callable:
    """Brings the window with the given executable to the foreground."""
    def action() -> None:
        try:
            window = Utils.get_window_from_exe(exe)
            if window:
                Utils.raise_window(window)
        except:
            pass

    return action


def send_keys_to_active_window(keys: str) -> Callable:
    """Sends the given keys to the foreground window."""
    def action() -> None:
        print(f"sending {keys} to {win32gui.GetWindowText(win32gui.GetForegroundWindow())}")  # type: ignore
        keyboard.press_and_release(keys)

    return action


def send_keys_to_window_by_title(title: str, keys: str) -> Callable:
    """Sends the given keys to the window with the given title."""
    def action() -> None:
        win32gui.SetForegroundWindow(Utils.get_window_from_partial_title(title))  # type: ignore
        keyboard.press_and_release(keys)

    return action


def search_web(title: str, url: str) -> Callable:
    """Open a web search in the default browser. %s in the url is replaced with the query."""

    def action() -> None:
        query: str = Utils.get_text_input(title)
        url_with_query: str = url.replace("%s", query)
        os.startfile(url_with_query)

    return action


def toggle_dark_mode() -> Callable:
    """Toggles the Windows 10 and Windows 11 dark mode."""
    def action() -> None:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, winreg.KEY_ALL_ACCESS)  # type: ignore
        value = winreg.QueryValueEx(key, "SystemUsesLightTheme")[0]  # type: ignore
        if value == 1:
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, 0)
        else:
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, 1)

    return action


def lock_workstation() -> Callable:
    """Locks the workstation. Equivalent to pressing Win+L."""
    def action() -> None:
        ctypes.windll.user32.LockWorkStation()

    return action


def empty_clipboard() -> Callable:
    """Empties the clipboard."""
    def action() -> None:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()

    return action


def minimize_active_window() -> Callable:
    """Minimizes the foreground window."""
    def action() -> None:
        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE)  # type: ignore

    return action


def maximize_active_window() -> Callable:
    """Maximizes the foreground window."""
    def action() -> None:
        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)  # type: ignore

    return action
