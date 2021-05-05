

import subprocess
import time

import psutil
import uiautomation
import win32gui
import win32process
from uiautomation import *

from .errors import ApplicationError


def controls_support(func):
    def wrapper(control, *args, **kwargs):
        if not control.isPidControl:
            return func(control, *args, **kwargs)
        else:
            # print(control.Controls)
            result = []
            for sub_control in control.Controls:
                result.append(func(sub_control, *args, **kwargs))
                # print("\n")
            return result
    return wrapper


class Application(uiautomation.Control):
    isPidControl = True
    pids_cache = []

    def __init__(self, exe_path, shell=False) -> None:
        self.exe_path = exe_path
        self.main_process = subprocess.Popen(exe_path, shell=shell)
        time.sleep(2)
        self.Pids

    @property
    def MainPid(self):
        return self.main_process.pid

    @property
    def Pids(self):
        try:
            parent = psutil.Process(self.MainPid)
            # print(self.MainPid)
            # print(parent)
        except psutil.NoSuchProcess:
            if self.pids_cache:
                remaining_pids = [x for x in self.pids_cache if psutil.pid_exists(x)]
                if remaining_pids:
                    return remaining_pids
            raise ApplicationError(f"'{self.exe_path}' and all children were terminated unexpectedly") from None
        all_process = [parent]
        children = parent.children(recursive=True)
        all_process.extend(children)
        self.pids_cache = [x.pid for x in all_process]
        return self.pids_cache

    @property
    def PidsControl(self):
        pass

    @property
    def MainControl(self):
        return GetControlFromPid(self.MainPid)

    @property
    def Controls(self):
        controls = [GetControlFromPid(x, ignore_error=True) for x in self.Pids]
        controls = [x for x in controls if x]
        return controls

    @staticmethod
    def get_handle_from_pid(pid):

        def callback(hwnd, hwnds):

            _, result = GetWindowThreadProcessId(hwnd)
            if result == pid:
                hwnds.append(hwnd)
            return True

        from win32gui import EnumWindows, IsWindowVisible
        from win32process import GetWindowThreadProcessId
        hwnds = []
        EnumWindows(callback, hwnds)
        return hwnds


def GetControlFromPid(pid, ignore_error=False):

    def get_hwnds(pid):
        """return a list of window handlers based on it process id"""
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
            return True
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if not hwnds:
            raise LookupError(f"Unable to find control with pid {pid}")
        return hwnds[0]

    try:
        return uiautomation.ControlFromHandle(get_hwnds(pid))
    except LookupError:
        if not ignore_error:
            raise RuntimeError(f"Unable to find control with pid {pid}") from None
        else:
            return None
