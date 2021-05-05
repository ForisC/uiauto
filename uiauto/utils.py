import time
import win32gui
import win32ui  # pylint: disable=reportMissingImports
from win32api import GetSystemMetrics


def draw_rectangle(x1, y1, x2, y2):
    width = 5
    dc = win32gui.GetDC(0)
    hwnd = win32gui.WindowFromPoint((0, 0))

    dcObj1 = win32ui.CreateDCFromHandle(dc)
    dcObj2 = win32ui.CreateDCFromHandle(dc)
    dcObj3 = win32ui.CreateDCFromHandle(dc)
    dcObj4 = win32ui.CreateDCFromHandle(dc)

    monitor = (0, 0, GetSystemMetrics(0), GetSystemMetrics(1))
    brush = win32ui.CreateBrush(0, 0x00FF00, 0)

    start_time = time.time()
    # win32gui.InvalidateRect(hwnd, monitor, False)

    win32gui.InvalidateRect(None, monitor, True)
    while time.time() - start_time < 5:
        dcObj1.FillRect((x1, y1, x2, y1 + width), brush)
        dcObj2.FillRect((x1, y1, x1 + width, y2), brush)
        dcObj3.FillRect((x1, y2, x2, y2 - width), brush)
        dcObj4.FillRect((x2, y1, x2 - width, y2), brush)
    win32gui.InvalidateRect(hwnd, monitor, True)


def embedded_image(filename: str) -> str:
    return f"<a href={filename} target=\"_blank\"><img src=\"{filename}\" width=\"200\"></a>"


def toggle_state(value):
    if str(value).lower() in ["on", "check", "checked", "yes", "y", "true", "1"]:
        return "On"
    elif str(value).lower() in ["off", "uncheck", "unchecked", "no", "n", "false", "0"]:
        return "Off"
    elif str(value).lower() in ["indeterminate", "2"]:
        return "Indeterminate"
    else:
        raise ValueError(f"Can not identify the state '{value}'")
