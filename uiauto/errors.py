import sys
from uiautomation import GetRootControl
from .utils import embedded_image


class PatternError(RuntimeError):
    __module__ = Exception.__module__


class ApplicationError(RuntimeError):
    __module__ = Exception.__module__


class UiLookupError(LookupError):
    __module__ = Exception.__module__
    ROBOT_SUPPRESS_NAME = True

    def __init__(self, message, filename="UiLookupError.jpg"):
        self.message = message.replace("*HTML*", "")
        self.filename = GetRootControl().Screenshot(filename=filename, log=False)

    def __str__(self):
        if "robot" in sys.modules:
            msg = (f"*HTML*{self.message}\n"
                   f"Screenshot:\n{embedded_image(self.filename)}")
        else:
            msg = (f"{self.message}\n"
                   f"Screenshot: {self.filename}")
        return msg
