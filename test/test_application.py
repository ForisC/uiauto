import time
from uiautomation.uiautomation import MenuBarControl
import uiauto
from .conftest import Utils
import psutil
from uiauto import ApplicationError


def test_open_app():
    app = uiauto.Application("notepad.exe")
    assert app.Exists()

    all_notepad_pids = Utils.get_image_pids("notepad.exe")
    assert len(all_notepad_pids) == 1


def test_main_pid():
    app = uiauto.Application("notepad.exe")
    all_notepad_pids = Utils.get_image_pids("notepad.exe")
    assert app.MainPid in all_notepad_pids


def test_all_pids():
    app = uiauto.Application("cmd /c start /wait notepad.exe")

    all_cmd_pids = Utils.get_image_pids("cmd.exe")
    cmd_pid = app.Pids[0]
    assert cmd_pid in all_cmd_pids

    all_notepad_pids = Utils.get_image_pids("notepad.exe")
    notepad_pid = app.Pids[1]
    assert notepad_pid == all_notepad_pids[0]


def test_access_child():
    app = uiauto.Application("cmd /c start /wait notepad.exe")
    notepad_search_by_app = app.WindowControl(ClassName="Notepad")
    assert notepad_search_by_app.Exists()

    notepad_search_directly = uiauto.WindowControl(ClassName="Notepad")
    assert notepad_search_directly.Exists()

    assert notepad_search_by_app.NativeWindowHandle == notepad_search_directly.NativeWindowHandle


def test_access_child_with_same_app():
    Utils.launch_notepad()
    original_notepad = uiauto.WindowControl(ClassName="Notepad")
    assert original_notepad.Exists()

    app = uiauto.Application("cmd /c start /wait notepad.exe")
    child_notepad = app.WindowControl(ClassName="Notepad")

    assert original_notepad.NativeWindowHandle != child_notepad.NativeWindowHandle
    assert original_notepad.ProcessId != child_notepad.ProcessId

    all_notepad_pids = Utils.get_image_pids("notepad.exe")
    assert sorted(all_notepad_pids) == sorted([original_notepad.ProcessId, child_notepad.ProcessId])


def test_terminate():
    app = uiauto.Application("notepad.exe")
    menu = app.MenuBarControl(searchDepth=1)
    assert menu.Exists()

    psutil.Process(app.MainPid).terminate()
    time.sleep(2)
    try:
        menu.Exists()
    except Exception as e:
        assert isinstance(e, ApplicationError)
    else:
        raise AssertionError("No ApplicationError was raised while accessing a terminated Application")


def test_terminate_main_with_cache_child():
    app = uiauto.Application("cmd /c start /wait notepad.exe")
    menu = app.MenuBarControl(searchDepth=1)
    assert menu.Exists()

    psutil.Process(app.MainPid).terminate()
    time.sleep(2)

    all_notepad_pids = Utils.get_image_pids("notepad.exe")
    assert app.Pids == all_notepad_pids
    assert menu.Exists()
    assert app.WindowControl(ClassName="Notepad").Exists()


def teardown_function():
    Utils.kill_notepad()


def setup_function():
    Utils.kill_notepad()
