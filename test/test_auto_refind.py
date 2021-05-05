from time import time
import uiauto
from .conftest import Utils
import logging
logging.basicConfig(level=logging.DEBUG)


def test_layer1():
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    assert window.Exists()

    menu_bar = window.MenuBarControl(searchDepth=1)

    Utils.launch_notepad()

    assert menu_bar.Exists()


def test_layer2():
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    assert window.Exists()

    menu_bar = window.MenuBarControl(searchDepth=1)
    assert menu_bar.Exists()

    Utils.launch_notepad()

    assert menu_bar.Exists()


def test_layer3():
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    assert window.Exists()

    menu_bar = window.MenuBarControl(searchDepth=1)
    assert menu_bar.Exists()

    file_item = menu_bar.MenuItemControl()
    assert file_item.Exists()

    Utils.launch_notepad()

    assert file_item.Exists()
    assert file_item.GetExpandCollapsePattern().Expand()


def test_search_by_get_children():
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    assert window.Exists()

    menu_bar = window.MenuBarControl(searchDepth=1)
    assert menu_bar.Exists()

    file_item = menu_bar.GetChildren()[0]
    assert file_item.Exists()

    Utils.launch_notepad()

    assert file_item.Exists()


def test_search_by_get_children2():
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    assert window.Exists()

    child1 = window.GetChildren()[0]
    assert child1.Exists()

    child2 = child1.GetChildren()[1]
    assert child2.Exists()

    Utils.launch_notepad()

    assert child2.Exists()
    child2.Log()


def test_search_with_child_control():
    # ==== pre-check ====
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    exit_item = window.MenuItemControl(AutomationId="7")
    assert not exit_item.Exists()

    menu_bar = window.MenuBarControl(searchDepth=1)
    file_item = menu_bar.MenuItemControl()
    file_item.GetExpandCollapsePattern().Expand()

    assert exit_item.Exists()
    file_item.GetExpandCollapsePattern().Collapse()
    assert not exit_item.Exists(1)

    menu = window.MenuControl()
    assert not menu.Exists()

    # ==== search with child control ====
    menu = window.MenuControl(Child=uiauto.MenuItemControl(AutomationId="7"))
    assert not menu.Exists(1)

    file_item.GetExpandCollapsePattern().Expand()
    assert menu.Exists()


def test_negative_1():
    # ==== pre-check ====
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    menu_bar = window.MenuBarControl(searchDepth=1)
    file_item = menu_bar.MenuItemControl()
    assert file_item.Exists()

    file_item.GetExpandCollapsePattern().Expand()
    exit_item = window.MenuItemControl(AutomationId="7")
    assert exit_item.Exists()

    file_item.GetExpandCollapsePattern().Collapse()
    assert not exit_item.Exists(1)

    try:
        exit_item.Log()
    except Exception as e:
        assert isinstance(e, LookupError)
    else:
        raise AssertionError("No LookupError was raised while access a non-exists instance")


def test_negative_2():
    # ==== pre-check ====
    Utils.launch_notepad()

    window = uiauto.WindowControl(ClassName="Notepad")
    menu_bar = window.MenuBarControl(searchDepth=1)
    file_item = menu_bar.MenuItemControl()
    assert file_item.Exists()

    file_item.GetExpandCollapsePattern().Expand()
    exit_item = window.MenuItemControl(AutomationId="7")
    assert exit_item.Exists()

    Utils.kill_notepad()
    import time
    start_time = time.time()
    assert not exit_item.Exists(30)
    assert time.time() - start_time > 29


def teardown_function():
    Utils.kill_notepad()


def teardown_package():
    Utils.kill_notepad()
