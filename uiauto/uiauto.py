import collections
from ctypes import DEFAULT_MODE
import logging
import os
from datetime import datetime
from pathlib import Path

import comtypes
import uiautomation
from uiautomation import *  # pylint: disable=unused-wildcard-import
from uiautomation.uiautomation import _AutomationClient

from .controls import *  # pylint: disable=unused-wildcard-import
from .errors import ApplicationError, PatternError, UiLookupError
from .utils import draw_rectangle, embedded_image, toggle_state

comtypes.logger.disabled = True

ENABLE_SCREENSHOT = True
DEBUG_MODE = False
DEFAULT_LOG_PATH = Path(os.getcwd()) / "LOGS"


def SetLogDir(value):
    global LOG_DIR
    LOG_DIR = Path(value)
    if not LOG_DIR.is_dir():
        LOG_DIR.mkdir(parents=True)
    uiautomation.Logger.SetLogFile(str(LOG_DIR / "@AutomationLog.txt"))


SetLogDir(DEFAULT_LOG_PATH)


def EnableDebugMode(value):
    global DEBUG_MODE
    DEBUG_MODE = value


def EnableScreenshot(value):
    global ENABLE_SCREENSHOT
    ENABLE_SCREENSHOT = value


def Screenshot(self=uiautomation.GetRootControl(), filename=f"{time.strftime(r'%Y%m%d-%H%M%S')}.jpg", log=True):
    if not ENABLE_SCREENSHOT:
        return "Screenshot disabled"
    file_path = LOG_DIR / filename
    stem = file_path.stem
    suffix = file_path.suffix

    num = 1
    while file_path.is_file():
        file_path = LOG_DIR / f"{stem}-{num}{suffix}"
        num += 1

    rect = self.Element.CurrentBoundingRectangle
    x = rect.left
    y = rect.top
    w = rect.right - rect.left
    h = rect.bottom - rect.top

    self.CaptureToImage(str(file_path), x, y, w, h)
    logging.debug(f"Screenshot to file: {file_path}")
    try:
        relative_path = file_path.relative_to(os.getcwd())
    except ValueError:
        relative_path = file_path

    if log:
        if "robot" in sys.modules:
            from robot.api import logger
            logger.info(f"Screenshot to file {file_path}\n"
                        f"{embedded_image(relative_path)}", html=True)
        else:
            logging.info(f"Screenshot to file {file_path}")

    return relative_path


def patched_Refind(self, maxSearchSeconds: float = TIME_OUT_SECOND, searchIntervalSeconds: float = SEARCH_INTERVAL, raiseException: bool = True) -> bool:
    """
    Refind the control every searchIntervalSeconds seconds in maxSearchSeconds seconds.
    maxSearchSeconds: float.
    searchIntervalSeconds: float.
    raiseException: bool, if True, raise a LookupError if timeout.
    Return bool, True if find.
    """
    if not self.Exists(maxSearchSeconds, searchIntervalSeconds, False if raiseException else DEBUG_EXIST_DISAPPEAR):
        if raiseException:
            Logger.ColorfullyLog('<Color=Red>Find Control Timeout: </Color>' + self.GetColorfulSearchPropertiesStr())
            raise UiLookupError('Find Control Timeout: ' + self.GetSearchPropertiesStr())
        else:
            return False
    return True


def patched_Element(self):
    """
    Property Element.
    Return `ctypes.POINTER(IUIAutomationElement)`.
    """
    if not self._element:
        self.Refind(maxSearchSeconds=TIME_OUT_SECOND, searchIntervalSeconds=self.searchInterval)

    try:
        _AutomationClient.instance().ViewWalker.GetFirstChildElement(self._element)
    except comtypes.COMError:
        logging.info(f"{self.GetSearchPropertiesStr()} fail to access instance, try to refind")
        self._element = 0
        self.Refind(maxSearchSeconds=TIME_OUT_SECOND, searchIntervalSeconds=self.searchInterval, raiseException=False)

    return self._element


def patched_WalkControl(control: Control, includeTop: bool = False, maxDepth: int = 0xFFFFFFFF):
    """
    control: `Control` or its subclass.
    includeTop: bool, if True, yield (control, 0) first.
    maxDepth: int, enum depth.
    Yield 2 items tuple (control: Control, depth: int).
    """
    if includeTop:
        yield control, 0
    if maxDepth <= 0:
        return

    if not control.isPidControl:
        depth = 0
        control.searchByWalk = True
        child = control.GetFirstChildControl()
        controlList = [child]
        while depth >= 0:
            lastControl = controlList[-1]
            if lastControl:
                yield lastControl, depth + 1
                lastControl.searchByWalk = True
                child = lastControl.GetNextSiblingControl()
                controlList[depth] = child
                if depth + 1 < maxDepth:
                    lastControl.searchByWalk = True
                    child = lastControl.GetFirstChildControl()
                    if child:
                        depth += 1
                        controlList.append(child)
            else:
                del controlList[depth]
                depth -= 1

    else:
        for sub_control in control.Controls:
            depth = 0
            yield sub_control, 0
            sub_control.searchByWalk = True
            child = sub_control.GetFirstChildControl()
            controlList = [child]
            while depth >= 0:
                lastControl = controlList[-1]
                if lastControl:
                    yield lastControl, depth + 1
                    lastControl.searchByWalk = True
                    child = lastControl.GetNextSiblingControl()
                    controlList[depth] = child
                    if depth + 1 < maxDepth:
                        lastControl.searchByWalk = True
                        child = lastControl.GetFirstChildControl()
                        if child:
                            depth += 1
                            controlList.append(child)
                else:
                    del controlList[depth]
                    depth -= 1


def patched_CompareFunction(self, control: 'Control', depth: int) -> bool:
    """
    Define how to search.
    control: `Control` or its subclass.
    depth: int, tree depth from searchFromControl.
    Return bool.
    """

    for key, value in self.searchProperties.items():
        if 'ControlType' == key:
            if value != control.ControlType:
                return False
        elif 'ClassName' == key:
            if value != control.ClassName:
                return False
        elif 'AutomationId' == key:
            if value != control.AutomationId:
                return False
        elif 'Depth' == key:
            if value != depth:
                return False
        elif 'Name' == key:
            if value != control.Name:
                return False
        elif 'SubName' == key:
            if value not in control.Name:
                return False
        elif 'RegexName' == key:
            if not self.regexName.match(control.Name):
                return False
        elif 'Compare' == key:
            if not value(control, depth):
                return False
        elif 'Child' == key:
            child_control = FindControl(control, value._CompareFunction, value.searchDepth, False, value.foundIndex)
            if not child_control:
                return False
    return True


def patched_GetFirstChildControl(self) -> 'Control':
    """
    Return `Control` subclass or None.
    """
    ele = _AutomationClient.instance().ViewWalker.GetFirstChildElement(self.Element)
    control = Control.CreateControlFromElement(ele)

    if self.searchByWalk:
        self.searchByWalk = False
    elif control:
        control.isFirstChildFrom = self
    return control


def patched_GetNextSiblingControl(self) -> 'Control':
    """
    Return `Control` subclass or None.
    """
    ele = _AutomationClient.instance().ViewWalker.GetNextSiblingElement(self.Element)
    control = Control.CreateControlFromElement(ele)

    if self.searchByWalk:
        self.searchByWalk = False
    elif control:
        control.isNextSiblingFrom = self
    return control


def patched_Exists(self, maxSearchSeconds: float = 5, searchIntervalSeconds: float = SEARCH_INTERVAL, printIfNotExist: bool = False) -> bool:
    """
    maxSearchSeconds: float
    searchIntervalSeconds: float
    Find control every searchIntervalSeconds seconds in maxSearchSeconds seconds.
    Return bool, True if find
    """
    if self.isFirstChildFrom:
        self.isFirstChildFrom.searchByWalk = True
        control = self.isFirstChildFrom.GetFirstChildControl()
        self._element = control.Element
        control._element = 0
        return True

    if self.isNextSiblingFrom:
        self.isNextSiblingFrom.searchByWalk = True
        control = self.isNextSiblingFrom.GetNextSiblingControl()
        self._element = control.Element
        control._element = 0
        return True

    if self.isPidControl:
        try:
            return self.Pids != []
        except ApplicationError:
            return False

    if self._element and self._elementDirectAssign:
        # if element is directly assigned, not by searching, just check whether self._element is valid
        # but I can't find an API in UIAutomation that can directly check
        rootElement = GetRootControl().Element
        if self._element == rootElement:
            return True
        else:
            parentElement = _AutomationClient.instance().ViewWalker.GetParentElement(self._element)
            if parentElement:
                return True
            else:
                return False
    # find the element
    if len(self.searchProperties) == 0:
        raise LookupError("control's searchProperties must not be empty!")
    self._element = None
    startTime = ProcessTime()
    # Use same timeout(s) parameters for resolve all parents
    prev = self.searchFromControl
    if prev and (not prev.isPidControl) and not prev.Exists(maxSearchSeconds, searchIntervalSeconds):
        if printIfNotExist or DEBUG_EXIST_DISAPPEAR:
            Logger.ColorfullyLog(self.GetColorfulSearchPropertiesStr() + '<Color=Red> does not exist.</Color>')
        return False
    startTime2 = ProcessTime()
    if DEBUG_SEARCH_TIME:
        startDateTime = datetime.datetime.now()
    while True:
        control = FindControl(self.searchFromControl, self._CompareFunction, self.searchDepth, False, self.foundIndex)
        if control:
            self._element = control.Element
            control._element = 0  # control will be destroyed, but the element needs to be stroed in self._element
            if DEBUG_SEARCH_TIME:
                Logger.ColorfullyLog('{} TraverseControls: <Color=Cyan>{}</Color>, SearchTime: <Color=Cyan>{:.3f}</Color>s[{} - {}]'.format(
                    self.GetColorfulSearchPropertiesStr(), control.traverseCount, ProcessTime() - startTime2,
                    startDateTime.time(), datetime.datetime.now().time()))
            if DEBUG_MODE:
                self.Draw()
            return True
        else:
            remain = startTime + maxSearchSeconds - ProcessTime()
            if remain > 0:
                time.sleep(min(remain, searchIntervalSeconds))
            else:
                if printIfNotExist or DEBUG_EXIST_DISAPPEAR:
                    Logger.ColorfullyLog(self.GetColorfulSearchPropertiesStr() + '<Color=Red> does not exist.</Color>')
                return False


def patched_GetPattern(self, patternId: int):
    """
    Call IUIAutomationElement::GetCurrentPattern.
    Get a new pattern by pattern id if it supports the pattern.
    patternId: int, a value in class `PatternId`.
    Refer https://docs.microsoft.com/en-us/windows/desktop/api/uiautomationclient/nf-uiautomationclient-iuiautomationelement-getcurrentpattern
    """

    pattern = self.Element.GetCurrentPattern(patternId)
    if pattern:
        subPattern = CreatePattern(patternId, pattern)
        self._supportedPatterns[patternId] = subPattern
        return subPattern
    else:
        raise PatternError(f"{self.GetSearchPropertiesStr()} does not support {PatternIdNames[patternId]}")


def patched_LogControl(control: Control, depth: int = 0, showAllName: bool = True, showPid: bool = False) -> None:
    """
    Print and log control's properties.
    control: `Control` or its subclass.
    depth: int, current depth.
    showAllName: bool, if False, print the first 30 characters of control.Name.
    """
    def getKeyName(theDict, theValue):
        for key in theDict:
            if theValue == theDict[key]:
                return key
    indent = ' ' * depth * 4
    Logger.Write('{0}ControlType: '.format(indent))
    Logger.Write(control.ControlTypeName, ConsoleColor.DarkGreen)
    Logger.Write('    ClassName: ')
    Logger.Write(control.ClassName, ConsoleColor.DarkGreen)
    Logger.Write('    AutomationId: ')
    Logger.Write(control.AutomationId, ConsoleColor.DarkGreen)
    Logger.Write('    Rect: ')
    Logger.Write(control.BoundingRectangle, ConsoleColor.DarkGreen)
    Logger.Write('    Name: ')
    Logger.Write(control.Name, ConsoleColor.DarkGreen,
                 printTruncateLen=0 if showAllName else 30)
    Logger.Write('    Handle: ')
    Logger.Write('0x{0:X}({0})'.format(
        control.NativeWindowHandle), ConsoleColor.DarkGreen)
    Logger.Write('    Depth: ')
    Logger.Write(depth, ConsoleColor.DarkGreen)
    if showPid:
        Logger.Write('    ProcessId: ')
        Logger.Write(control.ProcessId, ConsoleColor.DarkGreen)

    supportedPatterns = []
    for id_, name in PatternIdNames.items():
        try:
            supportedPatterns.append((control.GetPattern(id_), name))
        except PatternError:
            pass

    for pt, name in supportedPatterns:
        if isinstance(pt, ValuePattern):
            Logger.Write('    ValuePattern.Value: ')
            Logger.Write(pt.Value, ConsoleColor.DarkGreen,
                         printTruncateLen=0 if showAllName else 30)
        elif isinstance(pt, RangeValuePattern):
            Logger.Write('    RangeValuePattern.Value: ')
            Logger.Write(pt.Value, ConsoleColor.DarkGreen)
        elif isinstance(pt, TogglePattern):
            Logger.Write('    TogglePattern.ToggleState: ')
            Logger.Write('ToggleState.' + getKeyName(ToggleState.__dict__,
                                                     pt.ToggleState), ConsoleColor.DarkGreen)
        elif isinstance(pt, SelectionItemPattern):
            Logger.Write('    SelectionItemPattern.IsSelected: ')
            Logger.Write(pt.IsSelected, ConsoleColor.DarkGreen)
        elif isinstance(pt, ExpandCollapsePattern):
            Logger.Write('    ExpandCollapsePattern.ExpandCollapseState: ')
            Logger.Write('ExpandCollapseState.' + getKeyName(
                ExpandCollapseState.__dict__, pt.ExpandCollapseState), ConsoleColor.DarkGreen)
        elif isinstance(pt, ScrollPattern):
            Logger.Write('    ScrollPattern.HorizontalScrollPercent: ')
            Logger.Write(pt.HorizontalScrollPercent,
                         ConsoleColor.DarkGreen)
            Logger.Write('    ScrollPattern.VerticalScrollPercent: ')
            Logger.Write(pt.VerticalScrollPercent, ConsoleColor.DarkGreen)
        elif isinstance(pt, GridPattern):
            Logger.Write('    GridPattern.RowCount: ')
            Logger.Write(pt.RowCount, ConsoleColor.DarkGreen)
            Logger.Write('    GridPattern.ColumnCount: ')
            Logger.Write(pt.ColumnCount, ConsoleColor.DarkGreen)
        elif isinstance(pt, GridItemPattern):
            Logger.Write('    GridItemPattern.Row: ')
            Logger.Write(pt.Column, ConsoleColor.DarkGreen)
            Logger.Write('    GridItemPattern.Column: ')
            Logger.Write(pt.Column, ConsoleColor.DarkGreen)
        elif isinstance(pt, TextPattern):
            # issue 49: CEF Control as DocumentControl have no "TextPattern.Text" property, skip log this part.
            # https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nf-uiautomationclient-iuiautomationtextpattern-get_documentrange
            try:
                Logger.Write('    TextPattern.Text: ')
                Logger.Write(pt.DocumentRange.GetText(
                    30), ConsoleColor.DarkGreen)
            except comtypes.COMError:
                pass
    Logger.Write('\nSupportedPattern:\n')
    for pt, name in supportedPatterns:
        Logger.Write(' ' + f"Get{name}" + '\n', ConsoleColor.DarkGreen)
        method_list = [func for func in dir(pt) if callable(getattr(pt, func)) and not func.startswith("__")]
        for method in method_list:
            Logger.Write('     ' + method + '\n', ConsoleColor.DarkGreen)
    Logger.Write('\n')


def patched_GetSearchPropertiesStr(self) -> str:
    strs = []
    for k, v in self.searchProperties.items():
        if k == 'ControlType':
            strs.append('{}: {}'.format(k, ControlTypeNames[v]))
        elif k == 'Child':
            strs.append('{}: {}'.format(k, v.GetSearchPropertiesStr()))
        else:
            strs.append(repr(v))

    return '{' + ', '.join(strs) + '}'


def patch_Draw(self):
    rect = self.Element.CurrentBoundingRectangle
    draw_rectangle(rect.left, rect.top, rect.right, rect.bottom)


def patched_GetChildrenByName(self, Name) -> 'Control':
    child_list = self.GetChildren()
    child = [x for x in child_list if Name == x.Name]
    if len(child) == 0:
        raise UiLookupError(f"Can not find a child control named {Name}")
    if len(child) > 1:
        raise UiLookupError(f"Find multiple children control named {Name}")
    return child[0]


def patched_SelectByPath(self, path):
    """
    path = "root_name -> layer1_name -> layer2_name"
    """
    if "->" in path:
        path_list = path.split("->")
    else:
        path_list = path.split(">")
    path_list = [x.strip() for x in path_list]
    temp = self
    for item in path_list:
        logging.debug(f"To find Child '{item}'")
        temp = temp.GetChildrenByName(Name=item)
        if not temp.Exists():
            raise UiLookupError(f"Can not find tree item '{item}'") from None

        logging.debug(f"Select tree item '{temp.Name}'")
        if item is path_list[-1]:
            temp.GetSelectionItemPattern().Select()
            return temp
        logging.debug(f"To expand '{temp.Name}'")
        temp.GetExpandCollapsePattern().Expand()


def patched_Walk(self, Name="") -> list:
    def _control_walk(control, layer=1, Name=""):
        childs = control.GetChildren()
        childs_list = []
        for child in childs:
            child_dict = {}
            child_dict["Name"] = child.Name
            child_dict["Type"] = child.ControlTypeName
            if (Name and Name == child.Name) or not Name:
                print(
                    " " * layer * 8 + f'{child_dict["Name"]} ({child_dict["Type"]}, depth={layer})')
            child_dict["Child"] = _control_walk(child, layer + 1, Name=Name)
            childs_list.append(child_dict)
        return childs_list
    return _control_walk(self, Name=Name)


def patched_Walk(self, Name="") -> list:
    def _control_walk(control, layer=1, Name=""):
        childs = control.GetChildren()
        childs_list = []
        for child in childs:
            child_dict = {}
            child_dict["Name"] = child.Name
            child_dict["Type"] = child.ControlTypeName
            if (Name and Name == child.Name) or not Name:
                print(
                    " " * layer * 8 + f'{child_dict["Name"]} ({child_dict["Type"]}, depth={layer})')
            child_dict["Child"] = _control_walk(child, layer + 1, Name=Name)
            childs_list.append(child_dict)
        return childs_list
    return _control_walk(self, Name=Name)


def patched_GetValue(self):
    result = []
    children = self.GetChildren()
    for child in children:
        result.append([x.Name for x in child.GetChildren()])

    titles = result.pop(0)
    result = [x for x in result if len(x) == len(titles)]
    pair_mode = False
    for i in range(len(result) - 1):
        if result[i][0] == result[i + 1][0]:
            pair_mode = True
            break

    if len(titles) > 2:
        all = []
        for item in result:
            temp_dict = collections.OrderedDict()
            for i in range(len(titles)):
                temp_dict[titles[i]] = item[i]
            all.append(temp_dict)
        result = all

    else:
        all = []
        temp_dict = collections.OrderedDict()
        for item in result:

            if item != ['', '']:
                if item[0] in temp_dict:
                    if isinstance(temp_dict[item[0]], str):
                        temp_dict[item[0]] = [temp_dict[item[0]], item[1]]
                    else:
                        temp_dict[item[0]].append(item[1])
                else:
                    temp_dict[item[0]] = item[1]

            if item == ['', ''] or item is result[-1]:
                all.append(temp_dict)
                temp_dict = collections.OrderedDict()
        result = all

    return result


def patched_ToggleTo(self, status, waitTime: float = OPERATION_WAIT_TIME) -> bool:
    """
    Toggle to the status
    status = on, check, checked, yes, y, true, 1
                off,  uncheck , unchecked, no, n, false, 0
                indeterminate, 2
    """
    for _ in range(10):
        if toggle_state(self.pattern.CurrentToggleState) == toggle_state(status):
            return True
        self.Toggle()
        time.sleep(waitTime)
    else:
        raise UiLookupError(f"Fail to toggle checkbox to {status}")


def patched_StateShouldBe(self, expected_state):
    actual_state = toggle_state(self.pattern.CurrentToggleState)
    expected_state = toggle_state(expected_state)
    if actual_state != expected_state:
        raise UiLookupError(f"The expected checkbox state is '{expected_state}', however the actual state is '{actual_state}'")


def patched_Write(log: Any, consoleColor: int = ConsoleColor.Default, writeToFile: bool = True, printToStdout: bool = True, logFile: str = None, printTruncateLen: int = 0) -> None:
    """
    log: any type.
    consoleColor: int, a value in class `ConsoleColor`, such as `ConsoleColor.DarkGreen`.
    writeToFile: bool.
    printToStdout: bool.
    logFile: str, log file path.
    printTruncateLen: int, if <= 0, log is not truncated when print.
    """
    if not isinstance(log, str):
        log = str(log)
    if printToStdout and sys.stdout:
        isValidColor = (consoleColor >= ConsoleColor.Black and consoleColor <= ConsoleColor.White)
        if isValidColor:
            SetConsoleColor(consoleColor)
        try:
            if printTruncateLen > 0 and len(log) > printTruncateLen:
                sys.stdout.write(log[:printTruncateLen] + '...')
            else:
                sys.stdout.write(log)
        except Exception as ex:
            SetConsoleColor(ConsoleColor.Red)
            isValidColor = True
            sys.stdout.write(ex.__class__.__name__ + ': can\'t print the log!')
            if log.endswith('\n'):
                sys.stdout.write('\n')
        if isValidColor:
            ResetConsoleColor()
        sys.stdout.flush()
    if not writeToFile:
        return
    fileName = logFile if logFile else Logger.FileName
    folder = Path(fileName).parent
    if not folder.is_dir():
        folder.mkdir(parents=True)
    fout = None
    try:
        fout = open(fileName, 'a+', encoding='utf-8')
        fout.write(log)
    except Exception as ex:
        if sys.stdout:
            sys.stdout.write(ex.__class__.__name__ + ': can\'t write the log!')
    finally:
        if fout:
            fout.close()


uiautomation.LogControl = controls_support(patched_LogControl)
uiautomation.WalkControl = patched_WalkControl

uiautomation.Control.Element = property(patched_Element)
uiautomation.Control.isPidControl = False
uiautomation.Control.searchByWalk = None
uiautomation.Control.isFirstChildFrom = None
uiautomation.Control.isNextSiblingFrom = None
uiautomation.Control.Refind = patched_Refind
uiautomation.Control.GetFirstChildControl = patched_GetFirstChildControl
uiautomation.Control.GetNextSiblingControl = patched_GetNextSiblingControl
uiautomation.Control.GetPattern = patched_GetPattern
uiautomation.Control.Log = controls_support(patched_LogControl)
uiautomation.Control.Walk = controls_support(patched_Walk)
uiautomation.Control.Exists = patched_Exists
uiautomation.Control._CompareFunction = patched_CompareFunction
uiautomation.Control.GetSearchPropertiesStr = patched_GetSearchPropertiesStr
uiautomation.Control.Draw = patch_Draw
uiautomation.Control.Screenshot = Screenshot
uiautomation.Control.GetChildrenByName = patched_GetChildrenByName

uiautomation.TreeControl.SelectByPath = patched_SelectByPath
uiautomation.DataItemControl.GetInvokePattern = uiautomation.ButtonControl.GetInvokePattern
uiautomation.ListControl.GetValue = patched_GetValue
uiautomation.CheckBoxControl.GetSelectionItemPattern = uiautomation.TreeItemControl.GetSelectionItemPattern
uiautomation.CheckBoxControl.GetExpandCollapsePattern = uiautomation.TreeItemControl.GetExpandCollapsePattern
uiautomation.TogglePattern.ToggleTo = patched_ToggleTo
uiautomation.TogglePattern.StateShouldBe = patched_StateShouldBe
uiautomation.Logger.Write = staticmethod(patched_Write)


def window_should_exists(**kwarg):
    """
    Arg1:
        Name or SubName or RegexName

    Arg2:
        timeout, default is TIME_OUT_SECOND
    """
    name = kwarg.get("Name", None) or kwarg.get("SubName", None) or kwarg.get("RegexName", None)
    window = WindowControl(**kwarg)
    if not window.Exists(kwarg.get("timeout", TIME_OUT_SECOND)):
        raise UiLookupError(f"'{name}' window does not exist",
                            filename="WindowNotFound.jpg")


def window_should_not_exists(**kwarg):
    """
    Arg1:
        Name or SubName or RegexName

    Arg2:
        timeout, default is TIME_OUT_SECOND
    """
    name = kwarg.get("Name", None) or kwarg.get("SubName", None) or kwarg.get("RegexName", None)
    window = WindowControl(**kwarg)
    if window.Exists(kwarg.get("timeout", TIME_OUT_SECOND)):
        raise UiLookupError(f"Unexpected '{name}' found, the window should not exist",
                            filename="UnexpetedWindow.jpg")
