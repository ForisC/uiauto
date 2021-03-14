from datetime import datetime
from uiautomation import *
from uiautomation.uiautomation import _AutomationClient


def autoRefind(fn):
    def wrapper(self, *args, **kwargs):
        self.Refind()
        return fn(self, *args, **kwargs)
    return wrapper


def autoRefindProperty(fn):
    fn = fn.fget

    def wrapper(self, *args, **kwargs):
        self.Refind()
        return fn(self, *args, **kwargs)
    return property(wrapper)


def WalkControl(control: Control, includeTop: bool = False, maxDepth: int = 0xFFFFFFFF, autoRefind=True):
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
    depth = 0

    control.searchByWalk = True
    child = control.GetFirstChildControl(autoRefind=autoRefind)
    controlList = [child]
    while depth >= 0:
        lastControl = controlList[-1]
        if lastControl:
            yield lastControl, depth + 1
            child = lastControl.GetNextSiblingControl(autoRefind=autoRefind)
            controlList[depth] = child
            if depth + 1 < maxDepth:
                lastControl.searchByWalk = True
                child = lastControl.GetFirstChildControl(autoRefind=autoRefind)
                if child:
                    depth += 1
                    controlList.append(child)
        else:
            del controlList[depth]
            depth -= 1


def FindControl(control: Control, compare: Callable[[Control, int], bool], maxDepth: int = 0xFFFFFFFF, findFromSelf: bool = False, foundIndex: int = 1, autoRefind=True) -> Control:
    """
    control: `Control` or its subclass.
    compare: Callable[[Control, int], bool], function(control: Control, depth: int) -> bool.
    maxDepth: int, enum depth.
    findFromSelf: bool, if False, do not compare self.
    foundIndex: int, starts with 1, >= 1.
    Return `Control` subclass or None if not find.
    """
    foundCount = 0

    if not control:
        control = GetRootControl()
    elif autoRefind:
        control.Refind()
    traverseCount = 0

    for child, depth in WalkControl(control, findFromSelf, maxDepth, autoRefind=autoRefind):
        traverseCount += 1
        if compare(child, depth):
            foundCount += 1
            if foundCount == foundIndex:
                child.traverseCount = traverseCount
                return child


def LogControl(control: Control, depth: int = 0, showAllName: bool = True, showPid: bool = False) -> None:
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
            supportedPatterns.append((
                control.GetPattern(id_, autoRefind=False), name))
        except RuntimeError:
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
            Logger.Write(pt.HorizontalScrollPercent, ConsoleColor.DarkGreen)
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
            except comtypes.COMError as ex:
                pass
    Logger.Write('    SupportedPattern:')
    for pt, name in supportedPatterns:
        Logger.Write(' ' + name, ConsoleColor.DarkGreen)
    Logger.Write('\n')


def GetRootControl() -> PaneControl:
    """
    Get root control, the Desktop window.
    Return `PaneControl`.
    """
    control = Control.CreateControlFromElement(
        _AutomationClient.instance().IUIAutomation.GetRootElement())
    control.isRootControl = True
    return control


def EnumAndLogControl(control: Control, maxDepth: int = 0xFFFFFFFF, showAllName: bool = True, showPid: bool = False, startDepth: int = 0) -> None:
    """
    Print and log control and its descendants' propertyies.
    control: `Control` or its subclass.
    maxDepth: int, enum depth.
    showAllName: bool, if False, print the first 30 characters of control.Name.
    startDepth: int, control's current depth.
    """
    for c, d in WalkControl(control, True, maxDepth):
        LogControl(c, d + startDepth, showAllName, showPid)


class patched_function(object):
    def GetSearchPropertiesStr(self) -> str:
        strs = []
        for k, v in self.searchProperties.items():
            if k == 'ControlType':
                strs.append('{}: {}'.format(k, ControlTypeNames[v]))
            elif k == 'Child':
                strs.append('{}: {}'.format(k, v.GetSearchPropertiesStr()))
            else:
                strs.append(repr(v))

        return '{' + ', '.join(strs) + '}'

    def Refind(self, maxSearchSeconds: float = TIME_OUT_SECOND, searchIntervalSeconds: float = SEARCH_INTERVAL, raiseException: bool = True) -> bool:
        """
        Refind the control every searchIntervalSeconds seconds in maxSearchSeconds seconds.
        maxSearchSeconds: float.
        searchIntervalSeconds: float.
        raiseException: bool, if True, raise a LookupError if timeout.
        Return bool, True if find.
        """
        if self.isRootControl:
            # self = GetRootControl()
            return True
        if not self.Exists(maxSearchSeconds, searchIntervalSeconds, False if raiseException else DEBUG_EXIST_DISAPPEAR):
            if raiseException:
                Logger.ColorfullyLog(
                    '<Color=Red>Find Control Timeout: </Color>' + self.GetColorfulSearchPropertiesStr())
                raise LookupError('Find Control Timeout: ' +
                                  self.GetSearchPropertiesStr())
            else:
                return False
        return True

    def GetFirstChildControl(self, autoRefind=True) -> 'Control':
        """
        Return `Control` subclass or None.
        """
        if autoRefind:
            _ele = self.Element
        else:
            _ele = self._element
        ele = _AutomationClient.instance().ViewWalker.GetFirstChildElement(_ele)
        control = Control.CreateControlFromElement(ele)

        if self.searchByWalk:
            self.searchByWalk = False
        elif control:
            control.isFirstChildFrom = self
        return control

    def GetNextSiblingControl(self, autoRefind=True) -> 'Control':
        """
        Return `Control` subclass or None.
        """
        if autoRefind:
            _ele = self.Element
        else:
            _ele = self._element
        ele = _AutomationClient.instance().ViewWalker.GetNextSiblingElement(_ele)
        control = Control.CreateControlFromElement(ele)

        if self.searchByWalk:
            self.searchByWalk = False
        elif control:
            control.isNextSiblingFrom = self
        return control

    def Exists(self, maxSearchSeconds: float = 5, searchIntervalSeconds: float = SEARCH_INTERVAL, printIfNotExist: bool = False, autoRefind=True) -> bool:
        """
        maxSearchSeconds: float
        searchIntervalSeconds: float
        Find control every searchIntervalSeconds seconds in maxSearchSeconds seconds.
        Return bool, True if find
        """
        if self.isFirstChildFrom:
            self.isFirstChildFrom.Refind()
            self.isFirstChildFrom.searchByWalk = True
            control = self.isFirstChildFrom.GetFirstChildControl()
            self._element = control.Element
            control._element = 0
            return True

        if self.isNextSiblingFrom:
            self.isNextSiblingFrom.Refind()
            self.isNextSiblingFrom.searchByWalk = True
            control = self.isNextSiblingFrom.GetNextSiblingControl()
            self._element = control.Element
            control._element = 0
            return True

        if self._element and self._elementDirectAssign:
            # if element is directly assigned, not by searching, just check whether self._element is valid
            # but I can't find an API in UIAutomation that can directly check
            rootElement = GetRootControl().Element
            if self._element == rootElement:
                return True
            else:
                parentElement = _AutomationClient.instance(
                ).ViewWalker.GetParentElement(self._element)
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

        if autoRefind and prev and not prev._element and not prev.Exists(maxSearchSeconds, searchIntervalSeconds):
            if printIfNotExist or DEBUG_EXIST_DISAPPEAR:
                Logger.ColorfullyLog(self.GetColorfulSearchPropertiesStr(
                ) + '<Color=Red> does not exist.</Color>')
            return False
        startTime2 = ProcessTime()
        if DEBUG_SEARCH_TIME:
            startDateTime = datetime.datetime.now()
        while True:
            control = FindControl(
                self.searchFromControl, self._CompareFunction, self.searchDepth, False, self.foundIndex, autoRefind=autoRefind)
            if control:
                self._element = control.Element
                # control will be destroyed, but the element needs to be stroed in self._element
                control._element = 0
                if DEBUG_SEARCH_TIME:
                    Logger.ColorfullyLog('{} TraverseControls: <Color=Cyan>{}</Color>, SearchTime: <Color=Cyan>{:.3f}</Color>s[{} - {}]'.format(
                        self.GetColorfulSearchPropertiesStr(), control.traverseCount, ProcessTime() -
                        startTime2,
                        startDateTime.time(), datetime.datetime.now().time()))
                return True
            else:
                remain = startTime + maxSearchSeconds - ProcessTime()
                if remain > 0:
                    time.sleep(min(remain, searchIntervalSeconds))
                else:
                    if printIfNotExist or DEBUG_EXIST_DISAPPEAR:
                        Logger.ColorfullyLog(self.GetColorfulSearchPropertiesStr(
                        ) + '<Color=Red> does not exist.</Color>')
                    return False

    def _CompareFunction(self, control: 'Control', depth: int) -> bool:
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
                value.searchFromControl = control
                if not value.Exists(autoRefind=False):
                    return False
            # elif 'ChildProperty' == key:
            #     value.SetSearchFromControl(self)
            #     if not value.Exists():
            #         return False
        return True

    def GetPattern(self, patternId: int, autoRefind=True):
        """
        Call IUIAutomationElement::GetCurrentPattern.
        Get a new pattern by pattern id if it supports the pattern.
        patternId: int, a value in class `PatternId`.
        Refer https://docs.microsoft.com/en-us/windows/desktop/api/uiautomationclient/nf-uiautomationclient-iuiautomationelement-getcurrentpattern
        """
        if autoRefind:
            self.Refind()

        pattern = self.Element.GetCurrentPattern(patternId)
        if pattern:
            subPattern = CreatePattern(patternId, pattern)
            self._supportedPatterns[patternId] = subPattern
            return subPattern
        else:
            raise RuntimeError(
                f"The control didn't support {PatternIdNames[patternId]}")

    def Element(self):
        """
        Property Element.
        Return `ctypes.POINTER(IUIAutomationElement)`.
        """
        _error = None
        for _ in range(6):
            try:
                if not self._element:
                    self.Refind(maxSearchSeconds=TIME_OUT_SECOND,
                                searchIntervalSeconds=self.searchInterval)
                return self._element
            except comtypes.COMError as ex:
                _error = ex
                self.Refind()
                time.sleep(20)
        else:
            raise _error


# new property for auto refinding
Control.searchByWalk = False
Control.isFirstChildFrom = None
Control.isNextSiblingFrom = None
Control.isRootControl = False

# decorate functions, auto refind control when function execution
Control.SetFocus = autoRefind(Control.SetFocus)
Control.GetPattern = autoRefind(Control.GetPattern)
Control.__str__ = autoRefind(Control.__str__)

# patch functions
Control.GetSearchPropertiesStr = patched_function.GetSearchPropertiesStr
Control.Element = property(patched_function.Element)
Control.GetPattern = patched_function.GetPattern
Control.GetFirstChildControl = patched_function.GetFirstChildControl
Control.GetNextSiblingControl = patched_function.GetNextSiblingControl
Control.Exists = patched_function.Exists
Control._CompareFunction = patched_function._CompareFunction
Control.Refind = patched_function.Refind
Control.Log = LogControl
Control.Walk = EnumAndLogControl
