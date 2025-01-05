from win32com.client import DispatchEx, Dispatch, CDispatch
from typing import Any

__version__ = '0.1'

class ExcelApp:
    _registry: set
    visible: bool
    separate_instance: bool

    def __init__(self, visible: bool = True, separate_instance = False):
        self._registry = set()
        self.visible = visible
        self.separate_instance = separate_instance

    def __enter__(self):
        if self.separate_instance:
            dispatch = DispatchEx
        else:
            dispatch = Dispatch

        self._app = COMWrapper(dispatch("Excel.Application"), self)
        self._app.Visible=bool(self.visible)  # Casting to ensure nothing weird happens
        return self._app

    def __exit__(self, exc_type, exc_value, traceback):
        while self._app.Workbooks.Count > 0:
            self._app.Workbooks(1).Close(False)

        self._app.Quit()
        del self._app
        self.cleanup()

    def register(self, obj):
        self._registry.add(obj)

    def cleanup(self, namespace: dict=None):
        if namespace is None:
            namespace = globals()

        for key, value in namespace.items():
            if isinstance(value, COMWrapper) and value in self._registry:
                del namespace[key]

        while self._registry:
            item = self._registry.pop()
            item._self = None

    def __del__(self):
        self.cleanup()

class COMWrapper:
    _self: Any = None
    _is_method: bool = False
    _app: list

    def __init__(self, obj, app: ExcelApp):
        self._self = obj
        if isinstance(obj, CDispatch):
            self._is_method = False
        elif str(type(obj)) == "<class 'method'>":
            self._is_method = True
        else:
            raise TypeError(type(obj))
        self._parent = app
        app.register(self)

    def __call__(self, *args, **kwargs):
        if not self._is_method and not hasattr(self._self, 'Count'):
            raise AttributeError("__call__")
        result = self._self(*args, **kwargs)
        if isinstance(result, CDispatch) or str(type(result)) == "<class 'method'>":
            return COMWrapper(result, self._parent)
        return result


    def __getattr__(self, key):
        if self._is_method:
            raise AttributeError(key)
        if not hasattr(self._self, key):
            raise AttributeError(key)

        result = getattr(self._self, key)
        if isinstance(result, (int, float, str, bool)):
            return result
        return COMWrapper(result, self._parent)

    def __setattr__(self, key, value):
        if key[0] == '_':
            super().__setattr__(key, value)
            return None
        if object.__getattribute__(self,'_is_method') or not hasattr(self._self, key):
            raise AttributeError(key)
        setattr(self._self, key, value)

    def __del__(self):
        del self._self

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, type(self)):
            return other._self is self._self
        return False

    def __hash__(self):
        return hash((self.__class__.__name__,id(self._self)))