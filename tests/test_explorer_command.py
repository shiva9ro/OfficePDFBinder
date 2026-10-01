"""Check the compiled COM DLL's lifetime contract without changing registration."""
import ctypes
from pathlib import Path
import sys
import uuid

import pytest


@pytest.mark.skipif(sys.platform != "win32", reason="Windows COM DLL")
def test_factory_keeps_dll_loaded_and_command_metadata(tmp_path):
    path = Path(__file__).parents[1] / "native/shell_bridge/sparse.build/x64/OfficePDFBinder_Explorer.dll"
    if not path.exists():
        pytest.skip("Build sparse package DLLs first")
    dll = ctypes.WinDLL(str(path))
    pointer = ctypes.c_void_p
    guid = lambda value: (ctypes.c_ubyte * 16).from_buffer_copy(uuid.UUID(value).bytes_le)
    clsid = guid("f9869918-6f6d-4fc9-a8df-70f43498a804")
    iid_factory = guid("00000001-0000-0000-c000-000000000046")
    iid_command = guid("a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9")

    def method(obj, index, result, *arguments):
        table = ctypes.cast(obj, ctypes.POINTER(ctypes.POINTER(pointer))).contents
        return ctypes.WINFUNCTYPE(result, pointer, *arguments)(table[index])

    factory = pointer()
    command = pointer()
    dll.DllGetClassObject.argtypes = [pointer, pointer, ctypes.POINTER(pointer)]
    assert dll.DllCanUnloadNow() == 0
    assert dll.DllGetClassObject(clsid, iid_factory, ctypes.byref(factory)) == 0
    try:
        # The reference project's factory is omitted from this count.
        assert dll.DllCanUnloadNow() == 1
        create = method(factory, 3, ctypes.c_long, pointer, pointer, ctypes.POINTER(pointer))
        assert create(factory, None, iid_command, ctypes.byref(command)) == 0
        method(factory, 2, ctypes.c_ulong)(factory)
        factory = pointer()
        assert dll.DllCanUnloadNow() == 1
        state = ctypes.c_uint()
        get_state = method(command, 7, ctypes.c_long, pointer, ctypes.c_int, ctypes.POINTER(ctypes.c_uint))
        assert get_state(command, None, 0, ctypes.byref(state)) == 0
        assert state.value == 2  # ECS_HIDDEN without a selection
        shell = ctypes.WinDLL("shell32")
        shell.SHParseDisplayName.argtypes = [ctypes.c_wchar_p, pointer, ctypes.POINTER(pointer),
                                            ctypes.c_uint, pointer]
        shell.SHCreateShellItemArrayFromIDLists.argtypes = [ctypes.c_uint, pointer, ctypes.POINTER(pointer)]
        ctypes.windll.ole32.CoTaskMemFree.argtypes = [pointer]
        for names, expected in [(["a.pdf"], 0), (["a.PDF", "b.png"], 0),
                                (["a.pdf", "b.txt"], 2), (["b.txt"], 2)]:
            pidls = []
            selection = pointer()
            try:
                for name in names:
                    file = tmp_path / name
                    file.touch()
                    pidl = pointer()
                    assert shell.SHParseDisplayName(str(file), None, ctypes.byref(pidl), 0, None) == 0
                    pidls.append(pidl)
                values = (pointer * len(pidls))(*pidls)
                assert shell.SHCreateShellItemArrayFromIDLists(len(pidls), values, ctypes.byref(selection)) == 0
                assert get_state(command, selection, 0, ctypes.byref(state)) == 0
                assert state.value == expected
            finally:
                if selection:
                    method(selection, 2, ctypes.c_ulong)(selection)
                for pidl in pidls:
                    ctypes.windll.ole32.CoTaskMemFree(pidl)
        title = pointer()
        get_title = method(command, 3, ctypes.c_long, pointer, ctypes.POINTER(pointer))
        assert get_title(command, None, ctypes.byref(title)) == 0
        try:
            assert "Office PDF Binder" in ctypes.wstring_at(title)
        finally:
            ctypes.windll.ole32.CoTaskMemFree.argtypes = [pointer]
            ctypes.windll.ole32.CoTaskMemFree(title)
    finally:
        if command:
            method(command, 2, ctypes.c_ulong)(command)
        if factory:
            method(factory, 2, ctypes.c_ulong)(factory)
    assert dll.DllCanUnloadNow() == 0
