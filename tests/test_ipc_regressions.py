"""Exercise actual local sockets and the GUI queue at worker boundaries."""
import uuid
import sys
import subprocess
import threading
from pathlib import Path

import pytest
from PySide6.QtCore import Qt
from PySide6.QtNetwork import QLocalServer, QLocalSocket

import OfficePDFBinder_Main as app


def frame(paths):
    body = "\n".join(paths).encode("utf-8")
    return b"OPB2" + len(body).to_bytes(4, "little") + body


def test_native_and_python_use_same_session_pipe():
    driver = Path(__file__).parents[1] / "native/shell_bridge/out.build/shell_bridge_test.exe"
    if not driver.exists():
        pytest.skip("Build native -Tests first")
    result = subprocess.run([str(driver), "--pipe-name"], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == app._instance_server_name("Installed", app._current_session_id())
    assert app._instance_server_name("Installed", 1) != app._instance_server_name("Installed", 2)
    assert app._SINGLE_INSTANCE_MUTEX_NAME.startswith("Local\\")


@pytest.fixture
def ipc_server(main_window):
    server = QLocalServer(main_window)
    name = "OfficePDFBinder_test_" + uuid.uuid4().hex
    assert server.listen(name)
    server.newConnection.connect(lambda: app._handle_new_connection(server, main_window))
    yield server, name
    server.close()


def send_part(socket, data):
    assert socket.write(data) == len(data)
    socket.flush()


def test_valid_request_restores_minimized_window(main_window, ipc_server, pdf_factory, qtbot, monkeypatch):
    calls = []
    monkeypatch.setattr(main_window, "activateWindow", lambda: calls.append("activate"))
    main_window.showMinimized()
    socket = QLocalSocket(main_window)
    socket.connectToServer(ipc_server[1])
    assert socket.waitForConnected(1000)
    send_part(socket, frame([str(pdf_factory("front.pdf", ["page"]))]))
    qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 1)
    qtbot.waitUntil(lambda: main_window.threadpool.activeThreadCount() == 0)
    assert not main_window.isMinimized()
    assert calls == ["activate"]


@pytest.mark.parametrize("split", [False, True])
def test_two_paths_over_real_socket(main_window, ipc_server, pdf_factory, qtbot, split):
    paths = [str(pdf_factory(name, ["page"])) for name in ["01_first.pdf", "02_日本語.pdf"]]
    payload = "\n".join(paths).encode("utf-8")
    socket = QLocalSocket(main_window)
    socket.connectToServer(ipc_server[1])
    assert socket.waitForConnected(1000)
    if split:
        # Split inside a multibyte character, after the first complete path.
        cut = payload.index("日".encode("utf-8")) + 1
        send_part(socket, payload[:cut])
        qtbot.wait(100)
        send_part(socket, payload[cut:])
    else:
        send_part(socket, payload)
    socket.disconnectFromServer()
    qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 2, timeout=5000)
    qtbot.waitUntil(lambda: main_window.threadpool.activeThreadCount() == 0)
    actual = [main_window.page_list_widget.item(i).data(Qt.UserRole)["original_path"] for i in range(2)]
    assert actual == paths


def test_ipc_arrives_after_completion_before_pool_release(main_window, monkeypatch, qtbot):
    active = [1]
    monkeypatch.setattr(main_window.threadpool, "activeThreadCount", lambda: active[0])
    monkeypatch.setattr(main_window, "_check_duplicate_files", lambda paths: paths)
    calls = []
    monkeypatch.setattr(main_window, "_run_task", lambda *args, **kwargs: calls.append(kwargs) if not active[0] else None)
    main_window.current_worker = None
    main_window._ipc_pending_file_paths = ["second.pdf"]
    main_window._flush_ipc_pending_files()
    assert not calls
    active[0] = 0
    qtbot.waitUntil(lambda: bool(calls), timeout=1500)
    assert calls[0]["file_paths"] == ["second.pdf"]


@pytest.mark.parametrize("count", [2, 16, 100])
def test_framed_selection_ack_and_all_files(main_window, ipc_server, pdf_factory, qtbot, count):
    paths = [str(pdf_factory(f"{i:03}_日本語.pdf", [str(i)])) for i in range(count)]
    socket = QLocalSocket(main_window)
    socket.connectToServer(ipc_server[1])
    assert socket.waitForConnected(1000)
    packet = frame(paths)
    # Deliberately fragment the protocol header and body.
    for part in [packet[:2], packet[2:7], packet[7:43], packet[43:]]:
        send_part(socket, part)
        qtbot.wait(10)
    qtbot.waitUntil(lambda: socket.bytesAvailable() >= 3)
    assert bytes(socket.readAll()) == b"OK\n"
    qtbot.waitUntil(lambda: main_window.page_list_widget.count() == count, timeout=10000)
    qtbot.waitUntil(lambda: main_window.current_worker is None and main_window.threadpool.activeThreadCount() == 0)


def test_truncated_frame_adds_nothing(main_window, ipc_server, pdf_factory, qtbot):
    source = str(pdf_factory("first.pdf", ["first"]))
    socket = QLocalSocket(main_window)
    socket.connectToServer(ipc_server[1])
    assert socket.waitForConnected(1000)
    send_part(socket, frame([source, "second.pdf"])[:-5])
    socket.disconnectFromServer()
    qtbot.wait(700)
    assert main_window.page_list_widget.count() == 0


def test_receive_timeout_disconnects_and_accepts_next_request(main_window, ipc_server, pdf_factory, qtbot):
    source = str(pdf_factory("after_timeout.pdf", ["page"]))
    stalled = QLocalSocket(main_window)
    stalled.connectToServer(ipc_server[1])
    assert stalled.waitForConnected(1000)
    send_part(stalled, frame([source])[:-1])
    server = ipc_server[0]
    qtbot.waitUntil(lambda: bool(server.findChildren(app.QTimer)))
    # Exercise the real timeout callback without waiting 30 seconds.
    timers = server.findChildren(app.QTimer)
    assert len(timers) == 1
    timers[0].start(1)
    qtbot.waitUntil(lambda: stalled.state() == QLocalSocket.UnconnectedState)
    assert main_window.page_list_widget.count() == 0

    next_socket = QLocalSocket(main_window)
    next_socket.connectToServer(ipc_server[1])
    assert next_socket.waitForConnected(1000)
    send_part(next_socket, frame([source]))
    qtbot.waitUntil(lambda: next_socket.bytesAvailable() >= 3)
    assert bytes(next_socket.readAll()) == b"OK\n"
    qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 1)
    qtbot.waitUntil(lambda: main_window.current_worker is None and main_window.threadpool.activeThreadCount() == 0)


def test_real_worker_tail_does_not_lose_next_selection(main_window, ipc_server, pdf_factory, qtbot, monkeypatch):
    first = str(pdf_factory("first.pdf", ["first"]))
    second = str(pdf_factory("second.pdf", ["second"]))
    release = threading.Event()
    tail = threading.Event()
    original = app.AppWorker.run

    def slow_tail(worker):
        original(worker)
        if worker.kwargs.get("file_paths") == [first]:
            tail.set()
            release.wait(10)

    monkeypatch.setattr(app.AppWorker, "run", slow_tail)
    main_window._run_task("add_files", "Test", file_paths=[first])
    try:
        qtbot.waitUntil(lambda: tail.is_set() and main_window.current_worker is None)
        socket = QLocalSocket(main_window)
        socket.connectToServer(ipc_server[1])
        assert socket.waitForConnected(1000)
        send_part(socket, frame([second]))
        qtbot.waitUntil(lambda: main_window._pending_file_paths == [second])
        assert main_window.page_list_widget.count() == 1
    finally:
        release.set()
    qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 2)
    qtbot.waitUntil(lambda: main_window.current_worker is None and main_window.threadpool.activeThreadCount() == 0)


def test_native_shell_item_array_to_python_receiver(main_window, ipc_server, pdf_factory, qtbot):
    driver = Path(__file__).resolve().parents[1] / "native/shell_bridge/out.build/shell_bridge_test.exe"
    if not driver.exists():
        pytest.skip("Build native tests with scripts/build_shell_bridge.ps1 -Tests")
    paths = [str(pdf_factory(name, ["page"])) for name in ["01_日本語 空白.pdf", "02_日本語.pdf"]]
    process = subprocess.Popen([str(driver), ipc_server[1], "", *paths],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        qtbot.waitUntil(lambda: process.poll() is not None, timeout=15000)
        stdout, stderr = process.communicate()
        assert process.returncode == 0, (stdout, stderr)
        qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 2)
        qtbot.waitUntil(lambda: main_window.current_worker is None and main_window.threadpool.activeThreadCount() == 0)
    finally:
        if process.poll() is None:
            process.kill()
            process.wait()


@pytest.mark.skipif(sys.platform != "win32", reason="Windows COM integration")
@pytest.mark.parametrize("prestarted", [False, True])
def test_local_com_activation_and_selection(main_window, pdf_factory, qtbot, prestarted):
    """Use a test-only CLSID/pipe; never change the installed application's verb."""
    import winreg
    import ctypes
    # Elevated COM clients ignore per-user registrations. Match Inno's HKA
    # behavior for the test-only CLSID, then remove it in either case.
    registry_root = winreg.HKEY_LOCAL_MACHINE if ctypes.windll.shell32.IsUserAnAdmin() else winreg.HKEY_CURRENT_USER
    root = Path(__file__).resolve().parents[1] / "native/shell_bridge/out.build"
    driver = root / "shell_bridge_test.exe"
    helper = root / "shell_bridge_integration.exe"
    if not driver.exists() or not helper.exists():
        pytest.skip("Build native -Tests and -IntegrationServer first")
    class_id = "{" + str(uuid.uuid4()) + "}"
    key = "Software\\Classes\\CLSID\\" + class_id
    app_key = "Software\\Classes\\AppID\\" + class_id
    try:
        with winreg.OpenKey(registry_root, key):
            pytest.fail("Test-only COM registration already exists; refusing to overwrite")
    except FileNotFoundError:
        pass
    paths = [str(pdf_factory(name, ["page"])) for name in ["01_日本語.pdf", "02_space name.pdf"]]
    server = QLocalServer(main_window)
    assert server.listen("OfficePDFBinder_ShellBridge_IntegrationTest")
    server.newConnection.connect(lambda: app._handle_new_connection(server, main_window))
    process = None
    helper_process = None
    try:
        with winreg.CreateKeyEx(registry_root, key, 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as registration:
            winreg.SetValueEx(registration, "", 0, winreg.REG_SZ, "Binder integration test")
            winreg.SetValueEx(registration, "AppID", 0, winreg.REG_SZ, class_id)
        with winreg.CreateKeyEx(registry_root, app_key, 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as registration:
            winreg.SetValueEx(registration, "", 0, winreg.REG_SZ, "Binder integration test")
        with winreg.CreateKeyEx(registry_root, key + r"\LocalServer32", 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as registration:
            winreg.SetValueEx(registration, "", 0, winreg.REG_SZ, f'"{helper}" {class_id}')
            winreg.SetValueEx(registration, "ServerExecutable", 0, winreg.REG_SZ, str(helper))
        assert helper.is_file()
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
        qtbot.wait(100)
        if prestarted:
            helper_process = subprocess.Popen([str(helper), class_id], creationflags=subprocess.CREATE_NO_WINDOW)
            qtbot.wait(500)
        process = subprocess.Popen([str(driver), "--com", class_id, *paths],
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   creationflags=subprocess.CREATE_NO_WINDOW)
        qtbot.waitUntil(lambda: process.poll() is not None, timeout=20000)
        stdout, stderr = process.communicate()
        assert process.returncode == 0, (stdout, stderr)
        qtbot.waitUntil(lambda: main_window.page_list_widget.count() == 2)
        qtbot.waitUntil(lambda: main_window.current_worker is None and main_window.threadpool.activeThreadCount() == 0)
    finally:
        if process is not None and process.poll() is None:
            process.kill()
            process.wait()
        if helper_process is not None and helper_process.poll() is None:
            helper_process.kill()
            helper_process.wait()
        server.close()
        winreg.DeleteKey(registry_root, key + r"\LocalServer32")
        winreg.DeleteKey(registry_root, key)
        winreg.DeleteKey(registry_root, app_key)
