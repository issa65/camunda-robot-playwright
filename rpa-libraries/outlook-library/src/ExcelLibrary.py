import subprocess
import time
import win32gui
import win32process
from robot.api import logger

from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from OfficeCommon import (
    get_execution_environment as _common_get_execution_environment,
    is_process_running as _is_process_running,
    close_process as _close_process,
    get_window_process_id as _get_window_process_id,
    wait_for_window as _wait_for_window,
    generate_test_id as _common_generate_test_id,
    control_metadata as _control_metadata,
)


ROBOT_LIBRARY_SCOPE = "GLOBAL"

EXCEL_PROCESS = "EXCEL.EXE"
EXCEL_MAIN_CLASS = "XLMAIN"

EXCEL_LOADING_CLASS = "MsoSplash"


EXCEL_START_HOME_ID = "PlaceTabHomeContent"
EXCEL_GRID_ID = "Grid"

_excel_main_handle = None
_excel_process_id = None

def generate_test_id():
    return _common_generate_test_id()

def _is_excel_running():
    return _is_process_running(EXCEL_PROCESS)


def get_execution_environment():
    return _common_get_execution_environment(
        route_host="192.168.178.30",
        route_port=26500
    )    

def inspect_excel_startup_windows(
    timeout=15,
    poll_interval=0.005
):
    """
    Fast Win32 startup inspector.

    Captures newly appearing top-level windows during Excel startup.
    Logging is done only after the observation loop so that short-lived
    splash windows are not missed because of console/UIA overhead.
    """

    if _is_excel_running():
        raise RuntimeError(
            "Excel is already running. "
            "Close Excel before inspecting startup windows."
        )

    existing_handles = set()

    def collect_existing(hwnd, _):
        existing_handles.add(hwnd)

    win32gui.EnumWindows(
        collect_existing,
        None
    )

    observed = {}
    start = time.perf_counter()

    subprocess.Popen(
        'start "" excel',
        shell=True
    )

    deadline = start + float(timeout)

    while time.perf_counter() < deadline:

        now = time.perf_counter()

        def inspect_window(hwnd, _):
            try:
                if hwnd in existing_handles:
                    return

                if hwnd in observed:
                    return

                if not win32gui.IsWindow(hwnd):
                    return

                if not win32gui.IsWindowVisible(hwnd):
                    return

                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)

                _, process_id = (
                    win32process.GetWindowThreadProcessId(hwnd)
                )

                observed[hwnd] = {
                    "firstSeenSeconds": round(
                        now - start,
                        3
                    ),
                    "handle": hwnd,
                    "name": title,
                    "className": class_name,
                    "processId": process_id,
                }

            except Exception:
                pass

        win32gui.EnumWindows(
            inspect_window,
            None
        )

        time.sleep(float(poll_interval))

    logger.console(
        "\n=== EXCEL FAST STARTUP WINDOW INSPECTOR ==="
    )

    ordered = sorted(
        observed.values(),
        key=lambda item: item["firstSeenSeconds"]
    )

    for metadata in ordered:
        logger.console(
            "WINDOW | "
            f"first_seen={metadata['firstSeenSeconds']}s | "
            f"name={metadata['name']!r} | "
            f"handle={metadata['handle']!r} | "
            f"class={metadata['className']!r} | "
            f"pid={metadata['processId']!r}"
        )

    logger.console(
        "=== END EXCEL FAST STARTUP WINDOW INSPECTOR ===\n"
    )

    return ordered


def _ensure_excel_workbook(window, timeout=30):
    """
    Ensures that Excel has an actual workbook open.

    Excel uses XLMAIN both for:
    - the startup/home screen
    - an opened workbook

    A workbook is considered ready when the Excel DataGrid exists.
    """

    global _excel_main_handle
    global _excel_process_id

    # Fast path: workbook is already open.
    grid = window.child_window(
        auto_id=EXCEL_GRID_ID,
        control_type="DataGrid"
    )

    if grid.exists(timeout=0.2):
        return window

    # Check whether Excel is currently showing its start page.
    start_page = window.child_window(
        auto_id=EXCEL_START_HOME_ID,
        control_type="Group"
    )

    if not start_page.exists(timeout=0.5):
        raise RuntimeError(
            "Excel is open, but neither a workbook "
            "nor the Excel start page was detected."
        )

    # Stable compound locator:
    # Excel start page -> blank workbook template.
    blank_workbook = start_page.child_window(
        title="Leere Arbeitsmappe",
        control_type="ListItem",
        class_name="NetUIListViewItem"
    )

    blank_workbook.wait(
        "visible",
        timeout=float(timeout)
    )

    blank_workbook.click_input()

    deadline = time.perf_counter() + float(timeout)
    desktop = Desktop(backend="uia")

    while time.perf_counter() < deadline:

        # Excel normally reuses the same XLMAIN window,
        # but reacquire it through the known handle.
        if _excel_main_handle is not None:
            try:
                window = desktop.window(
                    handle=_excel_main_handle
                )
            except Exception:
                pass

        grid = window.child_window(
            auto_id=EXCEL_GRID_ID,
            control_type="DataGrid"
        )

        if grid.exists(timeout=0.1):
            wrapper = window.wrapper_object()

            _excel_main_handle = wrapper.handle
            _excel_process_id = _get_window_process_id(
                wrapper
            )

            return window

        time.sleep(0.05)

    raise RuntimeError(
        "Blank Excel workbook was opened, "
        "but the worksheet grid was not detected."
    )


def _get_excel_main_window(timeout=30):
    global _excel_main_handle
    global _excel_process_id

    desktop = Desktop(backend="uia")

    # Bereits eindeutig erkanntes Excel-Fenster wiederverwenden.
    if _excel_main_handle is not None:
        try:
            window = desktop.window(
                handle=_excel_main_handle
            )

            if (
                window.exists(timeout=0.2)
                and window.is_visible()
            ):
                return window

        except Exception:
            _excel_main_handle = None

    window = _wait_for_window(
        class_name=EXCEL_MAIN_CLASS,
        control_type="Window",
        process_id=_excel_process_id,
        timeout=timeout,
        poll_interval=0.05
    )

    wrapper = window.wrapper_object()

    _excel_main_handle = wrapper.handle
    _excel_process_id = _get_window_process_id(
        wrapper
    )

    return window


def measure_excel_startup(
    timeout=30,
    poll_interval=0.005
):
    """
    Measures Excel cold startup.

    Measures:
    - process start -> MsoSplash
    - MsoSplash -> XLMAIN
    - process start -> XLMAIN

    Supports:
    - LOADING_TO_MAIN
    - DIRECT_MAIN
    """

    global _excel_main_handle
    global _excel_process_id

    if _is_excel_running():
        raise RuntimeError(
            "Excel is already running. "
            "Cold startup measurement aborted."
        )

    _excel_main_handle = None
    _excel_process_id = None

    existing_handles = set()

    def collect_existing(hwnd, _):
        existing_handles.add(hwnd)

    win32gui.EnumWindows(
        collect_existing,
        None
    )

    start = time.perf_counter()

    loading_seen_at = None
    loading_process_id = None

    main_seen_at = None
    main_handle = None
    main_process_id = None

    observed_windows = []
    observed_handles = set()

    subprocess.Popen(
        'start "" excel',
        shell=True
    )

    deadline = start + float(timeout)

    while time.perf_counter() < deadline:

        now = time.perf_counter()

        detected = []

        def inspect_window(hwnd, _):
            try:
                if hwnd in existing_handles:
                    return

                if not win32gui.IsWindow(hwnd):
                    return

                if not win32gui.IsWindowVisible(hwnd):
                    return

                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)

                _, process_id = (
                    win32process.GetWindowThreadProcessId(hwnd)
                )

                detected.append(
                    {
                        "handle": hwnd,
                        "name": title,
                        "className": class_name,
                        "processId": process_id,
                    }
                )

            except Exception:
                pass

        win32gui.EnumWindows(
            inspect_window,
            None
        )

        for metadata in detected:

            handle = metadata["handle"]
            class_name = metadata["className"]
            process_id = metadata["processId"]

            if handle not in observed_handles:
                observed_handles.add(handle)

                observed_windows.append(
                    {
                        **metadata,
                        "firstSeenSeconds": round(
                            now - start,
                            3
                        )
                    }
                )

            # ------------------------------------------------
            # Loading / Splash
            # ------------------------------------------------
            if (
                loading_seen_at is None
                and class_name == EXCEL_LOADING_CLASS
            ):
                loading_seen_at = now
                loading_process_id = process_id

            # ------------------------------------------------
            # Excel Main
            # ------------------------------------------------
            if class_name == EXCEL_MAIN_CLASS:

                # Falls Splash erkannt wurde:
                # Main muss zum selben Excel-Prozess gehören.
                if (
                    loading_process_id is not None
                    and process_id != loading_process_id
                ):
                    continue

                main_seen_at = now
                main_handle = handle
                main_process_id = process_id
                break

        if main_seen_at is not None:
            break

        time.sleep(float(poll_interval))

    if main_seen_at is None:
        raise RuntimeError(
            f"Excel main window was not detected "
            f"within {timeout} seconds."
        )

    startup_to_main = round(
        main_seen_at - start,
        3
    )

    if loading_seen_at is not None:

        startup_to_loading = round(
            loading_seen_at - start,
            3
        )

        loading_to_main = round(
            main_seen_at - loading_seen_at,
            3
        )

        startup_path = "LOADING_TO_MAIN"

    else:
        startup_to_loading = None
        loading_to_main = None
        startup_path = "DIRECT_MAIN"

    # Main Window noch einmal über UIA validieren.
    desktop = Desktop(
        backend="uia"
    )

    main_window = desktop.window(
        handle=int(main_handle)
    )

    main_window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    _excel_main_handle = int(main_handle)
    _excel_process_id = main_process_id

    return {
        "startupToLoadingSeconds": startup_to_loading,
        "loadingToMainSeconds": loading_to_main,
        "startupToMainSeconds": startup_to_main,
        "startupPath": startup_path,

        "windowTitle": main_window.window_text(),
        "windowHandle": _excel_main_handle,
        "processId": _excel_process_id,

        "startupType": "COLD",
        "observedWindows": observed_windows
    }

def open_excel(timeout=30):
    global _excel_main_handle
    global _excel_process_id

    # Cold/new Excel start.
    if not _is_excel_running():

        # Alte Handles/PIDs niemals über einen Prozess-Neustart behalten.
        _excel_main_handle = None
        _excel_process_id = None

        subprocess.Popen(
            'start "" excel',
            shell=True
        )

    window = _get_excel_main_window(timeout)

    window.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    # Wichtig:
    # XLMAIN kann entweder Startseite oder echte Arbeitsmappe sein.
    window = _ensure_excel_workbook(
        window,
        timeout=timeout
    )

    wrapper = window.wrapper_object()

    _excel_main_handle = wrapper.handle
    _excel_process_id = _get_window_process_id(
        wrapper
    )

    return window.window_text()


def select_excel_cell(cell="A1", timeout=10):
    """
    Selects one Excel cell by its unique UI Automation ID.

    Example:
        A1
        B2
        C10
    """

    window = _get_excel_main_window(timeout)

    cell = str(cell).upper().strip()

    cell_control = window.child_window(
        auto_id=cell,
        control_type="DataItem"
    )

    cell_control.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    cell_control.click_input()

    return cell


def write_excel_cell(cell, value, timeout=10):
    """
    Writes a value into a specific Excel cell.
    """

    select_excel_cell(
        cell,
        timeout=timeout
    )

    # F2 activates editing of the selected cell.
    send_keys("{F2}")

    # Existing cell content ersetzen.
    send_keys("^a")

    send_keys(
        str(value),
        with_spaces=True
    )

    send_keys("{ENTER}")

    return {
        "cell": str(cell).upper().strip(),
        "value": str(value)
    }


def read_excel_cell(cell, timeout=10):
    """
    Reads the currently stored value through Excel's Formula Bar.

    The selected cell itself uses its address as automation ID,
    while FormulaBar is a unique Excel control.
    """

    window = _get_excel_main_window(timeout)

    cell = str(cell).upper().strip()

    select_excel_cell(
        cell,
        timeout=timeout
    )

    formula_bar = window.child_window(
        auto_id="FormulaBar",
        control_type="Edit"
    )

    formula_bar.wait(
        "visible",
        timeout=float(timeout)
    )

    wrapper = formula_bar.wrapper_object()

    value = ""

    # UIA ValuePattern bevorzugen.
    try:
        value = wrapper.iface_value.CurrentValue
    except Exception:
        try:
            value = wrapper.window_text()
        except Exception:
            value = ""

    raw_value = str(value)

    # Excel FormulaBar kann bei einfachen Zellwerten
    # einen führenden "=" zurückgeben.
    normalized_value = raw_value

    if normalized_value.startswith("="):
        normalized_value = normalized_value[1:]

    return {
        "cell": cell,
        "value": normalized_value,
        "rawValue": raw_value
    }


def close_excel(timeout=10):
    global _excel_main_handle
    global _excel_process_id

    result = _close_process(
        EXCEL_PROCESS,
        timeout=timeout
    )

    _excel_main_handle = None
    _excel_process_id = None

    return result


def ensure_excel_closed(timeout=10):
    if not _is_excel_running():
        return True

    close_excel(timeout)

    if _is_excel_running():
        raise RuntimeError(
            "Excel is still running. "
            "Test environment could not be prepared."
        )

    return True