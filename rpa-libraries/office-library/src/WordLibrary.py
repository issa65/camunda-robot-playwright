import win32gui
import win32process

from time import perf_counter
import subprocess
import time

from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from OfficeCommon import (
    is_process_running as _is_process_running,
    close_process as _close_process,
    wait_for_window as _wait_for_window,
    get_window_process_id as _get_window_process_id,
    get_execution_environment as _common_get_execution_environment,
    generate_test_id as _common_generate_test_id,
    save_office_file as _save_office_file,
    save_office_file_as as _save_office_file_as,
    create_new_office_document as _create_new_office_document,
)


ROBOT_LIBRARY_SCOPE = "GLOBAL"


WORD_PROCESS = "WINWORD.EXE"
WORD_COMMAND = "winword"

WORD_LOADING_CLASS = "MsoSplash"
WORD_MAIN_CLASS = "OpusApp"

WORD_START_HOME_ID = "PlaceTabHomeContent"

WORD_BLANK_DOCUMENT_NAME = "Leeres Dokument"
WORD_BLANK_DOCUMENT_CLASS = "NetUIListViewItem"

WORD_BODY_ID = "Body"
WORD_SAVE_ID = "FileSave"


_word_main_handle = None
_word_process_id = None


def get_execution_environment():
    return _common_get_execution_environment(
        route_host="192.168.178.30",
        route_port=26500,
    )


def generate_test_id():
    return _common_generate_test_id()


def _get_word_main_window(timeout=15):
    global _word_main_handle
    global _word_process_id

    if _word_main_handle:
        try:
            window = Desktop(
                backend="uia"
            ).window(
                handle=_word_main_handle
            )

            if window.exists(timeout=1):
                return window
        except Exception:
            pass

    window = _wait_for_window(
        class_name=WORD_MAIN_CLASS,
        control_type="Window",
        timeout=timeout,
    )

    wrapper = window.wrapper_object()

    _word_main_handle = wrapper.handle
    _word_process_id = _get_window_process_id(
        wrapper
    )

    return window


def _get_word_body(window):
    body = window.child_window(
        auto_id=WORD_BODY_ID,
        control_type="Edit",
    )

    if not body.exists(timeout=1):
        return None

    return body


def _ensure_word_document(
    window,
    timeout=15
):
    body = _get_word_body(window)

    if body is not None:
        return body

    blank_document = window.child_window(
        title=WORD_BLANK_DOCUMENT_NAME,
        control_type="ListItem",
        class_name=WORD_BLANK_DOCUMENT_CLASS,
    )

    if not blank_document.exists(timeout=3):
        raise RuntimeError(
            "Blank Word document template "
            "could not be found."
        )

    blank_document.click_input()

    deadline = time.perf_counter() + float(
        timeout
    )

    while time.perf_counter() < deadline:
        body = _get_word_body(window)

        if body is not None:
            return body

        time.sleep(0.25)

    raise RuntimeError(
        "Word document body did not become ready."
    )


def open_word(timeout=20):
    if not _is_process_running(
        WORD_PROCESS
    ):
        subprocess.Popen(
            f'start "" {WORD_COMMAND}',
            shell=True,
        )

    window = _get_word_main_window(
        timeout
    )

    window.wait(
        "visible enabled ready",
        timeout=timeout,
    )

    _ensure_word_document(
        window,
        timeout,
    )

    return True


def write_word_text(text):
    window = _get_word_main_window()

    body = _ensure_word_document(
        window
    )

    body.set_focus()

    send_keys(
        str(text),
        with_spaces=True,
        pause=0.02,
    )

    return True


def read_word_text():
    window = _get_word_main_window()

    body = _ensure_word_document(
        window
    )

    try:
        text_pattern = body.iface_text

        document_range = (
            text_pattern.DocumentRange
        )

        value = document_range.GetText(-1)

    except Exception as exc:
        raise RuntimeError(
            "Could not read Word document text "
            f"using UI Automation Text Pattern: {exc}"
        )

    return str(value or "").rstrip("\r\n")


def close_word(timeout=10):
    global _word_main_handle
    global _word_process_id

    result = _close_process(
        WORD_PROCESS,
        timeout,
    )

    _word_main_handle = None
    _word_process_id = None

    return result


def ensure_word_closed(timeout=10):
    if not _is_process_running(
        WORD_PROCESS
    ):
        return True

    return close_word(timeout)

def measure_word_startup(
    timeout=30,
    poll_interval=0.005
):
    """
    Measures a cold Word startup.

    Detects:
        loading window: MsoSplash
        main window:    OpusApp

    Returns startup performance metrics.
    """

    global _word_main_handle
    global _word_process_id

    if _is_process_running(WORD_PROCESS):
        raise RuntimeError(
            "Word is already running. "
            "A cold startup measurement requires Word to be closed."
        )

    existing_handles = set()

    def collect_existing_window(hwnd, _):
        try:
            existing_handles.add(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(
        collect_existing_window,
        None
    )

    observed_handles = set()
    observed_windows = []

    loading_seen = None
    main_seen = None

    loading_pid = None
    main_pid = None
    main_handle = None
    main_title = ""

    start_time = perf_counter()

    subprocess.Popen(
        f'start "" {WORD_COMMAND}',
        shell=True
    )

    deadline = (
        start_time
        + float(timeout)
    )

    while perf_counter() < deadline:

        def inspect_window(hwnd, _):
            nonlocal loading_seen
            nonlocal main_seen
            nonlocal loading_pid
            nonlocal main_pid
            nonlocal main_handle
            nonlocal main_title

            try:
                if hwnd in existing_handles:
                    return

                if not win32gui.IsWindowVisible(hwnd):
                    return

                class_name = (
                    win32gui.GetClassName(hwnd)
                    or ""
                )

                if class_name not in (
                    WORD_LOADING_CLASS,
                    WORD_MAIN_CLASS,
                ):
                    return

                detected_at = perf_counter()

                title = (
                    win32gui.GetWindowText(hwnd)
                    or ""
                )

                _, process_id = (
                    win32process.GetWindowThreadProcessId(
                        hwnd
                    )
                )

                if hwnd not in observed_handles:
                    observed_handles.add(hwnd)

                    observed_windows.append({
                        "firstSeenSeconds": round(
                            detected_at - start_time,
                            3
                        ),
                        "name": title,
                        "className": class_name,
                        "processId": process_id,
                        "handle": hwnd,
                    })

                if (
                    class_name == WORD_LOADING_CLASS
                    and loading_seen is None
                ):
                    loading_seen = detected_at
                    loading_pid = process_id

                if (
                    class_name == WORD_MAIN_CLASS
                    and main_seen is None
                ):
                    if (
                        loading_pid is not None
                        and process_id != loading_pid
                    ):
                        return

                    main_seen = detected_at
                    main_pid = process_id
                    main_handle = hwnd
                    main_title = title

            except Exception:
                pass

        win32gui.EnumWindows(
            inspect_window,
            None
        )

        if main_seen is not None:
            break

        time.sleep(
            float(poll_interval)
        )

    if main_seen is None:
        raise RuntimeError(
            f"Word main window "
            f"({WORD_MAIN_CLASS}) was not detected "
            f"within {timeout} seconds."
        )

    # Validate that the detected native window can also
    # be accessed through UI Automation.
    main_window = Desktop(
        backend="uia"
    ).window(
        handle=main_handle
    )

    main_window.wait(
        "visible enabled",
        timeout=10
    )

    _word_main_handle = main_handle
    _word_process_id = main_pid

    startup_to_main = round(
        main_seen - start_time,
        3
    )

    if loading_seen is not None:
        startup_to_loading = round(
            loading_seen - start_time,
            3
        )

        loading_to_main = round(
            main_seen - loading_seen,
            3
        )

        startup_path = "LOADING_TO_MAIN"

    else:
        startup_to_loading = ""
        loading_to_main = ""
        startup_path = "DIRECT_MAIN"

    return {
        "startupToLoadingSeconds": startup_to_loading,
        "loadingToMainSeconds": loading_to_main,
        "startupToMainSeconds": startup_to_main,
        "startupPath": startup_path,

        "windowTitle": main_title,
        "windowHandle": main_handle,
        "processId": main_pid,

        "observedWindows": observed_windows,
    }

def save_word_file(
    file_path,
    timeout=20
):
    window = _get_word_main_window(
        timeout
    )

    return _save_office_file(
        window=window,
        file_path=file_path,
        timeout=timeout,
    )

def save_word_file_as(
    file_path,
    timeout=20
):
    window = _get_word_main_window(
        timeout
    )

    return _save_office_file_as(
        window=window,
        file_path=file_path,
        timeout=timeout
    ) 



def create_blank_word_document(
    timeout=15
):
    window = _get_word_main_window(
        timeout=timeout
    )

    _create_new_office_document(
        window=window,
        blank_item_title="Leeres Dokument",
        timeout=timeout,
    )

    deadline = time.time() + float(timeout)

    while time.time() < deadline:
        try:
            window = _get_word_main_window(
                timeout=2
            )

            body = window.child_window(
                auto_id=WORD_BODY_ID,
                control_type="Edit"
            )

            if body.exists(timeout=0.5):
                return True

        except Exception:
            pass

        time.sleep(0.2)

    raise RuntimeError(
        "New blank Word document was not ready "
        f"within {timeout} seconds."
    )               