import getpass
import socket
import win32gui
import win32con
import win32process
import subprocess
import time
import json
import win32clipboard
from pathlib import Path
from pywinauto.keyboard import send_keys
from time import perf_counter
from pywinauto import Desktop
from robot.api import logger
from uuid import uuid4

# ============================================================
# Environment
# ============================================================

def get_execution_environment(
    route_host="192.168.178.30",
    route_port=26500
):
    """
    Returns information about the machine executing the RPA test.

    route_host / route_port are only used to determine which local
    network interface Windows would use to reach the target.
    No application-level connection is established.
    """

    username = getpass.getuser()
    hostname = socket.gethostname()

    ip_address = ""

    sock = socket.socket(
        socket.AF_INET,
        socket.SOCK_DGRAM
    )

    try:
        sock.connect(
            (str(route_host), int(route_port))
        )
        ip_address = sock.getsockname()[0]

    except OSError:
        try:
            ip_address = socket.gethostbyname(hostname)
        except OSError:
            ip_address = ""

    finally:
        sock.close()

    return {
        "username": username,
        "hostname": hostname,
        "ipAddress": ip_address,

        "workerUser": username,
        "workerHost": hostname,
        "workerIpAddress": ip_address,
}

# ============================================================
# Timing
# ============================================================

def monotonic_time():
    """
    Returns a high-resolution monotonic timestamp.
    Suitable for performance measurements.
    """

    return perf_counter()


def elapsed_seconds(start_time, end_time=None):
    """
    Calculates elapsed seconds between two monotonic timestamps.
    """

    if end_time is None:
        end_time = perf_counter()

    return round(
        float(end_time) - float(start_time),
        3
    )


# ============================================================
# Process handling
# ============================================================

def is_process_running(process_name):
    """
    Returns True when a Windows process with the given image name
    is currently running.

    Example:
        is_process_running("OUTLOOK.EXE")
        is_process_running("EXCEL.EXE")
        is_process_running("POWERPNT.EXE")
    """

    process_name = str(process_name).strip()

    if not process_name:
        raise ValueError(
            "process_name must not be empty."
        )

    result = subprocess.run(
        [
            "tasklist",
            "/FI",
            f"IMAGENAME eq {process_name}",
            "/FO",
            "CSV",
            "/NH"
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        check=False
    )

    return (
        process_name.upper().encode()
        in result.stdout.upper()
    )


def generate_test_id():
    """
    Generates a unique ID for an E2E email test.
    """
    return uuid4().hex[:8].upper()



def close_process(process_name, timeout=10):
    """
    Force-closes a Windows process and waits until it has stopped.

    Intended for dedicated RPA test machines.
    """

    process_name = str(process_name).strip()

    if not process_name:
        raise ValueError(
            "process_name must not be empty."
        )

    if not is_process_running(process_name):
        return True

    subprocess.run(
        [
            "taskkill",
            "/IM",
            process_name,
            "/T",
            "/F"
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False
    )

    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        if not is_process_running(process_name):
            return True

        time.sleep(0.25)

    raise RuntimeError(
        f"Process {process_name!r} could not be "
        f"closed within {timeout} seconds."
    )


# ============================================================
# Window metadata
# ============================================================

def get_window_class(wrapper):
    try:
        return (
            wrapper.element_info.class_name
            or ""
        )
    except Exception:
        return ""


def get_window_control_type(wrapper):
    try:
        return (
            wrapper.element_info.control_type
            or ""
        )
    except Exception:
        return ""


def get_window_process_id(wrapper):
    try:
        return wrapper.element_info.process_id
    except Exception:
        return None


def get_window_handle(wrapper):
    try:
        return wrapper.handle
    except Exception:
        try:
            return wrapper.element_info.handle
        except Exception:
            return None


def control_metadata(control):
    """
    Returns UI Automation metadata for a window or control.

    This function is generic and can be used for Outlook,
    Excel, PowerPoint, Word, etc.
    """

    try:
        info = control.element_info
    except Exception:
        info = None

    try:
        name = control.window_text() or ""
    except Exception:
        name = ""

    try:
        rectangle = str(control.rectangle())
    except Exception:
        rectangle = ""

    try:
        visible = bool(control.is_visible())
    except Exception:
        visible = None

    handle = get_window_handle(control)

    return {
        "name": name,
        "handle": handle,

        "automationId": (
            getattr(info, "automation_id", "") or ""
            if info is not None
            else ""
        ),

        "controlType": (
            getattr(info, "control_type", "") or ""
            if info is not None
            else ""
        ),

        "className": (
            getattr(info, "class_name", "") or ""
            if info is not None
            else ""
        ),

        "processId": (
            getattr(info, "process_id", None)
            if info is not None
            else None
        ),

        "runtimeId": (
            getattr(info, "runtime_id", None)
            if info is not None
            else None
        ),

        "rectangle": rectangle,
        "visible": visible
    }


# ============================================================
# Window lookup
# ============================================================

def find_windows_by_class(
    class_name,
    control_type="Window",
    process_id=None,
    visible_only=True
):
    """
    Returns all windows matching the supplied stable attributes.

    No title regex is used.
    """

    desktop = Desktop(backend="uia")
    matches = []

    for wrapper in desktop.windows():
        try:
            if (
                visible_only
                and not wrapper.is_visible()
            ):
                continue

            if (
                get_window_class(wrapper)
                != class_name
            ):
                continue

            if (
                control_type is not None
                and get_window_control_type(wrapper)
                != control_type
            ):
                continue

            if (
                process_id is not None
                and get_window_process_id(wrapper)
                != int(process_id)
            ):
                continue

            matches.append(wrapper)

        except Exception:
            continue

    return matches


def find_window_by_class(
    class_name,
    control_type="Window",
    process_id=None,
    visible_only=True
):
    """
    Returns exactly one matching window.

    If multiple windows match, this function deliberately fails
    instead of silently selecting an ambiguous UI element.
    """

    matches = find_windows_by_class(
        class_name=class_name,
        control_type=control_type,
        process_id=process_id,
        visible_only=visible_only
    )

    if not matches:
        return None

    if len(matches) > 1:
        raise RuntimeError(
            "Window locator is ambiguous: "
            f"class_name={class_name!r}, "
            f"control_type={control_type!r}, "
            f"process_id={process_id!r}. "
            f"Found {len(matches)} matching windows."
        )

    wrapper = matches[0]

    return Desktop(
        backend="uia"
    ).window(
        handle=wrapper.handle
    )


def wait_for_window(
    class_name,
    control_type="Window",
    process_id=None,
    timeout=30,
    poll_interval=0.05
):
    """
    Waits until exactly one matching window is available.
    """

    deadline = (
        perf_counter()
        + float(timeout)
    )

    while perf_counter() < deadline:

        window = find_window_by_class(
            class_name=class_name,
            control_type=control_type,
            process_id=process_id,
            visible_only=True
        )

        if window is not None:
            return window

        time.sleep(
            float(poll_interval)
        )

    raise RuntimeError(
        "Window was not detected within "
        f"{timeout} seconds: "
        f"class_name={class_name!r}, "
        f"control_type={control_type!r}, "
        f"process_id={process_id!r}"
    )


# ============================================================
# Generic control inspector
# ============================================================

def inspect_window_controls(
    window,
    label="WINDOW"
):
    """
    Inspects all descendants of an already uniquely identified
    window.

    Reports whether automation_id + control_type is unique within
    that window.
    """

    try:
        wrapper = window.wrapper_object()
    except Exception:
        wrapper = window

    controls = []

    for control in wrapper.descendants():
        try:
            controls.append(
                control_metadata(control)
            )
        except Exception:
            continue

    locator_counts = {}

    for metadata in controls:
        automation_id = (
            metadata["automationId"]
        )

        control_type = (
            metadata["controlType"]
        )

        if (
            not automation_id
            or not control_type
        ):
            continue

        key = (
            automation_id,
            control_type
        )

        locator_counts[key] = (
            locator_counts.get(key, 0)
            + 1
        )

    window_metadata = (
        control_metadata(wrapper)
    )

    logger.console(
        f"\n=== UI INSPECTOR: {label} ==="
    )

    logger.console(
        "WINDOW | "
        f"name={window_metadata['name']!r} | "
        f"handle={window_metadata['handle']!r} | "
        f"auto_id={window_metadata['automationId']!r} | "
        f"control_type={window_metadata['controlType']!r} | "
        f"class={window_metadata['className']!r} | "
        f"pid={window_metadata['processId']!r}"
    )

    for metadata in controls:

        automation_id = (
            metadata["automationId"]
        )

        control_type = (
            metadata["controlType"]
        )

        unique_locator = False

        if (
            automation_id
            and control_type
        ):
            unique_locator = (
                locator_counts.get(
                    (
                        automation_id,
                        control_type
                    ),
                    0
                )
                == 1
            )

        metadata[
            "uniqueAutoIdLocator"
        ] = unique_locator

        logger.console(
            "CONTROL | "
            f"name={metadata['name']!r} | "
            f"auto_id={automation_id!r} | "
            f"control_type={control_type!r} | "
            f"class={metadata['className']!r} | "
            f"handle={metadata['handle']!r} | "
            f"pid={metadata['processId']!r} | "
            f"unique_auto_id_locator="
            f"{unique_locator!r}"
        )

    logger.console(
        f"=== END UI INSPECTOR: "
        f"{label} ===\n"
    )

    return controls

OFFICE_APPLICATIONS = {
    "word": {
        "command": "winword",
        "process": "WINWORD.EXE",
    },
    "excel": {
        "command": "excel",
        "process": "EXCEL.EXE",
    },
    "powerpoint": {
        "command": "powerpnt",
        "process": "POWERPNT.EXE",
    },
    "outlook": {
        "command": "outlook",
        "process": "OUTLOOK.EXE",
    },
}

# ============================================================
# Startup window observation
# ============================================================

def observe_startup_windows(
    start_command,
    relevant_classes,
    stop_class_name=None,
    timeout=30,
    poll_interval=0.05
):
    """
    Starts an application and records newly observed visible
    top-level windows matching the supplied class names.

    Generic for Office applications.

    Example:
        relevant_classes=[
            "MsoSplash",
            "rctrl_renwnd32"
        ]

    The function does not contain Outlook/Excel/PowerPoint
    specific constants.
    """

    if isinstance(
        relevant_classes,
        str
    ):
        relevant_classes = [
            relevant_classes
        ]

    relevant_classes = set(
        relevant_classes
    )

    desktop = Desktop(
        backend="uia"
    )

    existing_handles = {
        wrapper.handle
        for wrapper
        in desktop.windows()
    }

    observed_handles = set()
    observed_windows = []

    start_time = perf_counter()

    subprocess.Popen(
        str(start_command),
        shell=True
    )

    deadline = (
        start_time
        + float(timeout)
    )

    stop_window = None

    while perf_counter() < deadline:

        now = perf_counter()

        for wrapper in desktop.windows():
            try:
                if not wrapper.is_visible():
                    continue

                handle = wrapper.handle

                if handle in existing_handles:
                    continue

                class_name = (
                    get_window_class(wrapper)
                )

                if (
                    relevant_classes
                    and class_name
                    not in relevant_classes
                ):
                    continue

                if (
                    handle
                    not in observed_handles
                ):
                    observed_handles.add(
                        handle
                    )

                    metadata = (
                        control_metadata(wrapper)
                    )

                    metadata[
                        "firstSeenSeconds"
                    ] = round(
                        now - start_time,
                        3
                    )

                    observed_windows.append(
                        metadata
                    )

                if (
                    stop_class_name
                    and class_name
                    == stop_class_name
                ):
                    stop_window = (
                        Desktop(
                            backend="uia"
                        ).window(
                            handle=handle
                        )
                    )
                    break

            except Exception:
                continue

        if stop_window is not None:
            break

        time.sleep(
            float(poll_interval)
        )

    return {
        "startedAt": start_time,
        "elapsedSeconds": (
            elapsed_seconds(start_time)
        ),
        "observedWindows": (
            observed_windows
        ),
        "stopWindowHandle": (
            None
            if stop_window is None
            else stop_window.handle
        )
    }

def inspect_office_startup(
    application,
    timeout=10,
    poll_interval=0.005
):
    """
    Starts an Office application and discovers newly created
    visible top-level windows using fast Win32 enumeration.

    Intended for discovering splash/loading and main window
    classes before stable locators are known.
    """

    app_key = str(application).strip().lower()

    if app_key not in OFFICE_APPLICATIONS:
        raise ValueError(
            f"Unsupported Office application: {application!r}. "
            f"Supported applications: "
            f"{', '.join(OFFICE_APPLICATIONS.keys())}"
        )

    config = OFFICE_APPLICATIONS[app_key]

    command = config["command"]
    process_name = config["process"]

    if is_process_running(process_name):
        raise RuntimeError(
            f"{process_name} is already running. "
            "Close the application before startup inspection."
        )

    existing_handles = set()

    def collect_existing_window(hwnd, _):
        existing_handles.add(hwnd)

    win32gui.EnumWindows(
        collect_existing_window,
        None
    )

    observed_handles = set()
    observed_windows = []

    start_time = perf_counter()

    logger.console(
        f"\n=== OFFICE STARTUP INSPECTOR: "
        f"{app_key.upper()} ==="
    )

    logger.console(
        f"command={command!r} | "
        f"process={process_name!r}"
    )

    subprocess.Popen(
        f'start "" {command}',
        shell=True
    )

    deadline = (
        start_time
        + float(timeout)
    )

    while perf_counter() < deadline:

        now = perf_counter()

        def inspect_window(hwnd, _):
            try:
                if hwnd in existing_handles:
                    return

                if hwnd in observed_handles:
                    return

                if not win32gui.IsWindowVisible(hwnd):
                    return

                title = (
                    win32gui.GetWindowText(hwnd)
                    or ""
                )

                class_name = (
                    win32gui.GetClassName(hwnd)
                    or ""
                )

                _, process_id = (
                    win32process.GetWindowThreadProcessId(
                        hwnd
                    )
                )

                observed_handles.add(hwnd)

                first_seen = round(
                    now - start_time,
                    3
                )

                metadata = {
                    "firstSeenSeconds": first_seen,
                    "name": title,
                    "className": class_name,
                    "processId": process_id,
                    "handle": hwnd,
                }

                observed_windows.append(
                    metadata
                )

                logger.console(
                    f"first_seen={first_seen}s | "
                    f"name={title!r} | "
                    f"class={class_name!r} | "
                    f"pid={process_id} | "
                    f"handle={hwnd}"
                )

            except Exception:
                pass

        win32gui.EnumWindows(
            inspect_window,
            None
        )

        time.sleep(
            float(poll_interval)
        )

    logger.console(
        f"=== END OFFICE STARTUP INSPECTOR: "
        f"{app_key.upper()} ===\n"
    )

    return {
        "application": app_key,
        "processName": process_name,
        "observedWindows": observed_windows,
    }

def save_locator_catalog(
    controls,
    file_path,
    application="",
    window_label="",
    unique_only=False
):
    """
    Saves stable UI Automation locator metadata to JSON.

    Runtime-specific data such as handle, processId and runtimeId
    is deliberately excluded.
    """

    locator_entries = []

    for control in controls:
        if (
            unique_only
            and not control.get("uniqueAutoIdLocator", False)
        ):
            continue

        locator_entries.append({
            "name": control.get("name", ""),
            "automationId": control.get("automationId", ""),
            "controlType": control.get("controlType", ""),
            "className": control.get("className", ""),
            "uniqueAutoIdLocator": control.get(
                "uniqueAutoIdLocator",
                False
            ),
        })

    catalog = {
        "application": str(application),
        "windowLabel": str(window_label),
        "locators": locator_entries,
    }

    path = Path(file_path)

    path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    with path.open(
        "w",
        encoding="utf-8"
    ) as file:
        json.dump(
            catalog,
            file,
            ensure_ascii=False,
            indent=2,
        )

    return str(path)

def save_office_file(
    window,
    file_path,
    timeout=20
):
    """
    Saves the currently open Office document to the supplied path.

    Generic for Word, Excel and PowerPoint.

    Uses the Office Save As shortcut and handles the standard
    Windows Save As dialog.

    Returns metadata about the saved file.
    """

    if not file_path:
        raise ValueError(
            "file_path must not be empty."
        )

    path = Path(str(file_path))

    parent = path.parent

    if not parent.exists():
        parent.mkdir(
            parents=True,
            exist_ok=True
        )

    try:
        wrapper = window.wrapper_object()
    except Exception:
        wrapper = window

    process_id = get_window_process_id(
        wrapper
    )

    wrapper.set_focus()

    # F12 opens "Save As" in Word/Excel/PowerPoint.
    send_keys("{F12}")

    desktop = Desktop(
        backend="uia"
    )

    deadline = perf_counter() + float(
        timeout
    )

    save_dialog = None

    while perf_counter() < deadline:

        for candidate in desktop.windows():

            try:
                if not candidate.is_visible():
                    continue

                if (
                    process_id is not None
                    and get_window_process_id(
                        candidate
                    )
                    != process_id
                ):
                    continue

                class_name = get_window_class(
                    candidate
                )

                if class_name != "#32770":
                    continue

                save_dialog = desktop.window(
                    handle=candidate.handle
                )

                break

            except Exception:
                continue

        if save_dialog is not None:
            break

        time.sleep(0.1)

    if save_dialog is None:
        raise RuntimeError(
            "Office Save As dialog was not detected."
        )

    save_dialog.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_name_field = save_dialog.child_window(
        auto_id="1001",
        control_type="Edit"
    )

    file_name_field.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_name_field.set_edit_text(
        str(path)
    )

    save_button = save_dialog.child_window(
        auto_id="1",
        control_type="Button"
    )

    save_button.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    save_button.click_input()

    deadline = perf_counter() + float(
        timeout
    )

    while perf_counter() < deadline:

        if path.exists():

            return {
                "saved": True,
                "filePath": str(path),
                "fileName": path.name,
                "fileSizeBytes": path.stat().st_size,
            }

        time.sleep(0.2)

    raise RuntimeError(
        f"Office file was not created within "
        f"{timeout} seconds: {path}"
    )

def open_office_save_as_dialog(
    window,
    timeout=15
):
    """
    Opens Office File -> Save As -> Browse.

    Generic for Word, Excel and PowerPoint.
    """

    try:
        wrapper = window.wrapper_object()
    except Exception:
        wrapper = window

    wrapper.set_focus()

    # ---------------------------------------------------------
    # Datei öffnen
    # ---------------------------------------------------------
    file_tab = wrapper.child_window(
        auto_id="FileTabButton"
    )

    file_tab.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_tab.click_input()

    # ---------------------------------------------------------
    # Backstage-Ansicht abwarten
    # ---------------------------------------------------------
    backstage = wrapper.child_window(
        auto_id="BackstageView",
        control_type="Pane"
    )

    backstage.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    # ---------------------------------------------------------
    # Speichern unter
    # ---------------------------------------------------------
    save_as = wrapper.child_window(
        title="Speichern unter",
        control_type="ListItem",
        class_name="NetUIRibbonTab"
    )

    save_as.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    save_as.click_input()

    # ---------------------------------------------------------
    # Save-As-Seite abwarten
    # ---------------------------------------------------------
    save_group = wrapper.child_window(
        auto_id="SaveGroup",
        control_type="Group"
    )

    save_group.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    # ---------------------------------------------------------
    # Durchsuchen
    # ---------------------------------------------------------
    browse_button = wrapper.child_window(
        title="Durchsuchen",
        control_type="Button",
        class_name="NetUISimpleButton"
    )

    browse_button.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    browse_button.click_input()

    return True

def _find_office_save_as_dialog(
    timeout=15
):
    """
    Finds the native Windows Save As dialog opened by
    Word, Excel or PowerPoint.
    """

    desktop = Desktop(
        backend="uia"
    )

    deadline = perf_counter() + float(
        timeout
    )

    while perf_counter() < deadline:

        for wrapper in desktop.windows():
            try:
                if not wrapper.is_visible():
                    continue

                window = desktop.window(
                    handle=wrapper.handle
                )

                file_name = window.child_window(
                    auto_id="1001",
                    control_type="Edit"
                )

                save_button = window.child_window(
                    auto_id="1",
                    control_type="Button"
                )

                if (
                    file_name.exists(timeout=0.05)
                    and save_button.exists(timeout=0.05)
                ):
                    return window

            except Exception:
                continue

        time.sleep(0.1)

    raise RuntimeError(
        "Windows Save As dialog was not detected."
    )

def save_office_file_as(
    window,
    file_path,
    timeout=20,
    save_as_title="Speichern unter",
    browse_title="Durchsuchen"
):
    """
    Saves an Office document using:

        File
        -> Save As
        -> Browse
        -> Enter complete path into File name
        -> Save

    Generic for Word, Excel and PowerPoint.
    """

    from pathlib import Path

    if not file_path:
        raise ValueError(
            "file_path must not be empty."
        )

    path = Path(
        str(file_path)
    )

    target_directory = path.parent

    # ---------------------------------------------------------
    # Zielordner prüfen / erstellen
    # ---------------------------------------------------------
    try:
        target_directory.mkdir(
            parents=True,
            exist_ok=True
        )

    except Exception as exc:
        raise RuntimeError(
            "Could not access or create target directory "
            f"'{target_directory}': {exc}"
        )

    if not target_directory.exists():
        raise RuntimeError(
            "Target directory is not accessible: "
            f"{target_directory}"
        )

    # ---------------------------------------------------------
    # Office-Hauptfenster
    # ---------------------------------------------------------
    try:
        wrapper = window.wrapper_object()
        handle = wrapper.handle

    except Exception:
        handle = get_window_handle(
            window
        )

    if handle is None:
        raise RuntimeError(
            "Could not determine Office window handle."
        )

    office_window = Desktop(
        backend="uia"
    ).window(
        handle=int(handle)
    )

    office_window.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    office_window.set_focus()

    # ---------------------------------------------------------
    # Datei
    # ---------------------------------------------------------
    file_tab = office_window.child_window(
        auto_id="FileTabButton"
    )

    file_tab.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_tab.click_input()

    # ---------------------------------------------------------
    # Backstage
    # ---------------------------------------------------------
    backstage = office_window.child_window(
        auto_id="BackstageView",
        control_type="Pane"
    )

    backstage.wait(
        "visible",
        timeout=float(timeout)
    )

    # ---------------------------------------------------------
    # Speichern unter
    # ---------------------------------------------------------
    save_as = office_window.child_window(
        title=save_as_title,
        control_type="ListItem"
    )

    save_as.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    save_as.click_input()

    # ---------------------------------------------------------
    # Durchsuchen
    # ---------------------------------------------------------
    browse = office_window.child_window(
        title=browse_title,
        control_type="Button"
    )

    browse.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    browse.click_input()

    # ---------------------------------------------------------
    # Auf Windows "Speichern unter"-Dialog warten
    # ---------------------------------------------------------
    deadline = (
        perf_counter()
        + float(timeout)
    )

    dialog_handle = None

    while perf_counter() < deadline:

        foreground_handle = (
            win32gui.GetForegroundWindow()
        )

        if foreground_handle:

            title = (
                win32gui.GetWindowText(
                    foreground_handle
                )
                or ""
            ).strip()

            if save_as_title.lower() in title.lower():
                dialog_handle = foreground_handle
                break

        time.sleep(0.1)

    if dialog_handle is None:
        raise RuntimeError(
            "Windows Save As dialog was not detected."
        )

    # ---------------------------------------------------------
    # Dialog über UIA ansprechen
    # ---------------------------------------------------------
    save_dialog = Desktop(
        backend="uia"
    ).window(
        handle=int(dialog_handle)
    )

    # ---------------------------------------------------------
    # Vollständiger Zielpfad
    #
    # Beispiel:
    # \\ISSA\Gemeinsame Dateien\test-files\word-1234.docx
    # ---------------------------------------------------------
    expected_path = str(path)

        # ---------------------------------------------------------
    # Dateiname-Feld eindeutig über UIA finden
    #
    # auto_id = 1001
    # control_type = Edit
    # ---------------------------------------------------------
    file_name_spec = save_dialog.child_window(
        auto_id="1001",
        control_type="Edit"
    )

    file_name_spec.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_name_wrapper = (
        file_name_spec.wrapper_object()
    )


    # ---------------------------------------------------------
    # Kompletten Zielpfad vorbereiten
    # ---------------------------------------------------------
    expected_path = str(path)

    logger.info(
        f"Target Save As path: {expected_path}"
    )


    # ---------------------------------------------------------
    # Vollständigen Pfad in Windows-Zwischenablage schreiben
    # ---------------------------------------------------------
    clipboard_set = False

    for attempt in range(10):
        try:
            win32clipboard.OpenClipboard()

            try:
                win32clipboard.EmptyClipboard()

                win32clipboard.SetClipboardText(
                    expected_path,
                    win32clipboard.CF_UNICODETEXT
                )

            finally:
                win32clipboard.CloseClipboard()

            clipboard_set = True
            break

        except Exception:
            time.sleep(0.1)


    if not clipboard_set:
        raise RuntimeError(
            "Could not write target path "
            "to Windows clipboard."
        )


    # ---------------------------------------------------------
    # Clipboard kontrollieren
    # ---------------------------------------------------------
    clipboard_text = ""

    try:
        win32clipboard.OpenClipboard()

        try:
            clipboard_text = (
                win32clipboard.GetClipboardData(
                    win32clipboard.CF_UNICODETEXT
                )
                or ""
            )

        finally:
            win32clipboard.CloseClipboard()

    except Exception as exc:
        raise RuntimeError(
            f"Could not verify Windows clipboard: {exc}"
        )


    if clipboard_text != expected_path:
        raise RuntimeError(
            "Clipboard does not contain expected path. "
            f"Expected: '{expected_path}', "
            f"Actual: '{clipboard_text}'"
        )


    # ---------------------------------------------------------
    # Dateiname-Feld aktivieren
    # ---------------------------------------------------------
    file_name_wrapper.click_input()

    time.sleep(0.2)


    # ---------------------------------------------------------
    # Vorhandenen Word-Vorschlag markieren
    # ---------------------------------------------------------
    file_name_wrapper.type_keys(
        "^a",
        set_foreground=True
    )

    time.sleep(0.1)


    # ---------------------------------------------------------
    # KOMPLETTEN UNC-Pfad einfügen
    #
    # Beispiel:
    # \\ISSA\Gemeinsame Dateien\test-files\word-12345678.docx
    # ---------------------------------------------------------
    file_name_wrapper.type_keys(
        "^v",
        set_foreground=True
    )

    time.sleep(0.5)


    # ---------------------------------------------------------
    # UIA-Wert nur zur Diagnose lesen.
    #
    # Kein Abbruch, falls Windows den Wert hier nicht
    # zurückliefert. Die endgültige Validierung erfolgt
    # anschließend über das Dateisystem.
    # ---------------------------------------------------------
    try:
        current_file_name = str(
            file_name_wrapper
            .iface_value
            .CurrentValue
            or ""
        ).strip()

        logger.info(
            "File name after paste: "
            f"'{current_file_name}'"
        )

    except Exception as exc:
        logger.info(
            "Could not read File name field "
            f"after paste: {exc}"
        )


    # ---------------------------------------------------------
    # Speichern
    # ---------------------------------------------------------
    save_button_spec = save_dialog.child_window(
        auto_id="1",
        control_type="Button"
    )

    save_button_spec.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    save_button_wrapper = (
        save_button_spec.wrapper_object()
    )

    save_button_wrapper.click_input()

    logger.info(
        f"Save button clicked for: {expected_path}"
    )

    # ---------------------------------------------------------
    # Datei tatsächlich im Zielordner prüfen
    # ---------------------------------------------------------
    deadline = (
        perf_counter()
        + float(timeout)
    )

    actual_path = None
    last_error = None

    while perf_counter() < deadline:

        try:
            if path.is_file():

                file_size = (
                    path.stat().st_size
                )

                if file_size > 0:
                    actual_path = path
                    break

        except Exception as exc:
            last_error = exc

        time.sleep(0.5)

    # ---------------------------------------------------------
    # Datei wurde nicht gefunden
    # ---------------------------------------------------------
    if actual_path is None:

        message = (
            f"File was not created within "
            f"{timeout} seconds: {path}"
        )

        if last_error is not None:
            message += (
                " | Last filesystem error: "
                f"{last_error}"
            )

        raise RuntimeError(
            message
        )

    # ---------------------------------------------------------
    # Dateigröße
    # ---------------------------------------------------------
    try:
        file_size_bytes = (
            actual_path.stat().st_size
        )

    except OSError as exc:
        raise RuntimeError(
            "Saved file exists but file size "
            f"could not be determined: {exc}"
        )

    if file_size_bytes <= 0:
        raise RuntimeError(
            "Saved file exists but is empty: "
            f"{actual_path}"
        )

    # ---------------------------------------------------------
    # Ergebnis
    # ---------------------------------------------------------
    return {
        "saved": True,
        "filePath": str(actual_path),
        "fileName": actual_path.name,
        "fileSizeBytes": file_size_bytes,
    }

def create_new_office_document(
    window,
    blank_item_title,
    timeout=15,
    home_content_id="PlaceTabHomeContent"
):
    """
    Creates a new blank Office document.

    Supports both situations:

    1. Office start screen is already visible
       -> click blank document directly

    2. A document is already open
       -> File -> New -> blank document
    """

    office_window = window

    # ---------------------------------------------------------
    # Fall 1:
    # Office befindet sich bereits auf der Startseite
    # ---------------------------------------------------------
    start_home = office_window.child_window(
        auto_id=home_content_id,
        control_type="Group"
    )

    if start_home.exists(timeout=1):

        blank_item = office_window.child_window(
            title=blank_item_title,
            control_type="ListItem",
            class_name="NetUIListViewItem"
        )

        blank_item.wait(
            "visible enabled",
            timeout=float(timeout)
        )

        blank_item.click_input()

        return True


    # ---------------------------------------------------------
    # Fall 2:
    # Bereits ein Dokument geöffnet
    # -> Datei
    # ---------------------------------------------------------
    file_tab = office_window.child_window(
        auto_id="FileTabButton"
    )

    file_tab.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    file_tab.click_input()


    # ---------------------------------------------------------
    # Backstage-Ansicht
    # ---------------------------------------------------------
    backstage = office_window.child_window(
        auto_id="BackstageView",
        control_type="Pane"
    )

    backstage.wait(
        "visible enabled",
        timeout=float(timeout)
    )


    # ---------------------------------------------------------
    # Neu
    # ---------------------------------------------------------
    new_item = office_window.child_window(
        title="Neu",
        control_type="ListItem"
    )

    new_item.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    new_item.click_input()


    # ---------------------------------------------------------
    # Leeres Dokument / Arbeitsmappe / Präsentation
    # ---------------------------------------------------------
    blank_item = office_window.child_window(
        title=blank_item_title,
        control_type="ListItem",
        class_name="NetUIListViewItem"
    )

    blank_item.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    blank_item.click_input()

    return True

def _set_clipboard_text(text):
    """
    Writes Unicode text to the Windows clipboard.
    """

    win32clipboard.OpenClipboard()

    try:
        win32clipboard.EmptyClipboard()

        win32clipboard.SetClipboardText(
            str(text),
            win32clipboard.CF_UNICODETEXT
        )

    finally:
        win32clipboard.CloseClipboard()

def _find_native_child_control(
    parent_handle,
    control_id,
    class_name=None
):
    """
    Finds a native child control by Windows control ID.
    """

    matches = []

    def callback(handle, _):
        try:
            if win32gui.GetDlgCtrlID(handle) != int(control_id):
                return True

            if (
                class_name
                and win32gui.GetClassName(handle) != class_name
            ):
                return True

            matches.append(handle)

        except Exception:
            pass

        return True

    win32gui.EnumChildWindows(
        int(parent_handle),
        callback,
        None
    )

    if not matches:
        raise RuntimeError(
            f"Native control not found: "
            f"id={control_id}, "
            f"class={class_name}"
        )

    return matches[0]               