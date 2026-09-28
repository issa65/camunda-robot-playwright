import getpass
import socket
import subprocess
import time
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