import win32gui

from pywinauto import Desktop

from OfficeCommon import (
    control_metadata,
    inspect_window_controls,
    get_window_process_id,
)


ROBOT_LIBRARY_SCOPE = "GLOBAL"


def inspect_foreground_window_controls():
    """
    Inspects the currently focused top-level window.

    Useful for discovering stable UI Automation locators
    in Outlook, Excel, PowerPoint, Word, etc.
    """

    handle = win32gui.GetForegroundWindow()

    if not handle:
        raise RuntimeError(
            "No foreground window could be detected."
        )

    desktop = Desktop(backend="uia")

    window = desktop.window(
        handle=handle
    )

    return inspect_window_controls(
        window,
        "FOREGROUND WINDOW"
    )


def inspect_window_by_handle(handle):
    """
    Inspects a top-level window using its native Windows handle.
    """

    if handle is None:
        raise ValueError(
            "handle must not be None."
        )

    desktop = Desktop(backend="uia")

    window = desktop.window(
        handle=int(handle)
    )

    if not window.exists(timeout=1):
        raise RuntimeError(
            f"Window with handle {handle} was not found."
        )

    return inspect_window_controls(
        window,
        f"WINDOW HANDLE {handle}"
    )


def inspect_process_windows(process_id):
    """
    Returns metadata for all visible top-level windows
    belonging to the specified Windows process ID.
    """

    process_id = int(process_id)

    desktop = Desktop(backend="uia")

    result = []

    for wrapper in desktop.windows():
        try:
            if not wrapper.is_visible():
                continue

            if (
                get_window_process_id(wrapper)
                != process_id
            ):
                continue

            result.append(
                control_metadata(wrapper)
            )

        except Exception:
            continue

    return result


def inspect_foreground_process_windows():
    """
    Detects the foreground window, determines its process ID,
    and returns all visible top-level windows belonging to
    the same process.
    """

    handle = win32gui.GetForegroundWindow()

    if not handle:
        raise RuntimeError(
            "No foreground window could be detected."
        )

    desktop = Desktop(backend="uia")

    window = desktop.window(
        handle=handle
    )

    wrapper = window.wrapper_object()

    process_id = get_window_process_id(
        wrapper
    )

    if process_id is None:
        raise RuntimeError(
            "Could not determine process ID "
            "of foreground window."
        )

    return inspect_process_windows(
        process_id
    )


def get_foreground_window_metadata():
    """
    Returns metadata for the currently focused window.
    """

    handle = win32gui.GetForegroundWindow()

    if not handle:
        raise RuntimeError(
            "No foreground window could be detected."
        )

    desktop = Desktop(backend="uia")

    window = desktop.window(
        handle=handle
    )

    return control_metadata(
        window.wrapper_object()
    )