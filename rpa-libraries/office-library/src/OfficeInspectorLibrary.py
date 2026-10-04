import win32gui

from pywinauto import Desktop

from OfficeCommon import (
    control_metadata,
    inspect_window_controls,
    get_window_process_id,
    inspect_office_startup as _inspect_office_startup_common,
    save_locator_catalog as _save_locator_catalog,
    inspect_office_startup as _inspect_office_startup_common,
    close_process,
)


ROBOT_LIBRARY_SCOPE = "GLOBAL"




def inspect_foreground_window_controls(
    output_file=None,
    application="",
    unique_only=False
):
    """
    Inspects the currently focused top-level window.

    Optionally saves stable locator metadata to a JSON file.
    """

    handle = win32gui.GetForegroundWindow()

    if not handle:
        raise RuntimeError(
            "No foreground window could be detected."
        )

    desktop = Desktop(
        backend="uia"
    )

    window = desktop.window(
        handle=handle
    )

    controls = inspect_window_controls(
        window,
        "FOREGROUND WINDOW"
    )

    if output_file:
        _save_locator_catalog(
            controls=controls,
            file_path=output_file,
            application=application,
            window_label="FOREGROUND WINDOW",
            unique_only=unique_only,
        )

    return controls


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


def inspect_office_startup(
    application,
    timeout=10,
    poll_interval=0.005
):
    """
    Starts an Office application and records newly created
    visible top-level windows.

    Intended for discovering splash/loading and main window
    classes.

    Examples:
        Inspect Office Startup    word
        Inspect Office Startup    excel
        Inspect Office Startup    powerpoint
        Inspect Office Startup    outlook
    """

    return _inspect_office_startup_common(
        application=application,
        timeout=timeout,
        poll_interval=poll_interval,
    )

def ensure_powerpoint_closed(timeout=10):
    """
    Ensures that no PowerPoint process is running.
    """

    return close_process(
        "POWERPNT.EXE",
        timeout=timeout,
    )       