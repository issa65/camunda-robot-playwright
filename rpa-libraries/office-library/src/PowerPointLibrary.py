import subprocess
import time

from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from OfficeCommon import (
    close_process,
    generate_test_id as _generate_test_id,
    get_execution_environment as _get_execution_environment,
    save_office_file_as as _save_office_file_as,
)


# ============================================================
# PowerPoint constants
# ============================================================

POWERPOINT_PROCESS = "POWERPNT.EXE"
POWERPOINT_COMMAND = "powerpnt"

POWERPOINT_LOADING_CLASS = "MsoSplash"
POWERPOINT_MAIN_CLASS = "PPTFrameClass"

POWERPOINT_START_HOME_ID = "PlaceTabHomeContent"

POWERPOINT_BLANK_PRESENTATION_NAME = "Leere Präsentation"
POWERPOINT_BLANK_PRESENTATION_CLASS = "NetUIListViewItem"

POWERPOINT_WORKSPACE_ID = "3244"
POWERPOINT_TEXT_AREA_ID = "Text Area"

POWERPOINT_SAVE_ID = "FileSave"


def generate_test_id():
    return _generate_test_id()


def get_execution_environment():
    return _get_execution_environment()

# ============================================================
# Internal helpers
# ============================================================

def _get_powerpoint_main_window(timeout=20):
    """
    Returns the visible PowerPoint main window.
    """

    desktop = Desktop(
        backend="uia"
    )

    deadline = (
        time.perf_counter()
        + float(timeout)
    )

    while time.perf_counter() < deadline:

        for wrapper in desktop.windows():

            try:
                if not wrapper.is_visible():
                    continue

                if (
                    wrapper.element_info.class_name
                    != POWERPOINT_MAIN_CLASS
                ):
                    continue

                return desktop.window(
                    handle=wrapper.handle
                )

            except Exception:
                continue

        time.sleep(0.1)

    raise RuntimeError(
        "PowerPoint main window was not detected "
        f"within {timeout} seconds."
    )


def _ensure_powerpoint_presentation(
    window,
    timeout=20
):
    """
    Ensures that a PowerPoint presentation is open.

    If PowerPoint is currently on its start screen,
    clicks 'Leere Präsentation'.
    """

    # ---------------------------------------------------------
    # Prüfen, ob bereits eine Präsentation geöffnet ist
    # ---------------------------------------------------------
    workspace = window.child_window(
        auto_id=POWERPOINT_WORKSPACE_ID,
        control_type="Pane"
    )

    try:
        if workspace.exists(timeout=1):

            # Auf der PowerPoint-Startseite kann 3244 ebenfalls
            # vorhanden sein. Deshalb zusätzlich prüfen, ob die
            # Startseite sichtbar ist.
            start_home = window.child_window(
                auto_id=POWERPOINT_START_HOME_ID,
                control_type="Group"
            )

            if not start_home.exists(timeout=0.5):
                return window

    except Exception:
        pass


    # ---------------------------------------------------------
    # PowerPoint Startseite prüfen
    # ---------------------------------------------------------
    start_home = window.child_window(
        auto_id=POWERPOINT_START_HOME_ID,
        control_type="Group"
    )

    if not start_home.exists(timeout=2):
        return window


    # ---------------------------------------------------------
    # Leere Präsentation öffnen
    # ---------------------------------------------------------
    blank_presentation = window.child_window(
        title=POWERPOINT_BLANK_PRESENTATION_NAME,
        control_type="ListItem",
        class_name=POWERPOINT_BLANK_PRESENTATION_CLASS
    )

    blank_presentation.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    blank_presentation.click_input()


    # ---------------------------------------------------------
    # Warten, bis Startseite verschwindet
    # ---------------------------------------------------------
    deadline = (
        time.perf_counter()
        + float(timeout)
    )

    while time.perf_counter() < deadline:

        if not start_home.exists(timeout=0.1):
            return _get_powerpoint_main_window(
                timeout=timeout
            )

        time.sleep(0.1)

    raise RuntimeError(
        "Blank PowerPoint presentation "
        "did not open."
    )


# ============================================================
# Process handling
# ============================================================

def ensure_powerpoint_closed(timeout=10):
    """
    Ensures that PowerPoint is not running.
    """

    return close_process(
        POWERPOINT_PROCESS,
        timeout=timeout,
    )


def close_powerpoint(timeout=10):
    """
    Closes PowerPoint.
    """

    return close_process(
        POWERPOINT_PROCESS,
        timeout=timeout,
    )


# ============================================================
# Open PowerPoint
# ============================================================

def open_powerpoint(timeout=20):
    """
    Starts PowerPoint and returns the main window.
    """

    subprocess.Popen(
        f'start "" {POWERPOINT_COMMAND}',
        shell=True
    )

    return _get_powerpoint_main_window(
        timeout=timeout
    )


# ============================================================
# Presentation
# ============================================================

def create_blank_powerpoint_presentation(
    timeout=20
):
    """
    Ensures that a blank PowerPoint presentation is open.
    """

    window = _get_powerpoint_main_window(
        timeout=timeout
    )

    _ensure_powerpoint_presentation(
        window=window,
        timeout=timeout,
    )

    return True


# ============================================================
# Save
# ============================================================

def save_powerpoint_file_as(
    file_path,
    timeout=20
):
    """
    Saves the currently open PowerPoint presentation using:

        File
        -> Save As
        -> Browse
        -> target path
        -> Save

    Uses the generic Office implementation.
    """

    window = _get_powerpoint_main_window(
        timeout=timeout
    )

    window = _ensure_powerpoint_presentation(
        window=window,
        timeout=timeout,
    )

    return _save_office_file_as(
        window=window,
        file_path=file_path,
        timeout=timeout,
    )


# ============================================================
# Common test metadata
# ============================================================

def get_powerpoint_execution_environment():
    """
    Returns execution environment information.
    """

    return get_execution_environment()


def generate_powerpoint_test_id():
    """
    Generates a test run ID.
    """

    return generate_test_id()