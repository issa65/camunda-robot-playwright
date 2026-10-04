import time
from time import perf_counter
import win32gui
from comtypes import COMError
from datetime import datetime, timezone
from pywinauto.keyboard import send_keys
from robot.api import logger
import subprocess

from pywinauto import Desktop

from OfficeCommon import (
    is_process_running as _is_process_running_common,
    close_process as _close_process,
    get_execution_environment as _common_get_execution_environment,
    get_window_process_id as _get_window_process_id,
    wait_for_window as _wait_for_window,
    observe_startup_windows as _observe_startup_windows,
    control_metadata as _control_metadata,
    inspect_window_controls as _inspect_window_controls,
    generate_test_id as _common_generate_test_id,
)


ROBOT_LIBRARY_SCOPE = "GLOBAL"

OUTLOOK_LOADING_CLASS = "MsoSplash"
OUTLOOK_MAIN_CLASS = "rctrl_renwnd32"

_main_handle = None
_outlook_process_id = None
_compose_handle = None


def _is_outlook_running():
    return _is_process_running_common("OUTLOOK.EXE")


def _get_outlook_main_window(timeout=30):
    global _main_handle
    global _outlook_process_id

    desktop = Desktop(backend="uia")

    # Bereits eindeutig erkanntes Hauptfenster wiederverwenden.

    if _main_handle is not None:
        try:
            window = desktop.window(handle=_main_handle)
            if (
                window.exists(timeout=0.2)
                and window.is_visible()
            ):
                return window

        except Exception:
            _main_handle = None



    # Vor dem Öffnen eines Compose-Fensters sollte genau ein
    # Outlook-Hauptfenster mit dieser Klasse existieren.

    window = _wait_for_window(
        class_name=OUTLOOK_MAIN_CLASS,
        control_type="Window",
        process_id=_outlook_process_id,
        timeout=timeout,
        poll_interval=0.05
    )

    wrapper = window.wrapper_object()
    _main_handle = wrapper.handle
    _outlook_process_id = _get_window_process_id(wrapper)

    return window

def generate_test_id():
    return _common_generate_test_id()
   

def open_outlook(timeout=30):
    global _main_handle
    global _outlook_process_id

    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    window = _get_outlook_main_window(timeout)
    window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    wrapper = window.wrapper_object()
    _main_handle = wrapper.handle
    _outlook_process_id = _get_window_process_id(wrapper)

    return window.window_text()


def measure_outlook_startup(timeout=30):
    """
    Measures Outlook cold startup.
    """
    global _main_handle
    global _outlook_process_id

    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is already running. "
            "Cold startup measurement aborted."
        )

    measured_at = datetime.now(
        timezone.utc
    ).isoformat()

    observation = _observe_startup_windows(
        start_command='start "" outlook',

        relevant_classes=[
            OUTLOOK_LOADING_CLASS,
            OUTLOOK_MAIN_CLASS
        ],

        stop_class_name=OUTLOOK_MAIN_CLASS,
        timeout=timeout,
        poll_interval=0.05
    )

    observed_windows = observation[
        "observedWindows"
    ]

    loading_windows = [
        window
        for window in observed_windows
        if window.get("className")
        == OUTLOOK_LOADING_CLASS

    ]

    main_windows = [
        window
        for window in observed_windows
        if window.get("className")
        == OUTLOOK_MAIN_CLASS

    ]

    # Main Window muss eindeutig gefunden worden sein.

    if len(main_windows) != 1:
        raise RuntimeError(

            "Expected exactly one Outlook main window "
            f"during startup, found {len(main_windows)}."
        )

    # Splash darf entweder fehlen oder genau einmal erscheinen.

    if len(loading_windows) > 1:
        raise RuntimeError(

            "Expected at most one Outlook loading window "
            f"during startup, found {len(loading_windows)}."
        )

    main = main_windows[0]
    startup_to_main = float(
        main["firstSeenSeconds"]
    )

    main_process_id = main.get(
        "processId"
    )

    # --------------------------------------------------------
    # Splash wurde gesehen
    # --------------------------------------------------------

    if loading_windows:
        loading = loading_windows[0]
        loading_process_id = loading.get(

            "processId"
        )
        # Splash und Main müssen zum selben Outlook-Prozess gehören.

        if (
            loading_process_id is not None
            and main_process_id is not None
            and loading_process_id
            != main_process_id

        ):
            raise RuntimeError(
                "Outlook loading window and main window "
                "belong to different processes."
            )

        startup_to_loading = float(
            loading["firstSeenSeconds"]
        )

        loading_to_main = round(
            startup_to_main
            - startup_to_loading,
            3
        )

        startup_path = (
            "LOADING_TO_MAIN"
        )

    # --------------------------------------------------------
    # Main Window erschien direkt
    # --------------------------------------------------------

    else:
        startup_to_loading = ""
        loading_to_main = ""
        startup_path = "DIRECT_MAIN"

    main_handle = main.get("handle")

    if main_handle is None:
        raise RuntimeError(
            "Outlook main window has no native handle."
        )

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

    _main_handle = int(main_handle)
    _outlook_process_id = (
        main_process_id
    )

    return {

        "startupToLoadingSeconds": (
            startup_to_loading
        ),

        "loadingToMainSeconds": (
            loading_to_main
        ),

        "startupToMainSeconds": round(
            startup_to_main,
            3
        ),
        "startupPath": startup_path,
        "windowTitle": (
            main_window.window_text()
        ),

        "windowHandle": (
            _main_handle
        ),

        "processId": (
            _outlook_process_id
        ),

        "startupType": "COLD",
        "measuredAt": measured_at,
        "observedWindows": (
            observed_windows
        )
    }


def prepare_outlook_performance_test():
    """
    Checks whether the environment is ready for a cold-start measurement.
    """
    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is already running. "
            "Close Outlook before starting the performance test."
        )

    return "READY"


def prepare_outlook_ui(timeout=1):
    """
    Handles already visible Outlook/Office startup dialogs.

    Fast path:
    If no popup is currently visible, return immediately instead
    of waiting several seconds for every possible dialog.
    """

    onboarding_handled = dismiss_outlook_onboarding(
        timeout=0.2
    )

    license_popup_handled = dismiss_outlook_license_popup(
        timeout=0.2
    )

    readonly_popup_handled = dismiss_outlook_readonly_popup(
        timeout=0.2
    )

    return {
        "onboardingHandled": onboarding_handled,
        "licensePopupHandled": license_popup_handled,
        "readonlyPopupHandled": readonly_popup_handled,
    }


def dismiss_outlook_readonly_popup(timeout=5):
    """
    Handles the Office dialog "Möchten Sie wirklich nicht anmelden?"
    by clicking "Nur zur Anzeige verwenden".
    """
    desktop = Desktop(backend="uia")
    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                window = desktop.window(handle=wrapper.handle)
                readonly_button = window.child_window(
                    title="Nur zur Anzeige verwenden",
                    control_type="Button"
                )

                if readonly_button.exists(timeout=0.2):
                    readonly_button.click_input()
                    time.sleep(0.5)
                    return True
            except Exception:
                continue

        time.sleep(0.2)
    return False


def open_new_email_window(timeout=10):
    global _compose_handle
    global _outlook_process_id

    if not _is_outlook_running():
        open_outlook(timeout)

    desktop = Desktop(backend="uia")
    main_window = _get_outlook_main_window(timeout)
    main_wrapper = main_window.wrapper_object()
    outlook_pid = _get_window_process_id(main_wrapper)

    if outlook_pid is None:
        raise RuntimeError(
            "Could not determine Outlook process ID."
        )
    _outlook_process_id = outlook_pid

    # Nur bereits existierende Fenster dieses Outlook-Prozesses merken.
    existing_handles = set()
    for wrapper in desktop.windows():
        try:
            if (
                _get_window_process_id(wrapper)
                == outlook_pid
            ):
                existing_handles.add(wrapper.handle)
        except Exception:
            continue

    main_window.set_focus()
    # Neues Compose-Fenster öffnen
    send_keys("^n")
    deadline = perf_counter() + float(timeout)
    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                if not wrapper.is_visible():
                    continue
                # Nur Fenster desselben Outlook-Prozesses

                if (
                    _get_window_process_id(wrapper)
                    != outlook_pid
                ):
                    continue
                # Nur neu entstandene Fenster

                if wrapper.handle in existing_handles:
                    continue
                compose_window = desktop.window(
                    handle=wrapper.handle
                )

                # Ein einziger stabiler Locator reicht zur Identifikation.
                # Die restlichen Controls werden später beim Befüllen
                # ohnehin separat validiert.
                subject_field = compose_window.child_window(
                    auto_id="4101",
                    control_type="Edit"
                )

                if not subject_field.exists(timeout=0.05):
                    continue

                _compose_handle = wrapper.handle
                compose_window.wait(
                    "visible enabled",
                    timeout=2
                )
                return compose_window

            except Exception:
                continue

        time.sleep(0.05)

    raise RuntimeError(
        "New Outlook email window was not detected."
    )

def open_new_email(timeout=30):
    compose_window = open_new_email_window(timeout)
    return compose_window.window_text()

def inspect_open_email_controls(timeout=10):

    """
    Inspects the Outlook compose window.
    The compose window is resolved through the unique native handle
    captured when the window was opened.
    """
    window = _get_compose_window(timeout)

    return _inspect_window_controls(
        window,
        "OUTLOOK COMPOSE"
    )

def inspect_outlook_main_controls(timeout=10):

    """

    Inspects the currently detected Outlook main window.

    """
    window = _get_outlook_main_window(timeout)
    return _inspect_window_controls(
        window,
        "OUTLOOK MAIN"
    )

def inspect_outlook_startup_windows(
    timeout=15,
    poll_interval=0.05

):

    """
    Starts Outlook from a closed state and records every new visible
    top-level window that appears.
    This is a diagnostic keyword for discovering stable unique
    locators for Outlook startup/loading/main windows.
    """
    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is already running. "
            "Close Outlook before inspecting startup windows."
        )

    desktop = Desktop(backend="uia")
    existing_handles = {
        wrapper.handle
        for wrapper in desktop.windows()
    }

    observed_windows = []
    observed_handles = set()
    start = perf_counter()

    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    deadline = start + float(timeout)

    logger.console(

        "\n=== OUTLOOK STARTUP WINDOW INSPECTOR ==="
    )

    while perf_counter() < deadline:
        now = perf_counter()

        for wrapper in desktop.windows():
            try:
                if not wrapper.is_visible():
                    continue

                handle = wrapper.handle

                # Ignore windows that already existed before Outlook start.
                if handle in existing_handles:
                    continue

                if handle in observed_handles:
                    continue

                observed_handles.add(handle)
                metadata = _control_metadata(wrapper)
                metadata["firstSeenSeconds"] = round(
                    now - start,
                    3
                )
                observed_windows.append(metadata)
                logger.console(
                    "WINDOW | "
                    f"first_seen={metadata['firstSeenSeconds']}s | "
                    f"name={metadata['name']!r} | "
                    f"handle={metadata['handle']!r} | "
                    f"auto_id={metadata['automationId']!r} | "
                    f"control_type={metadata['controlType']!r} | "
                    f"class={metadata['className']!r} | "
                    f"pid={metadata['processId']!r} | "
                    f"runtime_id={metadata['runtimeId']!r}"

                )
            except Exception:
                continue
        time.sleep(float(poll_interval))

    logger.console(

        "=== END OUTLOOK STARTUP WINDOW INSPECTOR ===\n"
    )

    return observed_windows        

def wait_for_outlook_email(
    subject,
    send_started,
    timeout=90,
    poll_interval=5
):
    """Waits for an email with the given subject and returns delivery time."""
    main_window = _get_outlook_main_window(15)
    main_window.wait(
        "visible enabled ready",
        timeout=15
    )

    main_window.set_focus()
    send_keys("^+i")
    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        main_window.set_focus()
        send_keys("{F9}")
        time.sleep(float(poll_interval))

        # Re-acquire the main window by handle-based helper in case Outlook
        # refreshed or recreated its UI tree.
        main_window = _get_outlook_main_window(10)
        main_window.wait(
            "visible enabled",
            timeout=10
        )
        for control in main_window.descendants():
            try:
                text = control.window_text()
            except Exception:
                continue

            if text and subject in text:
                received_at = perf_counter()
                return {
                    "received": True,
                    "deliveryDurationSeconds": round(
                        received_at - float(send_started),
                        3
                    )
                }

    raise RuntimeError(
        f"Email was not received within {timeout} seconds. "
        f"Subject: {subject}"
    )    


def _get_compose_window(timeout=10):
    global _compose_handle
    if _compose_handle is None:
        raise RuntimeError(

            "No Outlook compose window has been opened by this RPA run."
        )
    desktop = Desktop(backend="uia")

    window = desktop.window(
        handle=_compose_handle
    )

    window.wait(
        "visible enabled",
        timeout=float(timeout)
    )
    return window


def fill_outlook_email(
    recipient,
    subject,
    body,
    cc=None,
    timeout=10

):

    window = _get_compose_window(timeout)

    to_field = window.child_window(
        auto_id="4117",
        control_type="Edit"
    )

    cc_field = window.child_window(
        auto_id="4126",
        control_type="Edit"
    )

    subject_field = window.child_window(
        auto_id="4101",
        control_type="Edit"
    )

    body_field = window.child_window(
        control_type="Document",
        class_name="_WwG"
    )

    to_field.wait("visible enabled", timeout=float(timeout))
    subject_field.wait("visible enabled", timeout=float(timeout))
    body_field.wait("visible enabled", timeout=float(timeout))
    to_field.set_edit_text(recipient)

    if cc:
        cc_field.wait("visible enabled", timeout=float(timeout))
        cc_field.set_edit_text(cc)

    subject_field.set_edit_text(subject)
    # Outlook Body ist kein normales Edit-Feld
    body_field.click_input()
    send_keys(body, with_spaces=True)

    return {
        "recipient": recipient,
        "cc": cc or "",
        "subject": subject,
        "windowTitle": window.window_text()
    }

def send_outlook_email(timeout=15):
    window = _get_compose_window(timeout)
    compose_wrapper = window.wrapper_object()
    compose_handle = compose_wrapper.handle

    send_button = window.child_window(
        auto_id="4256",
        control_type="Button"
    )
    send_button.wait(
        "visible enabled",
        timeout=float(timeout)
    )
    # Measure delivery time from the actual send action, not from the point
    # where the compose window finally disappears.
    send_started = perf_counter()
    known_com_error = None
    try:
        send_button.click_input()
    except COMError as exc:
        if exc.args and exc.args[0] == -2147220991:
            known_com_error = exc
        else:
            raise
    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        # Window was fully destroyed.
        if not win32gui.IsWindow(compose_handle):
            return {
                "sent": True,
                "sendStarted": send_started
            }
        # The native window may still exist internally although it is already
        # closed for the user.
        if not win32gui.IsWindowVisible(compose_handle):
            return {
                "sent": True,
                "sendStarted": send_started
            }
        time.sleep(0.2)

    if known_com_error:
        raise RuntimeError(
            "Outlook send action raised a UI Automation COM error "
            "and the compose window remained visible."
        )


    raise RuntimeError(
        "Outlook send command was executed, "
        "but the compose window remained visible."
    )


def dismiss_outlook_onboarding(timeout=5):
    """
    Detects the Office onboarding dialog and clicks
    'Vorerst überspringen'.
    Returns True if the dialog was handled,
    otherwise False.
    """
    desktop = Desktop(backend="uia")

    onboarding_window = desktop.window(
        title="Anmelden, um Office einzurichten"
    )

    if not onboarding_window.exists(timeout=float(timeout)):
        return False

    onboarding_window.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    skip_link = onboarding_window.child_window(
        title="Vorerst überspringen",
        control_type="Hyperlink",
        class_name="NetUIElement"
    )

    skip_link.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    skip_link.click_input()

    # Warten bis das störende Fenster wirklich verschwunden ist
    deadline = perf_counter() + float(timeout)
    while perf_counter() < deadline:
        if not onboarding_window.exists(timeout=0.2):
            return True
        time.sleep(0.2)
    raise RuntimeError(
        "Outlook onboarding dialog was detected, "
        "but could not be closed."
    )

def dismiss_outlook_license_popup(timeout=5):
    """
    Detects the Office license/trial popup and closes it.
    Returns True if the popup was found and closed,
    otherwise False.
    """
    desktop = Desktop(backend="uia")
    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                window = desktop.window(handle=wrapper.handle)
                buy_button = window.child_window(
                    title="Microsoft 365 kaufen",
                    control_type="Button"
                )

                if buy_button.exists(timeout=0.2):
                    window.set_focus()
                    # Schließt nur das zuvor eindeutig erkannte Popup
                    send_keys("%{F4}")
                    time.sleep(0.5)
                    return True
            except Exception:
                continue
        time.sleep(0.2)
    return False

def inspect_outlook_onboarding_controls(timeout=10):
    """
    Finds the Outlook/Microsoft onboarding dialog and prints
    UI Automation information for the relevant controls.
    """
    desktop = Desktop(backend="uia")
    deadline = perf_counter() + float(timeout)
    found = False
    logger.console("\n=== OUTLOOK ONBOARDING INSPECTOR ===")

    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                window = desktop.window(handle=wrapper.handle)
                for control in window.descendants():
                    try:
                        text = control.window_text() or ""
                        info = control.element_info
                    except Exception:
                        continue
                    if (
                        "Vorerst überspringen" in text
                        or "Melden Sie sich an" in text
                        or "Anmelden oder ein Konto erstellen" in text
                    ):
                        found = True

                        logger.console(
                            f"WINDOW={wrapper.window_text()!r} | "
                            f"name={text!r} | "
                            f"auto_id={info.automation_id!r} | "
                            f"control_type={info.control_type!r} | "
                            f"class={info.class_name!r}"
                        )
            except Exception:
                continue
        time.sleep(0.5)
    logger.console("=== END OUTLOOK ONBOARDING INSPECTOR ===\n")
    return found

def ensure_outlook_closed(timeout=10):
    """
    Ensures that Outlook is not running before a test starts.
    """
    if not _is_outlook_running():
        return True

    close_outlook(timeout)

    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is still running. "
            "Test environment could not be prepared."
        )

    return True

def close_outlook(timeout=10):
    global _main_handle
    global _outlook_process_id
    global _compose_handle

    result = _close_process(
        "OUTLOOK.EXE",
        timeout=timeout
    )

    _main_handle = None
    _outlook_process_id = None
    _compose_handle = None

    return result

def get_execution_environment():
    """
    Returns metadata about the machine executing the Outlook RPA test.
    """
    return _common_get_execution_environment(
        route_host="192.168.178.30",
        route_port=26500
    )    




