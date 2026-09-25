from time import perf_counter
import time
import win32gui
from comtypes import COMError
from datetime import datetime, timezone
from pywinauto.keyboard import send_keys
from robot.api import logger
from uuid import uuid4
import subprocess
import getpass
import socket
from pywinauto import Desktop


ROBOT_LIBRARY_SCOPE = "GLOBAL"
OUTLOOK_LOADING_CLASS = "MsoSplash"
OUTLOOK_MAIN_CLASS = "rctrl_renwnd32"


def _get_window_class(wrapper):
    try:
        return wrapper.element_info.class_name or ""
    except Exception:
        return ""


def _get_window_control_type(wrapper):
    try:
        return wrapper.element_info.control_type or ""
    except Exception:
        return ""


def _find_outlook_loading_window(desktop):
    matches = []

    for wrapper in desktop.windows():
        try:
            if not wrapper.is_visible():
                continue

            if (
                _get_window_class(wrapper) == OUTLOOK_LOADING_CLASS
                and _get_window_control_type(wrapper) == "Window"
            ):
                matches.append(wrapper)

        except Exception:
            continue

    if len(matches) > 1:
        raise RuntimeError(
            f"Expected one Outlook loading window, found {len(matches)}."
        )

    return matches[0] if matches else None


def _find_outlook_main_window(desktop):
    matches = []

    for wrapper in desktop.windows():
        try:
            if not wrapper.is_visible():
                continue

            if (
                _get_window_class(wrapper) == OUTLOOK_MAIN_CLASS
                and _get_window_control_type(wrapper) == "Window"
            ):
                matches.append(wrapper)

        except Exception:
            continue

    if len(matches) > 1:
        raise RuntimeError(
            f"Expected one Outlook main window, found {len(matches)}."
        )

    return matches[0] if matches else None

_compose_handle = None

def _is_outlook_running():
    result = subprocess.run(

        [
            "tasklist",
            "/FI",
            "IMAGENAME eq OUTLOOK.EXE",
            "/FO",
            "CSV",
            "/NH"
        ],

        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        check=False

    )


    return b"OUTLOOK.EXE" in result.stdout.upper()


def _get_outlook_main_window(timeout=30):
    """Returns the visible Outlook main window using its native handle."""
    desktop = Desktop(backend="uia")
    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                title = wrapper.window_text() or ""

                if (
                    title.endswith(" - Outlook")
                    and wrapper.is_visible()
                ):
                    return desktop.window(handle=wrapper.handle)
            except Exception:
                continue

        time.sleep(0.25)

    raise RuntimeError(
        f"Outlook main window was not detected within {timeout} seconds."
    )


def open_outlook(timeout=30):
    """Opens Outlook and waits until the real main window is ready."""
    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    window = _get_outlook_main_window(timeout)
    window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    return window.window_text()


def measure_outlook_startup(timeout=30):
    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is already running. "
            "Cold startup measurement aborted."
        )

    desktop = Desktop(backend="uia")

    observed_windows = []
    observed_handles = set()

    loading_seen_at = None
    main_seen_at = None

    loading_process_id = None
    main_window = None

    measured_at = datetime.now(timezone.utc).isoformat()

    start = perf_counter()

    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    deadline = start + float(timeout)

    while perf_counter() < deadline:
        now = perf_counter()

        for wrapper in desktop.windows():
            try:
                if not wrapper.is_visible():
                    continue

                class_name = _get_window_class(wrapper)
                control_type = _get_window_control_type(wrapper)

                if control_type != "Window":
                    continue

                # Nur bekannte Outlook Startup-Fenster beachten
                if class_name not in {
                    OUTLOOK_LOADING_CLASS,
                    OUTLOOK_MAIN_CLASS
                }:
                    continue

                handle = wrapper.handle
                process_id = wrapper.element_info.process_id

                if handle not in observed_handles:
                    observed_handles.add(handle)

                    window_type = (
                        "LOADING"
                        if class_name == OUTLOOK_LOADING_CLASS
                        else "MAIN"
                    )

                    observed_windows.append({
                        "type": window_type,
                        "name": wrapper.window_text() or "",
                        "handle": handle,
                        "className": class_name,
                        "controlType": control_type,
                        "processId": process_id,
                        "firstSeenSeconds": round(
                            now - start,
                            3
                        )
                    })

                # Loading Window
                if (
                    class_name == OUTLOOK_LOADING_CLASS
                    and loading_seen_at is None
                ):
                    loading_seen_at = now
                    loading_process_id = process_id

                # Main Window
                if class_name == OUTLOOK_MAIN_CLASS:
                    # Wenn vorher Splash gesehen wurde,
                    # muss das Main Window zum selben Outlook-Prozess gehören.
                    if (
                        loading_process_id is not None
                        and process_id != loading_process_id
                    ):
                        continue

                    main_seen_at = now
                    main_window = desktop.window(
                        handle=handle
                    )
                    break

            except Exception:
                continue

        if main_window is not None:
            break

        time.sleep(0.05)

    if main_window is None:
        raise RuntimeError(
            f"Outlook main window was not detected "
            f"within {timeout} seconds."
        )

    main_window.wait(
        "visible enabled ready",
        timeout=float(timeout)
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

    return {
        "startupToLoadingSeconds": startup_to_loading,
        "loadingToMainSeconds": loading_to_main,
        "startupToMainSeconds": startup_to_main,
        "startupPath": startup_path,
        "windowTitle": main_window.window_text(),
        "startupType": "COLD",
        "measuredAt": measured_at,
        "observedWindows": observed_windows
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


def prepare_outlook_ui(timeout=5):
    """Handles the known Outlook/Office startup dialogs."""
    onboarding_handled = dismiss_outlook_onboarding(timeout)

    time.sleep(0.5)

    license_popup_handled = dismiss_outlook_license_popup(timeout)

    time.sleep(0.5)

    readonly_popup_handled = dismiss_outlook_readonly_popup(timeout)

    return {
        "onboardingHandled": onboarding_handled,
        "licensePopupHandled": license_popup_handled,
        "readonlyPopupHandled": readonly_popup_handled
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


def open_new_email_window(timeout=30):
    global _compose_handle

    if not _is_outlook_running():
        open_outlook(timeout)

    desktop = Desktop(backend="uia")
    main_window = _get_outlook_main_window(timeout)
    main_window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    existing_handles = {
        wrapper.handle
        for wrapper in desktop.windows()
    }

    main_window.set_focus()
    send_keys("^n")

    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        for wrapper in desktop.windows():
            try:
                if (
                    wrapper.handle in existing_handles
                    or not wrapper.is_visible()
                ):
                    continue

                compose_window = desktop.window(handle=wrapper.handle)

                # Do not accept an arbitrary new window. Verify the Outlook
                # compose window by stable control automation IDs.
                to_field = compose_window.child_window(
                    auto_id="4117",
                    control_type="Edit"
                )
                subject_field = compose_window.child_window(
                    auto_id="4101",
                    control_type="Edit"
                )
                send_button = compose_window.child_window(
                    auto_id="4256",
                    control_type="Button"
                )

                if not (
                    to_field.exists(timeout=0.2)
                    and subject_field.exists(timeout=0.2)
                    and send_button.exists(timeout=0.2)
                ):
                    continue

                _compose_handle = wrapper.handle
                compose_window.wait(
                    "visible enabled",
                    timeout=5
                )
                return compose_window
            except Exception:
                continue

        time.sleep(0.25)

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

def _control_metadata(control):
    """
    Returns UI Automation metadata for a window/control.
    Used only for inspection and locator discovery.
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
        rect = str(control.rectangle())
    except Exception:
        rect = ""

    handle = None

    try:
        handle = control.handle
    except Exception:
        if info is not None:
            handle = getattr(info, "handle", None)

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
        "rectangle": rect,
    }


def _inspect_window_controls(window, label):
    """
    Generic inspector for an already uniquely identified window.

    The window itself must already have been resolved by handle.
    """

    controls = []

    for control in window.descendants():
        try:
            metadata = _control_metadata(control)
            controls.append(metadata)
        except Exception:
            continue

    # Determine whether auto_id + control_type is unique
    locator_counts = {}

    for metadata in controls:
        auto_id = metadata["automationId"]
        control_type = metadata["controlType"]

        if not auto_id or not control_type:
            continue

        key = (auto_id, control_type)

        locator_counts[key] = locator_counts.get(key, 0) + 1

    logger.console(
        f"\n=== UI INSPECTOR: {label} ==="
    )

    window_metadata = _control_metadata(window.wrapper_object())

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
        auto_id = metadata["automationId"]
        control_type = metadata["controlType"]

        unique_auto_id = False

        if auto_id and control_type:
            unique_auto_id = (
                locator_counts.get(
                    (auto_id, control_type),
                    0
                ) == 1
            )

        logger.console(
            "CONTROL | "
            f"name={metadata['name']!r} | "
            f"auto_id={auto_id!r} | "
            f"control_type={control_type!r} | "
            f"class={metadata['className']!r} | "
            f"handle={metadata['handle']!r} | "
            f"pid={metadata['processId']!r} | "
            f"unique_auto_id_locator={unique_auto_id!r}"
        )

    logger.console(
        f"=== END UI INSPECTOR: {label} ===\n"
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

    return controls

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


def generate_test_id():

    """

    Generates a unique ID for an E2E email test.

    """

    return uuid4().hex[:8].upper()


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

    """

    Closes Outlook completely.


    Returns True if Outlook is closed.

    """


    if not _is_outlook_running():

        return True


    subprocess.run(

        ["taskkill", "/IM", "OUTLOOK.EXE", "/T", "/F"],

        stdout=subprocess.DEVNULL,

        stderr=subprocess.DEVNULL,

        check=False

    )


    deadline = perf_counter() + float(timeout)


    while perf_counter() < deadline:

        if not _is_outlook_running():

            return True


        time.sleep(0.25)


    raise RuntimeError(

        "Outlook could not be closed within the timeout."

    )


def get_execution_environment():

    username = getpass.getuser()

    hostname = socket.gethostname()


    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)


    try:

        # Keine echte Verbindung nötig.

        # Ermittelt die IP des Interfaces Richtung Camunda-Haupt-PC.

        sock.connect(("192.168.178.30", 26500))

        ip_address = sock.getsockname()[0]

    finally:

        sock.close()


    return {

        "username": username,

        "hostname": hostname,

        "ipAddress": ip_address

    }


