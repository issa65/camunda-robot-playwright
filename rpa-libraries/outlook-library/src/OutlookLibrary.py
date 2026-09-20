from time import perf_counter
import time
import win32gui
from comtypes import COMError
from datetime import datetime, timezone
from pywinauto.keyboard import send_keys
from robot.api import logger
from uuid import uuid4
import subprocess

from pywinauto import Desktop


ROBOT_LIBRARY_SCOPE = "GLOBAL"
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


def open_outlook(timeout=30):
    """
    Opens Outlook and waits until an Outlook window is visible.
    """

    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    window = Desktop(backend="uia").window(
        title_re=r".*Outlook.*"
    )

    window.wait(
        "visible",
        timeout=float(timeout)
    )

    return window.window_text()


def measure_outlook_startup(timeout=30):
    """
    Measures a real Outlook cold start.

    The measurement is aborted if Outlook is already running.
    """

    if _is_outlook_running():
        raise RuntimeError(
            "Outlook is already running. "
            "Cold startup measurement aborted."
        )

    measured_at = datetime.now(timezone.utc).isoformat()

    start = perf_counter()

    subprocess.Popen(
        'start "" outlook',
        shell=True
    )

    window = Desktop(backend="uia").window(
        title_re=r".*Outlook.*"
    )

    window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    duration = perf_counter() - start

    return {
        "durationSeconds": round(duration, 3),
        "windowTitle": window.window_text(),
        "startupType": "COLD",
        "measuredAt": measured_at
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


def open_new_email_window(timeout=30):
    global _compose_handle

    if not _is_outlook_running():
        open_outlook(timeout)

    desktop = Desktop(backend="uia")

    main_window = desktop.window(
        title_re=r".*Outlook.*"
    )

    main_window.wait(
        "visible enabled ready",
        timeout=float(timeout)
    )

    existing_handles = {
        window.handle
        for window in desktop.windows()
    }

    main_window.set_focus()
    send_keys("^n")

    deadline = perf_counter() + float(timeout)

    while perf_counter() < deadline:
        for wrapper in desktop.windows():

            if (
                wrapper.handle not in existing_handles
                and wrapper.is_visible()
            ):
                _compose_handle = wrapper.handle

                compose_window = desktop.window(
                    handle=_compose_handle
                )

                compose_window.wait(
                    "visible enabled",
                    timeout=5
                )

                return compose_window

        time.sleep(0.25)

    raise RuntimeError(
        "New Outlook email window was not detected."
    )


def open_new_email(timeout=30):
    compose_window = open_new_email_window(timeout)

    return compose_window.window_text()

def inspect_open_email_controls(timeout=10):
    """
    Prints relevant controls of the currently open
    Outlook compose window to the Robot console.
    """

    desktop = Desktop(backend="uia")

    window = desktop.window(
        title_re=r".*Nachricht.*"
    )

    window.wait(
        "visible enabled",
        timeout=float(timeout)
    )

    logger.console("\n=== OUTLOOK COMPOSE CONTROLS ===")

    for control in window.descendants():
        info = control.element_info

        if info.control_type in {
            "Edit",
            "Button",
            "ComboBox",
            "Text",
            "Document"
        }:
            logger.console(
                f"type={info.control_type!r} | "
                f"name={control.window_text()!r} | "
                f"auto_id={info.automation_id!r} | "
                f"class={info.class_name!r}"
            )

    logger.console("=== END CONTROLS ===\n")

    return window.window_text()

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


def send_outlook_email(timeout=10):
    """
    Sends the currently open Outlook email.
    Returns send status and send timestamp.
    """

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
        if not win32gui.IsWindow(compose_handle):
            return {
                "sent": True,
                "sendStarted": send_started
            }

        time.sleep(0.2)

    if known_com_error:
        raise RuntimeError(
            "Outlook send action raised a UI Automation COM error "
            "and the compose window remained open."
        )

    raise RuntimeError(
        "Outlook send command was executed, "
        "but the compose window did not close."
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
    """
    Waits for an email with the given subject and
    returns delivery duration in seconds.
    """

    desktop = Desktop(backend="uia")

    main_window = desktop.window(
        title_re=r".*Outlook.*"
    )

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

        main_window = desktop.window(
            title_re=r".*Outlook.*"
        )

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

def prepare_outlook_ui(timeout=5):
    """
    Handles known Outlook/Office startup dialogs.
    """

    onboarding_handled = dismiss_outlook_onboarding(timeout)

    # Das zweite Fenster erscheint eventuell erst nach dem ersten
    time.sleep(1)

    license_popup_handled = dismiss_outlook_license_popup(timeout)

    return {
        "onboardingHandled": onboarding_handled,
        "licensePopupHandled": license_popup_handled
    }       


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