from time import perf_counter
from datetime import datetime, timezone
import subprocess

from pywinauto import Desktop


ROBOT_LIBRARY_SCOPE = "GLOBAL"


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