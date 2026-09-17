import subprocess

from pywinauto import Desktop


ROBOT_LIBRARY_SCOPE = "GLOBAL"


def open_outlook(timeout=30):
    """
    Opens Outlook and waits until an Outlook window is visible.
    Returns the detected window title.
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