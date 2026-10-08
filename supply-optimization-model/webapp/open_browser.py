"""
open_browser.py - start the propane model web app (if needed) and open it.

Run it any way you like:

    py webapp\\open_browser.py            start the app if it isn't running,
                                         wait until it answers, open the browser
    py webapp\\open_browser.py --wait     only wait for an app started elsewhere
                                         (this is what "Run Web App.bat" uses)
    py webapp\\open_browser.py --check    exit 0 if the app is running, 1 if not

When this script has to start the app itself, the app gets its OWN console
window titled "Propane Model Web App" - that window shows what the app is
doing, and closing it stops the app.

Opening the browser tries three ways in turn, because on locked-down PCs the
usual one sometimes does nothing without saying so:
Python's webbrowser module, then Windows' "open this URL" (os.startfile),
then explorer.exe with the URL. The address is always printed as well, so it
can be pasted into a browser by hand if all three are blocked.
"""
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser

PORT = int(os.environ.get("MODEL_WEB_PORT", "8050"))
URL = f"http://127.0.0.1:{PORT}/"
HERE = os.path.dirname(os.path.abspath(__file__))
WAIT_SECONDS = 120


def is_up() -> bool:
    """True when the propane web app (not just anything) answers on the port."""
    try:
        with urllib.request.urlopen(URL + "api/scenarios", timeout=2) as r:
            return r.status == 200
    except Exception:  # noqa: BLE001 - any failure just means "not (yet)"
        return False


def is_stale() -> bool:
    """
    True when the running app is older than the code on disk.

    New servers say so on /api/health. Servers from before that existed
    don't have it at all (404) - those are out of date by definition.
    """
    try:
        with urllib.request.urlopen(URL + "api/health", timeout=3) as r:
            import json
            return bool(json.loads(r.read().decode("utf-8")).get("stale"))
    except urllib.error.HTTPError as e:
        return e.code == 404
    except Exception:  # noqa: BLE001
        return False


def stop_running_app() -> bool:
    """Stop the out-of-date copy so a fresh one can start. True if it went away."""
    # Newer copies can be asked politely.
    try:
        req = urllib.request.Request(URL + "api/shutdown", data=b"{}", method="POST",
                                     headers={"X-Requested-With": "propane-model",
                                              "Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=3).close()
    except Exception:  # noqa: BLE001 - older copies don't have it; fall through
        pass
    for _ in range(20):
        if not port_taken():
            return True
        time.sleep(0.25)
    # Older copies: end the process that owns the port - but ONLY if it really
    # is this app (python running app.py), never anything else on the PC.
    if os.name == "nt":
        ps = ("$c=Get-NetTCPConnection -LocalPort %d -State Listen -ErrorAction SilentlyContinue;"
              "foreach($p in ($c.OwningProcess|Sort-Object -Unique)){"
              " $w=Get-CimInstance Win32_Process -Filter \"ProcessId=$p\";"
              " if($w.CommandLine -match 'app\\.py'){ Stop-Process -Id $p -Force; 'stopped' }"
              " else { 'not ours: ' + $w.CommandLine } }") % PORT
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=60)
        except Exception:  # noqa: BLE001
            pass
    for _ in range(40):
        if not port_taken():
            return True
        time.sleep(0.25)
    return False


def port_taken() -> bool:
    s = socket.socket()
    s.settimeout(1)
    try:
        return s.connect_ex(("127.0.0.1", PORT)) == 0
    finally:
        s.close()


def start_app() -> subprocess.Popen:
    """
    Start app.py in its own console window so its log stays visible.

    Python is launched directly (not through "cmd /k ...") on purpose: the
    project may live in a folder whose path contains spaces, and cmd.exe and Python
    quote paths with spaces differently, so wrapping it in cmd silently breaks
    the path. app.py names its own window.
    """
    flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
    return subprocess.Popen([sys.executable, os.path.join(HERE, "app.py")],
                            cwd=os.path.dirname(HERE), creationflags=flags)


def open_url() -> bool:
    """Open URL in the default browser. Returns True if some method accepted it."""
    try:
        if webbrowser.open(URL, new=2):
            return True
    except Exception:  # noqa: BLE001
        pass
    if os.name == "nt":
        try:
            os.startfile(URL)  # type: ignore[attr-defined]
            return True
        except Exception:  # noqa: BLE001
            pass
        try:
            subprocess.Popen(["explorer.exe", URL])
            return True
        except Exception:  # noqa: BLE001
            pass
    return False


def wait_until_up(proc=None) -> bool:
    deadline = time.time() + WAIT_SECONDS
    shown = False
    while time.time() < deadline:
        if is_up():
            return True
        if proc is not None and proc.poll() is not None:
            return False                       # the app window closed / crashed
        if not shown and time.time() > deadline - WAIT_SECONDS + 5:
            print("  still starting... (the first start can take a few seconds)")
            shown = True
        time.sleep(0.5)
    return False


def main() -> int:
    args = sys.argv[1:]
    if "--check" in args:
        return 0 if is_up() else 1

    print(f"Propane Model Web App  ->  {URL}")
    proc = None
    if is_up() and "--wait" not in args and is_stale():
        # A copy left running from before an update serves the new page on old
        # code - new buttons then fail. Swap it for a fresh one.
        print("  The running copy is out of date - restarting it with the latest version...")
        if not stop_running_app():
            print("\n  Couldn't stop the old copy. Close the 'Propane Model Web App' window")
            print("  (or end python.exe in Task Manager), then run this again.")
            return 1
    if is_up():
        print("  The app is already running.")
    elif "--wait" in args:
        print("  Waiting for the app to start...")
    else:
        if port_taken():
            print(f"\n  Something else is already using port {PORT}, and it isn't this app.")
            print(f"  Close it, or run with a different port, e.g.:  set MODEL_WEB_PORT=8051")
            return 1
        print("  Starting the app in a new window (close that window to stop the app)...")
        proc = start_app()

    if not is_up() and not wait_until_up(proc):
        print("\n  The app didn't start. Look at the 'Propane Model Web App' window for the error,")
        print("  or run it by hand to see it:   py webapp\\app.py")
        return 1

    if open_url():
        print("  Opened in your browser.")
    else:
        print("  Couldn't open a browser automatically.")
    print(f"  If no page appeared, paste this into your browser:  {URL}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
