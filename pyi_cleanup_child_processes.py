# -*- coding: utf-8 -*-
"""PyInstaller runtime hook: clean child process trees before one-file temp cleanup.

The app launches a bundled Node.js process which in turn launches Chrome/Edge.
If the GUI is closed while those children are still alive, Windows keeps files in
PyInstaller's _MEI directory locked and the bootloader shows a cleanup warning.
"""

import atexit
import os
import subprocess
import threading


_OriginalPopen = subprocess.Popen
_tracked_processes = []
_lock = threading.Lock()


class _TrackedPopen(_OriginalPopen):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        with _lock:
            _tracked_processes.append(self)


subprocess.Popen = _TrackedPopen


def _kill_process_tree(proc):
    try:
        if proc.poll() is not None:
            return
    except Exception:
        return

    if os.name == 'nt':
        try:
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            killer = _OriginalPopen(
                ['taskkill', '/PID', str(proc.pid), '/T', '/F'],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                creationflags=flags,
            )
            killer.wait(timeout=5)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
    else:
        try:
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    try:
        proc.wait(timeout=3)
    except Exception:
        pass


@atexit.register
def _cleanup_children_before_pyinstaller_temp_removal():
    with _lock:
        processes = list(_tracked_processes)
    for proc in reversed(processes):
        _kill_process_tree(proc)
