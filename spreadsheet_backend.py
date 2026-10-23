"""Auto-select Microsoft Excel COM or LibreOffice UNO for spreadsheet files.

UNO runs in a separate LibreOffice-compatible Python interpreter. This module
does not copy or vendor LibreOffice DLLs; it reuses the native runtime already
installed with LibreOffice, avoiding incompatible CPython/pyuno builds.
"""
from __future__ import annotations

import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
from pathlib import Path

from excel_translate import EXCEL_EXTENSIONS, translate_excel_file, translate_rich_runs


def _excel_available():
    """Probe Excel.Application, not just whether pywin32 is installed."""
    if os.name != "nt":
        return False
    try:
        import win32com.client
        application = win32com.client.DispatchEx("Excel.Application")
        try:
            _ = application.Version
        finally:
            application.Quit()
        return True
    except Exception as exc:
        print(f"Excel desktop unavailable ({type(exc).__name__}: {exc})")
        return False


def _locate_libreoffice():
    configured = os.environ.get("LIBREOFFICE_HOME", "").strip().strip('"')
    candidates = []

    if configured:
        home = Path(configured)
        candidates += [home / "program" / "soffice.exe", home / "soffice.exe",
                       home / "program" / "soffice", home / "soffice"]

    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(variable)
        if root:
            candidates += [
                Path(root) / "LibreOffice" / "program" / "soffice.exe",
                Path(root) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
            ]

    for executable in ("soffice.exe", "soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            candidates.append(Path(found))

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    raise RuntimeError(
        "LibreOffice was not found. Install LibreOffice or set "
        "LIBREOFFICE_HOME=C:\\Program Files\\LibreOffice."
    )


def _select_uno_python(program_dir):
    env = os.environ.copy()
    env["PATH"] = str(program_dir) + os.pathsep + env.get("PATH", "")
    env["PYTHONPATH"] = str(program_dir) + os.pathsep + env.get("PYTHONPATH", "")

    candidates = [
        program_dir / "python.exe",
        program_dir / "python.bin",
        program_dir / "python",
        Path(sys.executable),
    ]
    # On Linux a system Python can have python3-uno while a virtualenv cannot.
    if os.name != "nt":
        candidates.append(Path("/usr/bin/python3"))

    for interpreter in candidates:
        if not interpreter.is_file():
            continue
        try:
            probe = subprocess.run(
                [str(interpreter), "-c", "import uno"],
                env=env, capture_output=True, text=True, timeout=10,
            )
            if probe.returncode == 0:
                return str(interpreter), env
        except (OSError, subprocess.TimeoutExpired):
            continue

    raise RuntimeError(
        "LibreOffice exists but no compatible Python-UNO interpreter was "
        "found. The agent loads pyuno and its DLL dependencies dynamically "
        "from LibreOffice's program directory. Ensure the LibreOffice Python "
        "component is installed, or use a Python build compatible with pyuno. "
        f"LibreOffice directory: {program_dir}"
    )


def _free_local_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _libreoffice_translate(input_path, output_path, translate_func, max_chars):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()

    if input_path.suffix.lower() in {".xlsm", ".xlsb"}:
        raise RuntimeError(
            "LibreOffice backend cannot safely round-trip .xlsm macros or "
            ".xlsb files. Convert a COPY to .xlsx with LibreOffice first; "
            "the original file will not be modified."
        )

    soffice = _locate_libreoffice()
    python_exe, environment = _select_uno_python(soffice.parent)
    worker_file = Path(__file__).resolve().with_name("libreoffice_uno_worker.py")
    if not worker_file.is_file():
        raise RuntimeError(f"UNO worker missing: {worker_file}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    port = _free_local_port()

    print(f"LibreOffice runtime : {soffice}")
    print(f"UNO Python          : {python_exe}")
    print(f"UNO input           : {input_path}")
    print(f"UNO output          : {output_path}")
    print("WARNING: LibreOffice may reserialize XLSX drawings, charts, macros "
          "and other advanced formatting. Inspect the output before use.")

    office = None
    worker = None
    with tempfile.TemporaryDirectory(prefix="jpvn_uno_") as temporary:
        profile_url = (Path(temporary) / "profile").resolve().as_uri()
        office_args = [
            str(soffice), "-env:UserInstallation=" + profile_url,
            "--headless", "--nologo", "--nodefault", "--norestore", "--nofirststartwizard",
            "--accept=socket,host=127.0.0.1,port="
            + str(port) + ";urp;StarOffice.ServiceManager",
        ]

        try:
            office = subprocess.Popen(
                office_args, env=environment, stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            worker = subprocess.Popen(
                [python_exe, "-u", str(worker_file), str(port),
                 str(input_path), str(output_path)],
                env=environment, text=True, encoding="utf-8", errors="replace",
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, bufsize=1,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            finished = False
            for line in worker.stdout:
                line = line.strip()
                if not line:
                    continue
                try:
                    message = json.loads(line)
                except json.JSONDecodeError:
                    print("  UNO worker: " + line[:500])
                    continue

                kind = message.get("type")
                if kind == "translate":
                    try:
                        runs = message["runs"]
                        result = translate_rich_runs(runs, translate_func, max_chars)
                        reply = {"ok": True, "runs": result}
                    except Exception as exc:
                        reply = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
                    worker.stdin.write(json.dumps(reply, ensure_ascii=False) + "\n")
                    worker.stdin.flush()
                elif kind == "log":
                    print("  UNO: " + str(message.get("text", "")))
                elif kind == "error":
                    raise RuntimeError("LibreOffice UNO failed:\n" + str(message["text"]))
                elif kind == "done":
                    print("LibreOffice translation completed: " + str(message.get("stats", {})))
                    finished = True
                else:
                    print("  UNO: " + line[:500])

            worker.wait(timeout=20)
            if worker.returncode != 0 or not finished:
                raise RuntimeError(
                    f"LibreOffice worker exited with code {worker.returncode}; "
                    f"finished={finished}. Check UNO messages above."
                )
        finally:
            if worker is not None and worker.poll() is None:
                worker.kill()
                worker.wait(timeout=10)
            if office is not None and office.poll() is None:
                office.terminate()
                try:
                    office.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    office.kill()
                    office.wait(timeout=10)


def translate_spreadsheet(input_path, output_path, translate_func, max_chars=1200):
    """Excel COM if installed and working; LibreOffice UNO otherwise."""
    if Path(input_path).suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(f"Unsupported spreadsheet: {input_path}")

    if _excel_available():
        try:
            print("Spreadsheet backend : Microsoft Excel COM")
            return translate_excel_file(
                input_path, output_path, translate_func, max_chars=max_chars
            )
        except Exception as exc:
            print(
                "WARNING: Excel COM failed, attempting LibreOffice UNO: "
                f"{type(exc).__name__}: {exc}"
            )

    print("Spreadsheet backend : LibreOffice UNO")
    return _libreoffice_translate(input_path, output_path, translate_func, max_chars)
