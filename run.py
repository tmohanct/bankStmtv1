from __future__ import annotations

import os
import runpy
import subprocess
import sys
from pathlib import Path

REEXEC_GUARD_ENV = "BANKSTMT_SKIP_VENV_REEXEC"


def _read_pyvenv_cfg(cfg_path: Path) -> dict[str, str]:
    if not cfg_path.is_file():
        return {}

    config: dict[str, str] = {}
    for raw_line in cfg_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        if "=" not in raw_line:
            continue
        key, value = raw_line.split("=", 1)
        config[key.strip().lower()] = value.strip()
    return config


def _get_stale_venv_reason(venv_dir: Path) -> str | None:
    config = _read_pyvenv_cfg(venv_dir / "pyvenv.cfg")

    executable = config.get("executable")
    if executable and not Path(executable).exists():
        return f"base interpreter is missing: {executable}"

    home = config.get("home")
    if home and not Path(home).exists():
        return f"base Python home is missing: {home}"

    return None


def _maybe_reexec_into_venv(project_root: Path) -> None:
    if os.environ.get(REEXEC_GUARD_ENV) == "1":
        return

    venv_dir = project_root / ".venv"
    venv_python = venv_dir / "Scripts" / "python.exe"
    if not venv_python.is_file():
        return

    stale_reason = _get_stale_venv_reason(venv_dir)
    if stale_reason:
        print(
            f"Warning: ignoring stale virtual environment at {venv_dir} ({stale_reason}).",
            file=sys.stderr,
        )
        print(r"Run .\setup_windows.bat to recreate it on this machine.", file=sys.stderr)
        return

    current_python = Path(sys.executable).resolve()
    target_python = venv_python.resolve()
    if current_python == target_python:
        return

    env = os.environ.copy()
    env[REEXEC_GUARD_ENV] = "1"
    try:
        completed = subprocess.run(
            [str(target_python), str(project_root / "run.py"), *sys.argv[1:]],
            env=env,
            check=False,
        )
    except OSError as exc:
        print(
            f"Warning: unable to start project virtual environment at {venv_dir}: {exc}",
            file=sys.stderr,
        )
        print(r"Run .\setup_windows.bat to recreate it on this machine.", file=sys.stderr)
        return

    raise SystemExit(completed.returncode)


def main() -> None:
    project_root = Path(__file__).resolve().parent
    _maybe_reexec_into_venv(project_root)
    src_dir = project_root / "src"
    code_dir = project_root / "src" / "code"
    sys.path.insert(0, str(code_dir))
    sys.path.insert(1, str(src_dir))
    try:
        runpy.run_path(str(code_dir / "run.py"), run_name="__main__")
    except ModuleNotFoundError as exc:
        missing_module = getattr(exc, "name", None) or "a required dependency"
        print(f"Missing Python dependency: {missing_module}")
        print("Install or repair the project virtualenv first:")
        print(r"  .\install_fresh_machine.bat")
        print("If you prefer manual setup:")
        print(r"  py -3 -m venv .venv")
        print(r"  .\.venv\Scripts\python.exe -m pip install --upgrade pip")
        print(r"  .\.venv\Scripts\python.exe -m pip install -r requirements.txt")
        print("After that, `python run.py ...` will automatically switch into `.venv`.")
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
