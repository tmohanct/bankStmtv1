from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path


SKIP_DIR_NAMES = {
    ".git",
    ".venv",
    ".pytest_cache",
    ".mypy_cache",
    "__pycache__",
    "dist",
    "tmp_cli",
    "tmp4kxszc32",
}
SKIP_FILE_SUFFIXES = {".pyc", ".pyo", ".log"}
ROOT_FILES_TO_INCLUDE = [
    "README.md",
    "SETUP_WINDOWS.md",
    "requirements.txt",
    "setup_windows.ps1",
    "setup_windows.bat",
    "run_bank_parser.bat",
    "install_new_machine.bat",
    "build_fresh_machine_package.py",
    "build_fresh_machine_package.bat",
    "build_windows_package.ps1",
]
ROOT_FILES_OPTIONAL = {
    "build_windows_package.ps1",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a clean Windows package for another machine")
    parser.add_argument(
        "--include-input-pdfs",
        action="store_true",
        help="Include PDF files from input/ in the package",
    )
    return parser.parse_args()


def should_skip_file(path: Path, include_input_pdfs: bool) -> bool:
    if path.suffix.lower() in SKIP_FILE_SUFFIXES:
        return True

    if path.name.startswith("~$"):
        return True

    parts = set(path.parts)
    if "output" in parts:
        return True

    if "src" in parts and "logs" in parts:
        return True

    if path.parent.name == "input" and path.suffix.lower() == ".pdf" and not include_input_pdfs:
        return True

    return False


def copy_tree(source_dir: Path, target_dir: Path, include_input_pdfs: bool) -> None:
    if not source_dir.exists():
        return

    for item in source_dir.rglob("*"):
        relative = item.relative_to(source_dir)
        if any(part in SKIP_DIR_NAMES for part in relative.parts):
            continue

        destination = target_dir / relative
        if item.is_dir():
            destination.mkdir(parents=True, exist_ok=True)
            continue

        if should_skip_file(item, include_input_pdfs):
            continue

        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(item, destination)


def copy_root_files(repo_root: Path, package_root: Path) -> None:
    for file_name in ROOT_FILES_TO_INCLUDE:
        source = repo_root / file_name
        if not source.exists():
            if file_name in ROOT_FILES_OPTIONAL:
                continue
            raise FileNotFoundError(f"Required file is missing: {source}")

        try:
            shutil.copy2(source, package_root / file_name)
        except PermissionError:
            if file_name in ROOT_FILES_OPTIONAL:
                continue
            raise


def ensure_placeholder_dirs(package_root: Path) -> None:
    for relative_dir in ("input", "output", "src/logs"):
        path = package_root / relative_dir
        path.mkdir(parents=True, exist_ok=True)
        keep_file = path / ".gitkeep"
        if not keep_file.exists():
            keep_file.write_text("", encoding="utf-8")


def create_package(repo_root: Path, include_input_pdfs: bool) -> Path:
    dist_dir = repo_root / "dist"
    dist_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
    package_name = f"bankStmtv1_fresh_windows_{timestamp}"
    zip_path = dist_dir / f"{package_name}.zip"

    staging_root = dist_dir / f"_{package_name}_staging"
    if staging_root.exists():
        shutil.rmtree(staging_root, ignore_errors=True)

    package_root = staging_root / package_name
    try:
        package_root.mkdir(parents=True, exist_ok=True)

        copy_root_files(repo_root, package_root)
        copy_tree(repo_root / "src", package_root / "src", include_input_pdfs)
        copy_tree(repo_root / "input", package_root / "input", include_input_pdfs)
        copy_tree(repo_root / "tests", package_root / "tests", include_input_pdfs)
        ensure_placeholder_dirs(package_root)

        archive_base = dist_dir / package_name
        shutil.make_archive(str(archive_base), "zip", staging_root, package_name)
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)

    return zip_path


def main() -> int:
    args = parse_args()
    repo_root = Path(__file__).resolve().parent
    zip_path = create_package(repo_root, include_input_pdfs=args.include_input_pdfs)
    print(f"Created fresh Windows package: {zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
