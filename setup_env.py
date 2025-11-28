from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
REQUIREMENTS = ROOT / "requirements.txt"


def ensure_python_version() -> None:
    if sys.version_info < (3, 8):
        raise SystemExit("Python 3.8+ is required.")


def install_requirements() -> None:
    if not REQUIREMENTS.exists():
        print(f"requirements.txt not found at {REQUIREMENTS}")
        return
    print(f"Installing dependencies from {REQUIREMENTS} into current Python environment...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)])


def main() -> None:
    ensure_python_version()
    install_requirements()
    print("Environment is ready.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit("Aborted by user.")
