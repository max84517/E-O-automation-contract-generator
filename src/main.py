"""Entry point for E&O Settlement Contract Generator.

Run from the project root:
    python src/main.py
    # or, using Poetry:
    poetry run eo-generator
"""
import sys
from pathlib import Path

# Ensure the project root is on sys.path so that the 'src' package is
# importable regardless of the current working directory.
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from src.ui.app import App  # noqa: E402 — must follow sys.path setup


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
