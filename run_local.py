from __future__ import annotations

from pathlib import Path

from bootstrap import ensure_dependencies

def main() -> None:
    ensure_dependencies(Path(__file__).resolve().parent / "requirements.txt")
    from app import create_app

    app = create_app()
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
