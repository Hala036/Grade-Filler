"""Start the Grade Filler app:  python main.py

`--self-test` checks a packaged build without opening the app (see
gradefiller/selftest.py).
"""
import sys

if __name__ == "__main__":
    if "--self-test" in sys.argv:
        from gradefiller.selftest import run

        sys.exit(run())

    import flet as ft

    from gradefiller.app import main

    ft.run(main)
