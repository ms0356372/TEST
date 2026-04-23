"""Entry point for desktop application."""
from app.gui.app import UltrasoundCheckinApp


def main() -> None:
    app = UltrasoundCheckinApp()
    app.mainloop()


if __name__ == "__main__":
    main()
