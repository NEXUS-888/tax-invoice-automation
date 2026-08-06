import os
import sys
import traceback
import logging

# Ensure src directory is in Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

# Set up crash log file
CRASH_LOG = os.path.join(os.path.dirname(__file__), "crash_log.txt")

logging.basicConfig(
    filename=CRASH_LOG,
    level=logging.ERROR,
    format="%(asctime)s %(levelname)s: %(message)s",
)


def global_exception_handler(exc_type, exc_value, exc_tb):
    """Catches ANY unhandled exception and writes it to crash_log.txt."""
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    logging.error(f"UNHANDLED EXCEPTION:\n{error_msg}")
    # Also write to stderr so it shows in console
    print(f"\n{'='*60}", file=sys.stderr)
    print(f"CRASH CAUGHT — see {CRASH_LOG}", file=sys.stderr)
    print(f"{'='*60}", file=sys.stderr)
    print(error_msg, file=sys.stderr)


# Install global exception handler BEFORE importing anything else
sys.excepthook = global_exception_handler


if __name__ == "__main__":
    try:
        from app_gui import main
        main()
    except Exception:
        logging.error(f"STARTUP CRASH:\n{traceback.format_exc()}")
        print(f"\nCRASH — see {CRASH_LOG}", file=sys.stderr)
        print(traceback.format_exc(), file=sys.stderr)
        input("Press Enter to close...")
