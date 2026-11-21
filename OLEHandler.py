import argparse
from pathlib import Path
from OLELib.OLEUI import run_gtui
from OLEVersioning import __version__, __license__, __copyright__

def str2bool(value) -> bool:
    # argparse's type=bool treats any non-empty string (including "False") as True.
    if isinstance(value, bool):
        return value
    if value.lower() in ("true", "t", "yes", "y", "1"):
        return True
    if value.lower() in ("false", "f", "no", "n", "0"):
        return False
    raise argparse.ArgumentTypeError(f"Expected True/False, got {value!r}")


# ======================================================================================================================
# CMD Argument Parser
parser = argparse.ArgumentParser(
    description="The OLEHandler is an application that can extract files convert oleObject.bin files into PDFs from an "
                "Office file.",
    epilog="This project is currently in progress. Bugs should be reported to the GitHub Issues page."
)

# Argument Addon
parser.add_argument(
    "-x", "--extract", "-X",
    help="Extract the Office file that contains attachments.",
    dest="extract",
    type=Path or str
)
parser.add_argument(
    "-s", "--save", "-S",
    help="The location to save the attachments.",
    dest="saved_path",
    type=Path or str
)
parser.add_argument(
    "-d", "--debug",
    help="Whether to start the script in DEBUG mode.",
    dest="debug",
    type=str2bool,
    default=False
)
parser.add_argument(
    "-ui", "--ui",
    help="Run in GUI Mode. To run in CLI mode, UI flag should be set to False.",
    dest="ui",
    type=str2bool,
    default=True  # By default, running the EXEC starts GUI
)
parser.add_argument(
    "--version", "-ver", "-v",
    help="Show the current version of OLEHandler.",
    action="version",
    version=f"OLEHandler: {__version__}\nLicensed under the {__license__}.\n\n{__copyright__}"
)

# Parse Args
cmd_args = parser.parse_args()

# Register into Core Variables
DEBUG = cmd_args.debug
FILE_PATH = cmd_args.extract
SAVED_PATH = cmd_args.saved_path
UI = cmd_args.ui

if DEBUG is True:
    print(f"[CMD ARGS]: DEBUG: {DEBUG} | {FILE_PATH} | {SAVED_PATH} | {UI}")

# ======================================================================================================================
# Run Extraction
run_gtui(FILE_PATH, SAVED_PATH, UI, DEBUG)

