from pathlib import Path
import os
import re
import shutil
import tempfile
from zipfile import ZipFile, BadZipFile
from OLELib.OLEErrorUI import OLEErrorWindow
from OLELib.OLEExtInputDialog import OLEExtInputWindow, parse_attachment
from OLELib.OLEInfoUI import OLEInfoWindow
from OLELib.FileByteInfo import BYTE_START_LIST, BYTE_END_LIST, CONVERSION_LIST

BASE_DIR = Path(__file__).resolve().parent.parent
TEMP_DIR = os.path.join(BASE_DIR, "TempFile")

# FileType Acceptance
FOLDER_NAME = ["ppt", "doc", "word", "xls"]


def convert_file(filepath: Path or str, saved_path: Path or str = "", debug=False, ui=False):
    """
    Convert the DOCX, PPTX, XLSX files into Zips and extracts them.
    :param ui: bool True/False
    :param filepath: Path or str
    :param debug: bool - True/False
    :param saved_path: Path or str
    :return:
    """
    # Init Check Filepath
    if os.path.exists(filepath) is False:
        print(f"[ERROR 0]: The specified file does not exist.")

        if ui is True:
            error_window = OLEErrorWindow(f"[ERROR 0]: The specified file does not exist.")
            error_window.show()

        return

    # Extract into a fresh, per-run temp directory so leftovers from a previous run
    # (e.g. an old ppt/ folder) can never be picked up again.
    try:
        with tempfile.TemporaryDirectory(prefix="OLEHandler_") as temp_dir:
            with ZipFile(filepath, "r") as zip_ext:
                zip_ext.extractall(temp_dir)

            extract_attachments(saved_path, debug, ui, temp_dir=temp_dir)

    except BadZipFile as e:
        print(f"[ERROR 1]: File Extraction error. It is likely that the file is not a ZIP file.\n\n{e}")

        if ui is True:
            error_window = OLEErrorWindow(f"[ERROR 1]: File Extraction error. It is likely that the file "
                                          f"\nis not a ZIP file.\n\n{e}")
            error_window.show()


def _natural_key(name: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def _locate(data: bytearray, byte_start: bytes, byte_end: bytes):
    """
    Return (start, end) slice bounds of the embedded file, or (-1, -1) if not found.
    The end marker is searched with rfind so files containing the marker more than once
    (incrementally-updated PDFs with several %%EOF, JPEGs with EXIF thumbnails) aren't cut short.
    Matching is case-insensitive for text formats (e.g. <!DOCTYPE html>).
    """
    hay = bytes(data).lower() if byte_start.startswith(b"<") else data
    start = hay.find(byte_start)
    if start == -1:
        return -1, -1
    end = hay.rfind(byte_end, start)
    end = len(data) if end == -1 else end + len(byte_end)
    return start, end


def extract_attachments(saved_path: Path or str, debug, ui=False, temp_dir: str = TEMP_DIR):
    ATTACHMENT_COUNT = 0
    PASSED_COUNT = 0
    FAILED_COUNT = 0
    EXTRACTED = False

    file_dir = None
    # Handle Undefined Saved Dir
    # (the temp dir is deleted after the run, so never save into it)
    if not saved_path:
        saved_path = os.getcwd()

    for folder in FOLDER_NAME:
        if os.path.exists(os.path.join(temp_dir, folder)):
            file_dir = os.path.join(temp_dir, folder, "embeddings")
            break

    if file_dir is None:
        print("[ERROR 2] The Extracted file is NOT a 2007 onwards Excel/Word/Powerpoint File. (DOC, PPT and XLS are not"
              " accepted as they are from 97-2003).")

        if ui is True:
            error_window = OLEErrorWindow(
                "[ERROR 2] The Extracted file is NOT a 2007 onwards Excel/Word/Powerpoint File."
                "\n(DOC, PPT and XLS are not accepted as they are from 97-2003)."
            )
            error_window.show()

        return

    for (x, y, z) in os.walk(file_dir):
        if debug is True:
            print(f"{x} | {z}")

        z = sorted(z, key=_natural_key)
        ATTACHMENT_COUNT += len(z)

        for i in range(len(z)):
            # oleObject Saving Method
            if debug is True:
                print(z[i])

            if z[i].endswith("bin") and z[i].startswith("oleObject"):
                EXTRACTED = False

                if debug is True:
                    print(f"File [{z[i]}] is an oleObject.")

                # Name States
                init_sn = f"Attachment {i + 1}." + z[i].split(".")[1]
                p_init_sn = f"Attachment {i + 1}_Convert.oleconvert"
                print(f"{temp_dir} | {init_sn} | {p_init_sn}")
                temp_sn = os.path.join(temp_dir, init_sn)
                final_sn = os.path.join(saved_path, p_init_sn)
                # tbs = os.path.join(saved_path, z[i])

                shutil.copyfile(os.path.join(x, z[i]), temp_sn)

                with open(temp_sn, "rb") as file:
                    f_bin_data = file.read()
                    # print(f_bin_data)
                # ByteArray Conversion
                f_binarr_data = bytearray(f_bin_data)

                for byte_list_count in range(len(BYTE_START_LIST)):
                    byte_start = BYTE_START_LIST[byte_list_count]
                    byte_end = BYTE_END_LIST[byte_list_count]

                    bin_spt_1, bin_spt_2 = _locate(f_binarr_data, byte_start, byte_end)

                    if bin_spt_1 != -1:
                        # Keep exactly [start marker ... end marker] and drop the OLE wrapper around it.
                        f_binarr_data = f_binarr_data[bin_spt_1:bin_spt_2]
                        fin_bin_spt_2 = bin_spt_2

                        if debug is True:
                            print(f"BYTE LOCATION: {bin_spt_1} | {fin_bin_spt_2}")

                        with open(final_sn, "wb") as file:
                            file.write(f_binarr_data)

                        os.remove(temp_sn)
                        PASSED_COUNT += 1

                        fin_fn = f"Attachment {i + 1}.{CONVERSION_LIST[byte_list_count]}"
                        fin_fp = os.path.join(saved_path, fin_fn)

                        # Rename the file to the current filetype
                        try:
                            os.rename(final_sn, fin_fp)

                        except Exception as e:
                            print(e)

                        if debug is True:
                            print(f"Converted Attachment {i + 1} into {fin_fp}.")

                        EXTRACTED = True

                        break

                if EXTRACTED is False:
                    if ui is True:
                        ole_input = OLEExtInputWindow(
                            ui=ui,
                            debug=debug,
                            directory=saved_path,
                            file=init_sn,
                            data=f_binarr_data
                        )
                        ole_input.show()
                    else:
                        user_ext = input("OLEHandler failed to determine filetype, please input assumed filetype: ")

                        parse_attachment(ui, debug, saved_path, init_sn, user_ext, f_binarr_data)

                    os.remove(temp_sn)
                    EXTRACTED = True
                    PASSED_COUNT += 1

                if EXTRACTED is False:
                    if debug is True:
                        print(f"Failed to convert Attachment {i + 1}. Skipping...")

                    os.remove(temp_sn)
                    FAILED_COUNT += 1
                    continue

            # Non oleObject Saving Method
            else:
                if debug is True:
                    print(f"File [{z[i]}] is not an oleObject.")

                # Name States
                init_sn = f"Attachment {i + 1}." + z[i].split(".")[1]
                tbs = os.path.join(saved_path, z[i])
                final_sn = os.path.join(saved_path, init_sn)

                if debug is True:
                    print(f"{saved_path} | {init_sn}")

                # User Dialog
                print(f"Saving {z[i]} as {init_sn} in {saved_path}...")

                # Save Handling
                shutil.copyfile(os.path.join(x, z[i]), tbs)
                os.rename(tbs, final_sn)

                PASSED_COUNT += 1

    print(f"[FINAL RESULT]: TOTAL -> {ATTACHMENT_COUNT} | PASSED -> {PASSED_COUNT} | FAILED -> {FAILED_COUNT}")

    if ui is True:
        ole_info = OLEInfoWindow(
            title="Extraction Status",
            text=f"Total Count: {ATTACHMENT_COUNT}\nPassed: {PASSED_COUNT}\nFailed: {FAILED_COUNT}"
        )
        ole_info.show()
