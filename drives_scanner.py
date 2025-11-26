import psutil
import re
import warnings
import shutil
import pandas as pd
from pathlib import Path
import logging
import os
if os.name == "nt":
    import win32api
    import ctypes

"""
Drive scanning and analysis module.

It includes functions to:
- Detect external drives
- Filter drives and directories based on blacklists
- Scan directories and get certain properties
- Export drives (and directories) documentation as a pandas DataFrame or a nested dictionary

The main functionality is exposed through two key functions:
- scan_valid_drives_to_df(): Returns drive information as a pandas DataFrame
- scan_valid_drives_to_dict(): Returns drive information as a nested dictionary
"""


def get_dir_size(dir_path):
    """
    Computes total file size of a directory tree.
    Returns -1 if scanning is disabled or a PermissionError occurs.

    Args:
        dir_path (Path | str): Path object or string path of directory

    Returns:
        int: Total size in bytes, or -1 if scanning is disabled or could not be completed.
    """
    try:
        p = Path(dir_path)
        return sum(f.stat().st_size for f in p.rglob('*') if f.is_file())
    except PermissionError as e:
        warnings.warn(f"permission denied accessing {e.filename}, "
                      f"while calculating file sizes.\nFolder size could not be calculated.")
        return -1


def get_date_from_dir_name(dir_name: str) -> str | None:
    """
    Returns the date from the first 10 characters of the directory name, if it matches the YYYY-MM-DD format.
    Otherwise, returns None.

    Args:
        dir_name (str): directory name

    Returns:
        str: date in YYYY-MM-DD format or None
    """
    if re.fullmatch(r"\d{4}\W\d{2}\W\d{2}", dir_name[:10]):
        return dir_name[:10]
    else:
        return None


def bytes_to_gb(bytes_size: int) -> float:
    """Convert bytes to GB, rounded to 3 decimal places."""
    return round(bytes_size / (1024 ** 3), 3)


def is_external_drive(p: psutil._common.sdiskpart) -> bool:
    """
    Checks multiple (flawed) heuristics, if the drive is physical, external and no CD drive.
    On Windows DriveType fixed sometimes includes usb-connected external drives,
    therefore all fixed drives are considered external.

    Args:
        p (psutil._common.sdiskpart): disk partition object by psutil.disk_partitions()

    Returns:
        bool: whether the drive is considered external
    """
    if 'cdrom' in p.opts.lower():
        return False

    if not p.fstype:
        return False

    if os.name == "nt":
        if win32api.GetVolumeInformation(p.mountpoint)[0] == "":
            return False
        drive_type_code = ctypes.windll.kernel32.GetDriveTypeW(p.mountpoint)
        drive_removable = 2
        drive_fixed = 3
        # despite 3 means fixed, some removable USB-drives have DriveType 3 :(
        # thus all fixed drives are included
        return drive_type_code == drive_removable or drive_type_code == drive_fixed
    else:
        return p.mountpoint.startswith(("/media", "/mnt", "/Volumes"))


def get_external_drives() -> list[str]:
    """
    Returns a list of mountpoints paths of all drives considered external drives by is_external_drive().

    Returns:
        list[str]: A list of mountpoints of drives (e.g. ['C:\\', 'D:\\', 'F:\\']).
    """
    return [
        p.mountpoint
        for p in psutil.disk_partitions(all=False)
        if is_external_drive(p)
    ]


def get_drive_properties(mountpoint) -> dict:
    """
    Returns a dictionary containing properties of a drive.

    Args:
        mountpoint (str): The mountpoint of the drive.

    Returns:
        dict: A dictionary with drive properties.

    The returned dictionary contains the following keys:
        - "drive-name": The name of the drive.
        - "total-storage": The total storage size of the drive in GB.
        - "free-storage": The free storage size of the drive in GB.
    """
    storage_usage = shutil.disk_usage(mountpoint)
    if os.name == "nt":
        drive_name = win32api.GetVolumeInformation(mountpoint)[0]
    else:
        drive_name = Path(mountpoint).name
    logging.info(f"scanning {drive_name}...")
    return {
        "drive-name": drive_name,
        "total-storage": bytes_to_gb(storage_usage.total),
        "free-storage": bytes_to_gb(storage_usage.free),
    }


def is_valid_directory(subdir: Path, blacklist_directories: set[str]) -> bool:
    """
    Checks if a given Path is a valid directory, i.e. if it should be listed it in the drives' documentation.

    Parameters:
        subdir (Path): The path to check.
        blacklist_directories (set[str]): A set of directory names to blacklist.

    Returns:
        bool: True if the Path is a valid directory, False otherwise.
    """
    return (
            subdir.is_dir()
            and not subdir.name.startswith(".")
            and subdir.name not in blacklist_directories
    )


def scan_directories(parent_dir, blacklist_directories=()):
    """
    Scan directories (within a drive) and return their properties if they are considered valid by is_valid_directory.

    Args:
        parent_dir (str): directory path
        blacklist_directories (set): directories to blacklist

    Returns:
        list[dict]: A list of directories with their properties.
    """
    dirs = [
        {
            "project-name": subdir.name,
            "size": bytes_to_gb(get_dir_size(subdir)),
            "date": get_date_from_dir_name(subdir.name)
        }
        for subdir in Path(parent_dir).iterdir()
        if is_valid_directory(subdir, blacklist_directories)
    ]
    return dirs


def is_blacklisted_drive(mountpoint: str, blacklist_drives: set[str]):
    """
    Checks if a given drive is blacklisted.

    Parameters:
        mountpoint (str): The mountpoint of the drive to check.
        blacklist_drives (set[str]): A set of drive names to blacklist.

    Returns:
        bool: True if the drive is blacklisted, False otherwise.
    """
    if os.name == "nt":
        drive_name = win32api.GetVolumeInformation(mountpoint)[0]
        return drive_name in blacklist_drives
    else:
        return mountpoint in blacklist_drives


def scan_valid_drives_to_df(blacklist_drives=(), blacklist_directories=()) -> pd.DataFrame:
    """
    Scans all valid drives and their valid directories and returns a DataFrame with their properties, which is suited for Google Spreadsheet Documentation.
    A drive is considered valid if it is not blacklisted and considered external by is_external_drive().

    Args:
        blacklist_drives (set[str]): A set of drive names to blacklist.
        blacklist_directories (set[str]): A set of directory names to blacklist.

    Returns:
        pd.DataFrame: A DataFrame with the properties of the valid drives.
    """
    mountpoints = get_external_drives()
    directories = []
    for mountpoint in mountpoints:
        if is_blacklisted_drive(mountpoint, blacklist_drives):
            continue
        drive_properties = get_drive_properties(mountpoint)
        directories_properties = scan_directories(mountpoint, blacklist_directories)
        directories.extend([directory | drive_properties for directory in directories_properties])
    return pd.DataFrame(directories)


def scan_valid_drives_to_dict(blacklist_drives=(), blacklist_directories=()) -> dict:
    """
    Scan all external, non-blacklisted drives and return a nested dictionary
    containing their storage properties and scanned project directories.

    A drive is considered valid if:
      • it is detected as external by `get_external_drives()`
      • it is not in `blacklist_drives`

    Args:
        blacklist_drives (Iterable[str]): Drive names to exclude.
        blacklist_directories (Iterable[str]): Directory names to exclude
            when scanning each drive.

    Returns:
        dict: A nested dictionary with drive information and project lists.

            Structure:
            {
                "<drive-name>": {
                    "drive-name": str,
                    "total-storage": int,
                    "free-storage": int,
                    "projects": list[dict]
                }
            }
    """
    mountpoints = get_external_drives()
    drives = {}
    for mountpoint in mountpoints:
        if is_blacklisted_drive(mountpoint, blacklist_drives):
            continue
        drive_properties = get_drive_properties(mountpoint)
        drive_properties["projects"] = scan_directories(mountpoint, blacklist_directories)
        drives[drive_properties["drive-name"]] = drive_properties
    return drives

