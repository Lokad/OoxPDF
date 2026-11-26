"""Measure logical and stored file bytes, skipping links.

Bytes counts file lengths once per device/file ID. StoredBytes adjusts for
file compression and sparse storage. Filesystem metadata and allocation
slack are outside the measure. Separate copies count independently.
"""
import datetime
import json
import os
from pathlib import Path
import stat
import sys
import time

if os.name == "nt":
    from ctypes import POINTER, WinDLL, WinError, byref, get_last_error, set_last_error
    from ctypes.wintypes import DWORD, LPCWSTR

    _compressed_size = WinDLL("kernel32.dll", use_last_error=True).GetCompressedFileSizeW
    _compressed_size.argtypes = (LPCWSTR, POINTER(DWORD))
    _compressed_size.restype = DWORD


def stored_file_bytes(path, info):
    if os.name == "nt":
        if info.st_file_attributes & (stat.FILE_ATTRIBUTE_COMPRESSED | stat.FILE_ATTRIBUTE_SPARSE_FILE):
            high = DWORD()
            set_last_error(0)
            low = _compressed_size(path, byref(high))
            if low == 0xffffffff and get_last_error():
                raise WinError(get_last_error())
            return min(info.st_size, (high.value << 32) | low)
    elif stat.S_ISREG(info.st_mode) and hasattr(info, "st_blocks"):
        return min(info.st_size, info.st_blocks * 512)
    return info.st_size


def measure(path):
    root = str(Path(path).resolve())
    # scandir/stat need extended Windows paths even when .NET created the
    # directory successfully. Keep the ordinary resolved root in the report.
    walk_root = root
    if os.name == 'nt' and not root.startswith('\\\\?\\'):
        walk_root = '\\\\?\\UNC\\' + root[2:] if root.startswith('\\\\') else '\\\\?\\' + root
    if not os.path.isdir(walk_root):
        raise ValueError("Not a directory: " + root)
    started = time.monotonic()
    pending, seen, groups, errors = [walk_root], set(), {}, []
    total = stored = logical = files = directories = links = shared = 0
    while pending:
        directory = pending.pop()
        try:
            with os.scandir(directory) as entries:
                for entry in entries:
                    try:
                        # DirEntry.stat reports zero file IDs/link counts on Windows.
                        info = os.stat(entry.path, follow_symlinks=False)
                        if stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400):
                            links += 1
                            continue
                        if stat.S_ISDIR(info.st_mode):
                            directories += 1
                            pending.append(entry.path)
                            continue
                        files += 1
                        logical += info.st_size
                        charge = info.st_size
                        if info.st_nlink > 1 and info.st_ino:
                            identity = (info.st_dev, info.st_ino)
                            if identity in seen:
                                charge = 0
                                shared += info.st_size
                            else:
                                seen.add(identity)
                        total += charge
                        stored_charge = 0
                        if charge:
                            try:
                                stored_charge = stored_file_bytes(entry.path, info)
                            except OSError as error:
                                errors.append(entry.path + ": " + str(error))
                                stored_charge = charge
                        stored += stored_charge
                        relative = os.path.relpath(entry.path, walk_root)
                        group = relative.split(os.sep, 1)[0] if os.sep in relative else "."
                        item = groups.setdefault(group, dict(Path=group, Bytes=0, StoredBytes=0, Files=0))
                        item["Bytes"] += charge
                        item["StoredBytes"] += stored_charge
                        item["Files"] += 1
                    except OSError as error:
                        errors.append(entry.path + ": " + str(error))
        except OSError as error:
            errors.append(directory + ": " + str(error))
    for group in groups.values():
        group["GB"] = round(group["Bytes"] / 1e9, 3)
        group["StoredGB"] = round(group["StoredBytes"] / 1e9, 3)
    return dict(Root=root, CheckedUtc=datetime.datetime.now(datetime.timezone.utc).isoformat().replace("+00:00", "Z"),
                Bytes=total, GB=round(total / 1e9, 3), LogicalBytes=logical,
                StoredBytes=stored, StoredGB=round(stored / 1e9, 3),
                SharedHardLinkBytes=shared, Files=files, Directories=directories,
                SkippedLinks=links, Errors=errors, ElapsedSeconds=round(time.monotonic()-started, 2),
                Groups=sorted(groups.values(), key=lambda item: item["Bytes"], reverse=True))


if __name__ == "__main__":
    print(json.dumps(measure(sys.argv[1]), ensure_ascii=True))
