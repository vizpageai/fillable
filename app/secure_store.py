from __future__ import annotations

import ctypes
from ctypes import wintypes
from pathlib import Path

CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2
CREDUI_MAX_USERNAME_LENGTH = 513


class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD)]


class CREDENTIALW(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.c_void_p),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.c_void_p),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


_advapi32 = ctypes.WinDLL("Advapi32.dll")
_crypt32 = ctypes.WinDLL("Crypt32.dll")
_kernel32 = ctypes.WinDLL("Kernel32.dll")

CredReadW = _advapi32.CredReadW
CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(ctypes.POINTER(CREDENTIALW))]
CredReadW.restype = wintypes.BOOL

CredWriteW = _advapi32.CredWriteW
CredWriteW.argtypes = [ctypes.POINTER(CREDENTIALW), wintypes.DWORD]
CredWriteW.restype = wintypes.BOOL

CredDeleteW = _advapi32.CredDeleteW
CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
CredDeleteW.restype = wintypes.BOOL

CredFree = _advapi32.CredFree
CredFree.argtypes = [ctypes.c_void_p]
CredFree.restype = None

CRYPTPROTECT_UI_FORBIDDEN = 0x1


class DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


CryptProtectData = _crypt32.CryptProtectData
CryptProtectData.argtypes = [
    ctypes.POINTER(DATA_BLOB),
    wintypes.LPCWSTR,
    ctypes.POINTER(DATA_BLOB),
    ctypes.c_void_p,
    ctypes.c_void_p,
    wintypes.DWORD,
    ctypes.POINTER(DATA_BLOB),
]
CryptProtectData.restype = wintypes.BOOL

CryptUnprotectData = _crypt32.CryptUnprotectData
CryptUnprotectData.argtypes = [
    ctypes.POINTER(DATA_BLOB),
    ctypes.POINTER(wintypes.LPWSTR),
    ctypes.POINTER(DATA_BLOB),
    ctypes.c_void_p,
    ctypes.c_void_p,
    wintypes.DWORD,
    ctypes.POINTER(DATA_BLOB),
]
CryptUnprotectData.restype = wintypes.BOOL

LocalFree = _kernel32.LocalFree
LocalFree.argtypes = [ctypes.c_void_p]
LocalFree.restype = ctypes.c_void_p


def _dpapi_secret_path(target_name: str) -> Path:
    appdata = Path.home() / "AppData" / "Roaming" / "Fillable" / "secrets"
    appdata.mkdir(parents=True, exist_ok=True)
    safe_name = "".join(ch if ch.isalnum() or ch in {"-", "_", "."} else "_" for ch in target_name)
    return appdata / f"{safe_name}.bin"


def _dpapi_protect(payload: bytes) -> bytes:
    data_in_buffer = ctypes.create_string_buffer(payload)
    data_in = DATA_BLOB(len(payload), ctypes.cast(data_in_buffer, ctypes.POINTER(ctypes.c_byte)))
    data_out = DATA_BLOB()
    if not CryptProtectData(
        ctypes.byref(data_in),
        None,
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(data_out),
    ):
        raise OSError("Failed to encrypt secret with Windows DPAPI.")
    try:
        return ctypes.string_at(data_out.pbData, data_out.cbData)
    finally:
        LocalFree(data_out.pbData)


def _dpapi_unprotect(payload: bytes) -> bytes | None:
    if not payload:
        return b""
    data_in_buffer = ctypes.create_string_buffer(payload)
    data_in = DATA_BLOB(len(payload), ctypes.cast(data_in_buffer, ctypes.POINTER(ctypes.c_byte)))
    data_out = DATA_BLOB()
    ok = CryptUnprotectData(
        ctypes.byref(data_in),
        None,
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(data_out),
    )
    if not ok:
        return None
    try:
        return ctypes.string_at(data_out.pbData, data_out.cbData)
    finally:
        LocalFree(data_out.pbData)


def set_secret(target_name: str, secret: str, username: str = "fillable") -> None:
    payload = (secret or "").encode("utf-16-le")
    blob = ctypes.create_string_buffer(payload)
    cred = CREDENTIALW()
    cred.Type = CRED_TYPE_GENERIC
    cred.TargetName = ctypes.c_wchar_p(target_name)
    cred.CredentialBlobSize = len(payload)
    cred.CredentialBlob = ctypes.cast(blob, ctypes.c_void_p)
    cred.Persist = CRED_PERSIST_LOCAL_MACHINE
    cred.UserName = ctypes.c_wchar_p(username[:CREDUI_MAX_USERNAME_LENGTH])
    if not CredWriteW(ctypes.byref(cred), 0):
        raise OSError("Failed to write Windows Credential Manager secret.")


def get_secret(target_name: str) -> str | None:
    ptr = ctypes.POINTER(CREDENTIALW)()
    ok = CredReadW(target_name, CRED_TYPE_GENERIC, 0, ctypes.byref(ptr))
    if not ok:
        return None
    try:
        cred = ptr.contents
        if cred.CredentialBlobSize == 0 or not cred.CredentialBlob:
            return ""
        data = ctypes.string_at(cred.CredentialBlob, cred.CredentialBlobSize)
        return data.decode("utf-16-le", errors="ignore")
    finally:
        CredFree(ptr)


def delete_secret(target_name: str) -> None:
    CredDeleteW(target_name, CRED_TYPE_GENERIC, 0)


def set_user_openai_key(target_name: str, secret: str) -> None:
    payload = (secret or "").encode("utf-8")
    encrypted = _dpapi_protect(payload)
    _dpapi_secret_path(target_name).write_bytes(encrypted)


def get_user_openai_key(target_name: str) -> str | None:
    path = _dpapi_secret_path(target_name)
    if path.exists():
        encrypted = path.read_bytes()
        decrypted = _dpapi_unprotect(encrypted)
        if decrypted is None:
            return None
        return decrypted.decode("utf-8", errors="ignore")
    # Fallback for older installs: migrate from Credential Manager.
    legacy = get_secret(target_name)
    if legacy is None:
        return None
    try:
        set_user_openai_key(target_name, legacy)
        delete_secret(target_name)
    except OSError:
        return legacy
    return legacy


def delete_user_openai_key(target_name: str) -> None:
    path = _dpapi_secret_path(target_name)
    if path.exists():
        path.unlink(missing_ok=True)
    delete_secret(target_name)
