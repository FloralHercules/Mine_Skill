"""Windows dependencies; no Office process is attached or killed implicitly."""
import os
import sys
from pathlib import Path

_handles = []

def bootstrap():
    candidates = [Path(__file__).resolve().parents[1] / '.runtime']
    if os.environ.get('VISIO_SKILL_PYTHONPATH'):
        candidates = [Path(p) for p in os.environ['VISIO_SKILL_PYTHONPATH'].split(os.pathsep)] + candidates
    for base in candidates:
        if not base.is_dir():
            continue
        for path in [base, base/'win32', base/'win32/lib', base/'pywin32_system32']:
            if path.is_dir():
                sys.path.insert(0, str(path))
        dll = base/'pywin32_system32'
        if dll.is_dir() and os.name == 'nt':
            _handles.append(os.add_dll_directory(str(dll)))

def office():
    bootstrap()
    import win32com.client
    import pythoncom
    return win32com.client, pythoncom
