"""Read-only: dump verbatim makepy signatures for the OneNote COM methods we care about.

Does NOT instantiate OneNote (no mod.Application()) — EnsureModule only reads the registered
type library and generates the early-bound wrapper, so this is safe even in SSH session 0 and
never touches the running OneNote COM server. For Step-1 of the UpdateHierarchy concurrency
fact-finding: we want the EXACT parameter list (names, order, defaults) straight from the
generated module, not inferred.
"""

from __future__ import annotations

import inspect
import sys

from win32com.client import gencache, selecttlb

tlbs = [t for t in selecttlb.EnumTlbs() if "onenote" in (t.desc or "").lower()]
if not tlbs:
    print("NO ONENOTE TLB REGISTERED")
    sys.exit(1)

tlb = max(tlbs, key=lambda t: (int(t.major, 16), int(t.minor, 16)))
print(f"TLB: {tlb.desc!r}  clsid={tlb.clsid}  ver={tlb.major}.{tlb.minor}")

mod = gencache.EnsureModule(tlb.clsid, 0, int(tlb.major, 16), int(tlb.minor, 16))
print(f"MODULE FILE: {mod.__file__}")

WANT = [
    "UpdateHierarchy",
    "UpdatePageContent",
    "DeleteHierarchy",
    "DeletePageContent",
    "GetHierarchy",
    "GetPageContent",
    "CreateNewPage",
    "OpenHierarchy",
]

for cn in dir(mod):
    cls = getattr(mod, cn)
    if isinstance(cls, type) and hasattr(cls, "UpdateHierarchy"):
        print(f"\nINTERFACE CLASS: {cn}\n" + "=" * 60)
        for n in WANT:
            if not hasattr(cls, n):
                print(f"\n---- {n} ----  (NOT on this class)")
                continue
            print(f"\n---- {n} ----")
            try:
                print(inspect.getsource(getattr(cls, n)).rstrip())
            except Exception as exc:  # noqa: BLE001
                print(f"  (no source: {exc})")
        break
else:
    print("No interface class with UpdateHierarchy found in the module.")
