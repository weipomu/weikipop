"""Microbench: old `raw == raw` memcmp vs new sampled fingerprint used by
ScreenManager to detect screen changes between scan ticks.

Simulates a 4K BGRA buffer (3840 * 2160 * 4 = ~33 MB). Reports ms/op for:
  1. equal buffers — worst case for memcmp (must scan every byte)
  2. differ at end — also full scan before mismatch
  3. differ at start — best case for memcmp (early-exit at byte 0)

Run from the repo root:
    python scripts/microbench_screenshot_fingerprint.py
"""
import os
import sys
import time

# Make `src.*` importable when run from the repo root.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from src.utils.screen_fingerprint import content_fingerprint  # noqa: E402

W, H, BPP = 3840, 2160, 4
SIZE = W * H * BPP
N = 200

print(f"Buffer size: {SIZE / (1024 * 1024):.1f} MiB ({SIZE} bytes)")
print(f"Iterations:  {N}\n")

base = os.urandom(SIZE)
# Force a real copy — CPython's bytes(bytes_obj) returns the same object,
# which would short-circuit `==` via identity check and skew results.
equal = bytes(bytearray(base))
assert equal is not base and equal == base
diff_end = bytearray(base); diff_end[-1] ^= 0xFF; diff_end = bytes(diff_end)
diff_start = bytearray(base); diff_start[0] ^= 0xFF; diff_start = bytes(diff_start)

cases = [
    ("equal",      base, equal),
    ("differ end", base, diff_end),
    ("differ 0",   base, diff_start),
]


def bench(label, fn):
    t0 = time.perf_counter()
    for _ in range(N):
        fn()
    elapsed = time.perf_counter() - t0
    per_op_ms = (elapsed / N) * 1000
    print(f"  {label:20s} {per_op_ms:8.3f} ms/op    total {elapsed * 1000:8.1f} ms")


for name, a, b in cases:
    print(f"--- {name} ---")
    bench("old raw==raw",     lambda a=a, b=b: a == b)
    bench("new fingerprint",  lambda a=a, b=b: content_fingerprint(a) == content_fingerprint(b))
    print()
