"""Cheap change-detection fingerprint over a sampled subset of a raw screen
buffer. Used by ScreenManager to skip OCR when the screen hasn't changed,
without paying for a byte-wise compare of the full ~33 MB 4K buffer."""

# 8 x 4 KiB = 32 KiB sampled out of a full ~33 MB 4K buffer. Sampling spans
# the whole image (not just corners) so that text changes anywhere on screen
# are caught. False negatives just waste an OCR call; false positives would
# skip a needed lookup, so coverage matters more than chunk count.
_FP_CHUNKS = 8
_FP_CHUNK_SIZE = 4096


def content_fingerprint(raw: bytes) -> int:
    n = len(raw)
    if n <= _FP_CHUNKS * _FP_CHUNK_SIZE:
        return hash(raw)
    step = (n - _FP_CHUNK_SIZE) // (_FP_CHUNKS - 1)
    return hash(b"".join(
        raw[i * step : i * step + _FP_CHUNK_SIZE] for i in range(_FP_CHUNKS)
    ))
