import re

_A1 = re.compile(r"^([A-Za-z]+)(\d+)$")


def col_to_idx(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def idx_to_col(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def a1_to_rc(a1: str) -> tuple[int, int]:
    m = _A1.match(a1.strip())
    if not m:
        raise ValueError(f"invalid A1: {a1!r}")
    return int(m.group(2)), col_to_idx(m.group(1))


def rc_to_a1(row: int, col: int) -> str:
    return f"{idx_to_col(col)}{row}"


def parse_range(a1: str) -> tuple[int, int, int, int]:
    start, end = (a1.split(":", 1) + [a1])[:2] if ":" in a1 else (a1, a1)
    r1, c1 = a1_to_rc(start)
    r2, c2 = a1_to_rc(end)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
