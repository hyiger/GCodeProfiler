import re

def _ini_value_to_float(v: str):
    """Best-effort parse of numeric-ish values from config.ini.

    Handles:
      - plain floats/ints ("35", "0.2")
      - percentages ("20%" -> 20.0)
      - nil/none ("nil", "none" -> None)
      - quoted strings

    Returns float or None.
    """
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s.lower() in ("nil", "none"):
        return None
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1].strip()
    if s.endswith('%'):
        try:
            return float(s[:-1].strip())
        except Exception:
            return None
    try:
        return float(s)
    except Exception:
        return None


def parse_config_ini(path: str) -> dict:
    """Parse PrusaSlicer-style `key = value` config.ini into a dict (raw strings).

    Lines starting with `#` are treated as comments.
    """
    out: dict[str, str] = {}
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            line = line.rstrip("\n")
            if not line or line.lstrip().startswith('#'):
                continue
            m = re.match(r'^([^=]+?)\s*=\s*(.*)$', line)
            if not m:
                continue
            k = m.group(1).strip()
            v = m.group(2).strip()
            out[k] = v
    return out


def config_get_float(cfg: dict, key: str):
    return _ini_value_to_float(cfg.get(key))

