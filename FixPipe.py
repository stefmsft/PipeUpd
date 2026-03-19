"""FixPipe - Interactive Excel Pipe File Repair Tool."""

import configparser
import os
import re
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

import sys

from dotenv import load_dotenv

from templates.by_sales_charts import CHART1_XML, CHART2_XML, DRAWING_XML


def _s(text: str) -> str:
    """Return a console-safe string, replacing characters the terminal can't encode."""
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    return text.encode(enc, errors="replace").decode(enc)


# ---------------------------------------------------------------------------
# Environment loading via config.ini
# ---------------------------------------------------------------------------

def load_env_from_config() -> str:
    """Read config.ini for ENV_SUFFIX, load the matching .env file.

    Returns the suffix string (empty string = production).
    """
    suffix = ""
    config_path = os.path.join(os.path.dirname(__file__) or ".", "config.ini")

    if os.path.exists(config_path):
        cfg = configparser.ConfigParser()
        cfg.read(config_path, encoding="utf-8")
        suffix = cfg.get("Environment", "ENV_SUFFIX", fallback="").strip()

    env_file = f".env.{suffix}" if suffix else ".env"
    env_path = os.path.join(os.path.dirname(__file__) or ".", env_file)
    if not os.path.exists(env_path):
        print(f"WARNING: {env_file} not found, falling back to .env")
        env_file = ".env"
        env_path = os.path.join(os.path.dirname(__file__) or ".", ".env")
        suffix = ""

    load_dotenv(env_path, override=True)
    return suffix


# ---------------------------------------------------------------------------
# XML namespace / relationship-type constants
# ---------------------------------------------------------------------------

NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT   = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_WB   = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

RT_DRAWING     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
RT_CHART       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
RT_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
RT_PIVOT_CACHE         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
RT_PIVOT_CACHE_RECORDS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"

CT_CHART   = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"


# ---------------------------------------------------------------------------
# Session: in-memory ZIP working state
# ---------------------------------------------------------------------------

class Session:
    """In-memory representation of the XLSX ZIP archive.

    All fix functions read from and write to this object.  The file on disk is
    only written when the user chooses to exit (choice 0).

    Implements the same read interface as zipfile.ZipFile (read / namelist /
    infolist) so all helpers work with a Session unchanged.
    """

    def __init__(self, path: str) -> None:
        self._infos: dict[str, zipfile.ZipInfo] = {}
        self._data:  dict[str, bytes]           = {}
        self.applied_fixes: list[str]           = []

        with zipfile.ZipFile(path, "r") as zf:
            for info in zf.infolist():
                self._infos[info.filename] = info
                self._data[info.filename]  = zf.read(info.filename)

    # --- ZipFile-compatible read interface used by helpers ---

    def read(self, name: str) -> bytes:
        return self._data[name]

    def namelist(self) -> list[str]:
        return list(self._data.keys())

    def infolist(self) -> list[zipfile.ZipInfo]:
        return list(self._infos.values())

    # --- Mutation ---

    def write(self, name: str, data: bytes) -> None:
        """Add or replace a ZIP entry (marks session as dirty)."""
        self._data[name] = data
        # Keep the original ZipInfo if the file already exists; otherwise the
        # name alone is enough for writestr() at save time.

    # --- Persistence ---

    @property
    def dirty(self) -> bool:
        return bool(self.applied_fixes)

    def save(self, path: str) -> None:
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for name, data in self._data.items():
                info = self._infos.get(name)
                zf_out.writestr(info if info is not None else name, data)
        with open(path, "wb") as f:
            f.write(buf.getvalue())


# ---------------------------------------------------------------------------
# Shared OOXML helpers  (accept Session or any duck-typed ZipFile-like object)
# ---------------------------------------------------------------------------

def _find_sheet_file(zf, tab_name: str) -> str | None:
    """Return the sheetN.xml filename for the named tab, or None."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rid = None
    for sheet_el in wb_root.iter(f"{{{NS_WB}}}sheet"):
        if sheet_el.get("name") == tab_name:
            rid = sheet_el.get(f"{{{NS_R}}}id")
            break
    if rid is None:
        return None

    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.iter(f"{{{NS_RELS}}}Relationship"):
        if rel.get("Id") == rid:
            target = rel.get("Target", "")   # e.g. "worksheets/sheet4.xml"
            return target.split("/")[-1]      # "sheet4.xml"
    return None


def _resolve_zip_path(base_dir: str, relative_target: str) -> str:
    """Resolve a relative .rels target from a base directory to a ZIP path."""
    parts = base_dir.split("/")
    for seg in relative_target.split("/"):
        if seg == "..":
            parts.pop()
        elif seg and seg != ".":
            parts.append(seg)
    return "/".join(parts)


def _find_pivot_tables_on_sheet(zf, sheet_file: str) -> list[str]:
    """Return list of ZIP paths for all pivot tables attached to a sheet."""
    rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
    if rels_path not in zf.namelist():
        return []
    rels_root = ET.fromstring(zf.read(rels_path))
    results = []
    for rel in rels_root.iter(f"{{{NS_RELS}}}Relationship"):
        if rel.get("Type") == RT_PIVOT_TABLE:
            full = _resolve_zip_path("xl/worksheets", rel.get("Target", ""))
            results.append(full)
    return results


def _find_pivot_cache_path(zf, pt_path: str) -> str | None:
    """Return ZIP path to the cache definition for a pivot table, or None."""
    pt_filename = pt_path.split("/")[-1]
    rels_path = f"xl/pivotTables/_rels/{pt_filename}.rels"
    if rels_path not in zf.namelist():
        return None
    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.iter(f"{{{NS_RELS}}}Relationship"):
        if rel.get("Type") == RT_PIVOT_CACHE:
            return _resolve_zip_path("xl/pivotTables", rel.get("Target", ""))
    return None


def _find_field_in_cache(
    zf, cache_path: str, field_name: str,
) -> tuple[int, dict[int, str]] | tuple[None, None]:
    """Find a cacheField whose first line matches field_name (case-insensitive).

    Excel encodes newlines in field names as '_x000a_'.  We match only the text
    before the first newline, so "Quarter Invoice" matches the cache field named
    "Quarter Invoice_x000a_Facturation".

    Returns (field_idx, {item_index: value}) or (None, None) if not found.
    """
    cache_root = ET.fromstring(zf.read(cache_path))
    needle = field_name.strip().lower()
    for idx, cf in enumerate(cache_root.findall(f".//{{{NS_WB}}}cacheField")):
        raw_name = cf.get("name", "")
        first_line = raw_name.replace("_x000a_", "\n").split("\n")[0].strip().lower()
        if first_line == needle:
            shared = cf.find(f"{{{NS_WB}}}sharedItems")
            if shared is None:
                return None, None
            idx_to_val: dict[int, str] = {}
            for i, item in enumerate(list(shared)):
                tag = item.tag.split("}")[-1]
                idx_to_val[i] = item.get("v", "") if tag != "m" else ""
            return idx, idx_to_val
    return None, None


def _get_field_axis(pt_str: str, field_idx: int) -> str | None:
    """Return the axis attribute of pivotField[field_idx], or None.

    Possible values: "axisPage", "axisRow", "axisCol", "axisValues" (or None
    when the attribute is absent, which also means a values/data field).
    Only axisPage / axisRow / axisCol fields have filterable <items> blocks.
    """
    pf_starts = [m.start() for m in re.finditer(r"<pivotField(?![sS])\b", pt_str)]
    if len(pf_starts) <= field_idx:
        return None
    pf_start = pf_starts[field_idx]
    pf_end   = pf_starts[field_idx + 1] if field_idx + 1 < len(pf_starts) else len(pt_str)
    pf_chunk = pt_str[pf_start:pf_end]
    m = re.search(r'\baxis="([^"]*)"', pf_chunk)
    return m.group(1) if m else None


def _apply_filter_to_pivot(
    pt_str: str,
    field_idx: int,
    idx_to_val: dict[int, str],
    selected: set[str],
) -> str | None:
    """Rewrite the <items> block inside pivotField[field_idx].

    - Items whose value is in *selected*: h attribute removed (visible/checked)
    - All other items: h="1" (hidden/unchecked)
    - The <pageField> element is intentionally NOT modified; Excel derives the
      active filter state entirely from the h attrs on <item> elements.

    Returns the updated XML string, or None if no change is needed.
    """
    pf_starts = [m.start() for m in re.finditer(r"<pivotField(?![sS])\b", pt_str)]
    if len(pf_starts) <= field_idx:
        return None

    pf_start = pf_starts[field_idx]
    pf_end   = pf_starts[field_idx + 1] if field_idx + 1 < len(pf_starts) else len(pt_str)
    pf_chunk = pt_str[pf_start:pf_end]

    items_open  = pf_chunk.find("<items")
    items_close = pf_chunk.find("</items>")
    if items_open == -1 or items_close == -1:
        return None
    items_close += len("</items>")
    items_block = pf_chunk[items_open:items_close]

    def rewrite_item(m: re.Match) -> str:
        full = m.group(0)
        if 't="default"' in full:
            return full
        x_m = re.search(r'\bx="(\d+)"', full)
        if x_m is None:
            return full
        x = int(x_m.group(1))
        in_selected = idx_to_val.get(x, "") in selected

        attrs: dict[str, str] = {
            k: v for k, v in re.findall(r'(\w+)="([^"]*)"', full) if k != "h"
        }
        parts: list[str] = []
        if not in_selected:
            parts.append('h="1"')
        for k in ("sd", "n", "m", "x"):
            if k in attrs:
                parts.append(f'{k}="{attrs[k]}"')
        return f'<item {" ".join(parts)}/>'

    new_items_block = re.sub(r"<item\b[^/]*/>\s*", lambda m: rewrite_item(m), items_block)
    if new_items_block == items_block:
        return None

    new_pf_chunk = pf_chunk[:items_open] + new_items_block + pf_chunk[items_close:]
    return pt_str[:pf_start] + new_pf_chunk + pt_str[pf_end:]


# ---------------------------------------------------------------------------
# Fix 1: By Sales Charts
# ---------------------------------------------------------------------------

def _has_drawing_rel(zf, sheet_file: str) -> bool:
    rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
    if rels_path not in zf.namelist():
        return False
    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.iter(f"{{{NS_RELS}}}Relationship"):
        if rel.get("Type") == RT_DRAWING:
            return True
    return False


def _next_available_number(zf, pattern: str) -> int:
    existing = set()
    prefix, suffix = pattern.split("{}")
    for name in zf.namelist():
        if name.startswith(prefix) and name.endswith(suffix):
            mid = name[len(prefix):-len(suffix)] if suffix else name[len(prefix):]
            if mid.isdigit():
                existing.add(int(mid))
    return max(existing, default=0) + 1


def _next_free_rid(rels_xml_bytes: bytes) -> str:
    root = ET.fromstring(rels_xml_bytes)
    max_id = 0
    for rel in root.iter(f"{{{NS_RELS}}}Relationship"):
        rid = rel.get("Id", "")
        m = re.match(r"rId(\d+)", rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return f"rId{max_id + 1}"


def fix_by_sales_charts(session: Session) -> bool:
    """Restore the two missing pivot charts on the 'By Sales' tab."""
    sheet_file = _find_sheet_file(session, "By Sales")
    if sheet_file is None:
        print("  ERROR: 'By Sales' tab not found in workbook.")
        return False
    print(f"  Found 'By Sales' -> {sheet_file}")

    if _has_drawing_rel(session, sheet_file):
        print("  No fix needed: 'By Sales' already has a drawing relationship.")
        return False

    chart_n   = _next_available_number(session, "xl/charts/chart{}.xml")
    drawing_n = _next_available_number(session, "xl/drawings/drawing{}.xml")
    print(f"  Will create: chart{chart_n}.xml, chart{chart_n + 1}.xml, drawing{drawing_n}.xml")

    sheet_rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
    if sheet_rels_path in session.namelist():
        sheet_rels_bytes = session.read(sheet_rels_path)
    else:
        sheet_rels_bytes = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    drawing_rid = _next_free_rid(sheet_rels_bytes)
    new_rel = (
        f'<Relationship Id="{drawing_rid}" '
        f'Type="{RT_DRAWING}" '
        f'Target="../drawings/drawing{drawing_n}.xml"/>'
    )
    modified_sheet_rels = sheet_rels_bytes.decode("utf-8").replace(
        "</Relationships>", f"{new_rel}</Relationships>",
    )

    drawing_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_RELS}">'
        f'<Relationship Id="rId1" Type="{RT_CHART}" Target="../charts/chart{chart_n}.xml"/>'
        f'<Relationship Id="rId2" Type="{RT_CHART}" Target="../charts/chart{chart_n + 1}.xml"/>'
        '</Relationships>'
    )

    ws_path = f"xl/worksheets/{sheet_file}"
    ws_xml  = session.read(ws_path).decode("utf-8")
    drawing_tag = f'<drawing r:id="{drawing_rid}"/>'
    if "<drawing " not in ws_xml:
        ws_xml = ws_xml.replace("</worksheet>", f"{drawing_tag}</worksheet>")
        if "xmlns:r=" not in ws_xml:
            ws_xml = ws_xml.replace(
                f'xmlns="{NS_WB}"', f'xmlns="{NS_WB}" xmlns:r="{NS_R}"',
            )

    ct_str = session.read("[Content_Types].xml").decode("utf-8")
    new_overrides = (
        f'<Override PartName="/xl/charts/chart{chart_n}.xml" ContentType="{CT_CHART}"/>'
        f'<Override PartName="/xl/charts/chart{chart_n + 1}.xml" ContentType="{CT_CHART}"/>'
    )
    drawing_part = f"/xl/drawings/drawing{drawing_n}.xml"
    if drawing_part not in ct_str:
        new_overrides += f'<Override PartName="{drawing_part}" ContentType="{CT_DRAWING}"/>'
    modified_ct = ct_str.replace("</Types>", f"{new_overrides}</Types>")

    # Commit all changes to session
    session.write("[Content_Types].xml", modified_ct.encode("utf-8"))
    session.write(sheet_rels_path,       modified_sheet_rels.encode("utf-8"))
    session.write(ws_path,               ws_xml.encode("utf-8"))
    session.write(f"xl/charts/chart{chart_n}.xml",          CHART1_XML.strip().encode("utf-8"))
    session.write(f"xl/charts/chart{chart_n + 1}.xml",      CHART2_XML.strip().encode("utf-8"))
    session.write(f"xl/drawings/drawing{drawing_n}.xml",    DRAWING_XML.strip().encode("utf-8"))
    session.write(
        f"xl/drawings/_rels/drawing{drawing_n}.xml.rels",
        drawing_rels_xml.encode("utf-8"),
    )
    print(f"  Changes staged in session (not yet saved to disk).")
    return True


# ---------------------------------------------------------------------------
# Fix 2: Pivot page-filter selection
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Pivot cache surgery helpers
# ---------------------------------------------------------------------------

def _et_serialize_ns_safe(original_bytes: bytes, root) -> bytes:
    """Serialize an ET element while preserving namespace prefixes from the original.

    ET renames unknown namespace prefixes (mc: → ns1:, xr: → ns3:, …) which
    breaks markup-compatibility annotations like mc:Ignorable="xr": the value
    "xr" references the OLD prefix name, but ET has renamed it, so Excel rejects
    the file.  Registering every xmlns:prefix="uri" pair found in the original
    bytes before calling tostring() prevents that renaming.

    Additionally, ET drops namespace declarations for prefixes that are not
    actually used as element/attribute prefixes in the serialized tree.  For
    example, if the original had xmlns:xr="…" declared (needed so that
    mc:Ignorable="xr" resolves the "xr" prefix) but no element uses xr:*,
    ET would silently drop xmlns:xr — making mc:Ignorable="xr" invalid.
    We detect and re-inject any such "declared but unused" namespace declarations.
    """
    original_str = original_bytes.decode("utf-8", errors="replace")

    # Register default namespace: xmlns="uri"
    m = re.search(r'xmlns="([^"]*)"', original_str)
    if m:
        ET.register_namespace("", m.group(1))

    # Register all prefixed namespaces: xmlns:prefix="uri"
    original_ns: dict[str, str] = {}
    for prefix, uri in re.findall(r'xmlns:(\w+)="([^"]*)"', original_str):
        ET.register_namespace(prefix, uri)
        original_ns[prefix] = uri

    body = ET.tostring(root, encoding="unicode", xml_declaration=False)

    # Re-inject any namespace declarations that ET dropped because the prefix
    # was not actually used in any element/attribute name.  ET does this for
    # xmlns:xr when xr: is only referenced inside attribute *values* such as
    # mc:Ignorable="xr" — dropping xmlns:xr makes mc:Ignorable invalid XML.
    if original_ns:
        output_ns = set(re.findall(r'xmlns:(\w+)=', body))
        missing = {p: u for p, u in original_ns.items() if p not in output_ns}
        if missing:
            ns_decls = " ".join(f'xmlns:{p}="{u}"' for p, u in missing.items())
            first_gt = body.find(">")
            body = body[:first_gt] + " " + ns_decls + body[first_gt:]

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body).encode("utf-8")


def _disable_cache_auto_refresh(session: Session, cache_def_path: str) -> None:
    """Set refreshOnLoad="0" on the pivot cache definition if it was "1".

    When refreshOnLoad="1" is set (e.g. by Excel Services), Excel rebuilds
    the cache from source data on open — undoing our pruned sharedItems and
    causing pivot tables to crash while reconciling stale items blocks.
    Disabling it prevents the rebuild so our modifications are preserved.

    Uses pure string replacement to avoid ET serialization side-effects.
    """
    raw = session.read(cache_def_path)
    xml = raw.decode("utf-8")
    if 'refreshOnLoad="1"' in xml:
        xml = xml.replace('refreshOnLoad="1"', 'refreshOnLoad="0"', 1)
        session.write(cache_def_path, xml.encode("utf-8"))


def _find_pivot_cache_records_path(zf, cache_def_path: str) -> str | None:
    """Return ZIP path to the cache records file linked from cache_def_path, or None."""
    filename = cache_def_path.split("/")[-1]
    rels_path = f"xl/pivotCache/_rels/{filename}.rels"
    if rels_path not in zf.namelist():
        return None
    rels_root = ET.fromstring(zf.read(rels_path))
    for rel in rels_root.iter(f"{{{NS_RELS}}}Relationship"):
        if rel.get("Type") == RT_PIVOT_CACHE_RECORDS:
            return _resolve_zip_path("xl/pivotCache", rel.get("Target", ""))
    return None


def _get_used_shared_item_indices(zf, records_path: str, field_idx: int) -> set[int]:
    """Return the set of sharedItem indices for field_idx actually present in records."""
    records_root = ET.fromstring(zf.read(records_path))
    used: set[int] = set()
    for r in records_root:
        children = list(r)
        if field_idx < len(children):
            child = children[field_idx]
            if child.tag.split("}")[-1] == "x":
                v = child.get("v")
                if v is not None:
                    used.add(int(v))
    return used


def _recompute_shared_items_attrs(shared_el) -> None:
    """Recompute summary attributes on a <sharedItems> element after item removal.

    Only updates attributes that may have become inconsistent.  Attributes for
    types whose items are entirely absent are removed so Excel can recompute
    them on next open (absent optional attrs never trigger a repair prompt;
    *wrong* attrs do).
    """
    items = list(shared_el)
    tags  = {item.tag.split("}")[-1] for item in items}

    has_blank = "m" in tags
    has_str   = "s" in tags
    has_num   = "n" in tags
    has_date  = "d" in tags

    # containsBlank — 1 iff any <m/> present (ghost or active)
    if has_blank:
        shared_el.set("containsBlank", "1")
    else:
        shared_el.attrib.pop("containsBlank", None)

    # Numeric summary attrs — only safe if <n> items still exist
    if not has_num:
        for attr in ("containsNumber", "containsInteger", "minValue", "maxValue"):
            shared_el.attrib.pop(attr, None)

    # Date summary attrs — only safe if <d> items still exist
    if not has_date:
        for attr in ("containsDate", "containsNonDate", "minDate", "maxDate"):
            shared_el.attrib.pop(attr, None)

    # containsMixedTypes — true when more than one non-blank type present
    type_count = sum([has_str, has_num, has_date])
    if type_count > 1:
        shared_el.set("containsMixedTypes", "1")
    else:
        shared_el.attrib.pop("containsMixedTypes", None)

    # containsSemiMixedTypes — text values mixed with other types
    if has_str and (has_num or has_date):
        shared_el.set("containsSemiMixedTypes", "1")
    else:
        shared_el.attrib.pop("containsSemiMixedTypes", None)


def _update_si_attrs_str(si_tag: str, remaining_items: list[str]) -> str:
    """Update summary attributes in a <sharedItems …> opening tag string.

    Analyses *remaining_items* (raw XML strings like '<s v="Q1FY26"/>') and
    adjusts containsBlank / containsMixedTypes / containsSemiMixedTypes;
    drops numeric/date summary attrs when no such items remain.
    """
    has_blank = any(item.strip().startswith("<m") for item in remaining_items)
    has_str   = any(re.match(r"\s*<s\b", item) for item in remaining_items)
    has_num   = any(re.match(r"\s*<n\b", item) for item in remaining_items)
    has_date  = any(re.match(r"\s*<d\b", item) for item in remaining_items)

    def _set(tag: str, name: str, value: str | None) -> str:
        pat = re.compile(r"\s*\b" + re.escape(name) + r'="[^"]*"')
        if value is not None:
            if pat.search(tag):
                return pat.sub(f' {name}="{value}"', tag)
            return tag[:-1] + f' {name}="{value}">'   # insert before closing >
        return pat.sub("", tag)

    si_tag = _set(si_tag, "containsBlank", "1" if has_blank else None)

    if not has_num:
        for attr in ("containsNumber", "containsInteger", "minValue", "maxValue"):
            si_tag = _set(si_tag, attr, None)
    if not has_date:
        for attr in ("containsDate", "containsNonDate", "minDate", "maxDate"):
            si_tag = _set(si_tag, attr, None)

    type_count = sum([has_str, has_num, has_date])
    si_tag = _set(si_tag, "containsMixedTypes", "1" if type_count > 1 else None)
    if has_str and (has_num or has_date):
        si_tag = _set(si_tag, "containsSemiMixedTypes", "1")
    else:
        si_tag = _set(si_tag, "containsSemiMixedTypes", None)

    return si_tag


_SHARED_ITEM_RE  = re.compile(r"<(?:s|n|d|m|b|e)\b[^>]*/>\s*")
_RECORD_CHILD_RE = re.compile(r"<[a-z]\b[^>]*/>\s*")


def _prune_shared_items_str(
    xml: str,
    field_idx: int,
    remove_indices: set[int],
) -> tuple[str, dict[int, int]]:
    """Remove items at *remove_indices* from sharedItems of cacheField[field_idx].

    Pure string implementation — the original XML is preserved byte-for-byte
    except for the removed <item/> elements and the updated count/summary attrs.

    Returns (new_xml, remapping) where remapping = {old_idx: new_idx, …}.
    When remove_indices is empty, xml is returned unchanged.
    """
    # Locate cacheField[field_idx] by counting <cacheField start tags
    cf_starts = [m.start() for m in re.finditer(r"<cacheField\b", xml)]
    if field_idx >= len(cf_starts):
        return xml, {}

    cf_start = cf_starts[field_idx]
    cf_end   = cf_starts[field_idx + 1] if field_idx + 1 < len(cf_starts) else len(xml)
    cf_chunk = xml[cf_start:cf_end]

    # Locate <sharedItems …> opening tag
    si_open_m = re.search(r"<sharedItems\b[^>]*>", cf_chunk)
    if si_open_m is None:
        # Self-closing <sharedItems … /> — no items to prune; build identity mapping
        return xml, {}

    si_tag_str = si_open_m.group(0)                         # '<sharedItems count="18" …>'
    si_close   = cf_chunk.find("</sharedItems>", si_open_m.end())
    if si_close == -1:
        return xml, {}                                       # malformed — bail safely
    si_content = cf_chunk[si_open_m.end():si_close]

    # Find every item element within sharedItems
    items = list(_SHARED_ITEM_RE.finditer(si_content))
    n_old = len(items)

    # Build remapping {old_idx → new_idx} for retained items
    remapping: dict[int, int] = {}
    new_idx = 0
    for old_idx in range(n_old):
        if old_idx not in remove_indices:
            remapping[old_idx] = new_idx
            new_idx += 1

    if not remove_indices:
        return xml, remapping

    # Assemble new sharedItems content
    new_si_content = "".join(
        m.group(0) for i, m in enumerate(items) if i not in remove_indices
    )

    # Update count= in opening tag
    new_count = n_old - len(remove_indices)
    if re.search(r'\bcount="[^"]*"', si_tag_str):
        new_si_tag = re.sub(r'\bcount="[^"]*"', f'count="{new_count}"', si_tag_str)
    else:
        new_si_tag = si_tag_str[:-1] + f' count="{new_count}">'

    # Recompute summary attributes from remaining items
    remaining_strs = [m.group(0) for i, m in enumerate(items) if i not in remove_indices]
    new_si_tag = _update_si_attrs_str(new_si_tag, remaining_strs)

    # Splice back
    si_end       = si_close + len("</sharedItems>")
    new_cf_chunk = (
        cf_chunk[:si_open_m.start()]
        + new_si_tag
        + new_si_content
        + "</sharedItems>"
        + cf_chunk[si_end:]
    )
    return xml[:cf_start] + new_cf_chunk + xml[cf_end:], remapping


def _reindex_records_str(
    xml_bytes: bytes,
    field_idx: int,
    remapping: dict[int, int],
) -> bytes:
    """Rewrite <x v="N"/> at position *field_idx* in every <r> record row.

    Pure string implementation — only the v= attribute of the relevant <x>
    element is changed; everything else is preserved byte-for-byte.
    """
    xml = xml_bytes.decode("utf-8")

    def _process_record(m: re.Match) -> str:
        r_content = m.group(1)
        children  = list(_RECORD_CHILD_RE.finditer(r_content))
        if field_idx >= len(children):
            return m.group(0)
        child_m   = children[field_idx]
        child_str = child_m.group(0)
        if not child_str.lstrip().startswith("<x"):
            return m.group(0)
        v_m = re.search(r'\bv="(\d+)"', child_str)
        if v_m is None:
            return m.group(0)
        old_v = int(v_m.group(1))
        new_v = remapping.get(old_v)
        if new_v is None or new_v == old_v:
            return m.group(0)
        new_child     = re.sub(r'\bv="\d+"', f'v="{new_v}"', child_str, count=1)
        new_r_content = r_content[:child_m.start()] + new_child + r_content[child_m.end():]
        return "<r>" + new_r_content + "</r>"

    new_xml = re.sub(r"<r>(.*?)</r>", _process_record, xml, flags=re.DOTALL)
    return new_xml.encode("utf-8")


def _prune_cache(
    session: Session,
    cache_def_path: str,
    records_path: str | None,
    field_idx: int,
    remove_indices: set[int],
) -> dict[int, int]:
    """Remove items at remove_indices from sharedItems of cacheField[field_idx].

    Updates the cache definition (sharedItems count + elements) and the cache
    records (reindexes x="" references for field_idx) in the session.

    Uses pure string manipulation — no ElementTree serialisation — to preserve
    the original XML byte-for-byte and avoid artefacts that crash Excel's
    pivot cache parser (STATUS_STACK_BUFFER_OVERRUN in mso20win32client.dll).

    Returns remapping {old_idx: new_idx} for every retained item.
    Items in remove_indices are absent from the remapping.
    """
    cache_def_bytes = session.read(cache_def_path)
    xml = cache_def_bytes.decode("utf-8")

    new_xml, remapping = _prune_shared_items_str(xml, field_idx, remove_indices)

    if remove_indices:
        session.write(cache_def_path, new_xml.encode("utf-8"))

    # Reindex x= values in cache records when any index actually shifts
    if records_path and records_path in session.namelist():
        if any(old != new for old, new in remapping.items()):
            records_bytes     = session.read(records_path)
            new_records_bytes = _reindex_records_str(records_bytes, field_idx, remapping)
            if new_records_bytes != records_bytes:
                session.write(records_path, new_records_bytes)

    return remapping


def _remap_pivot_items_block(pt_str: str, field_idx: int, remapping: dict[int, int]) -> str:
    """Remap x="" indices in pivotField[field_idx]'s <items> block.

    Items whose old index is absent from the remapping are removed.
    The <items count="N"> attribute is updated to the new item count.
    Returns the updated pivot table XML string.
    """
    pf_starts = [m.start() for m in re.finditer(r"<pivotField(?![sS])\b", pt_str)]
    if len(pf_starts) <= field_idx:
        return pt_str

    pf_start = pf_starts[field_idx]
    pf_end   = pf_starts[field_idx + 1] if field_idx + 1 < len(pf_starts) else len(pt_str)
    pf_chunk = pt_str[pf_start:pf_end]

    items_open  = pf_chunk.find("<items")
    items_close = pf_chunk.find("</items>")
    if items_open == -1 or items_close == -1:
        return pt_str
    items_close += len("</items>")
    items_block = pf_chunk[items_open:items_close]

    def remap_item(m: re.Match) -> str:
        full = m.group(0)
        if 't="default"' in full:
            return full
        x_m = re.search(r'\bx="(\d+)"', full)
        if x_m is None:
            return full
        old_x = int(x_m.group(1))
        if old_x not in remapping:
            return ""   # item pruned from cache — drop it
        return re.sub(r'\bx="\d+"', f'x="{remapping[old_x]}"', full)

    new_items_block = re.sub(r"<item\b[^/]*/>\s*", lambda m: remap_item(m), items_block)

    # Update count attribute
    n_items = sum(1 for _ in re.finditer(r"<item\b", new_items_block))
    new_items_block = re.sub(
        r'(<items\b[^>]*\bcount=")[^"]*(")',
        lambda m: f'{m.group(1)}{n_items}{m.group(2)}',
        new_items_block,
    )

    new_pf_chunk = pf_chunk[:items_open] + new_items_block + pf_chunk[items_close:]
    return pt_str[:pf_start] + new_pf_chunk + pt_str[pf_end:]


_AXIS_LABELS = {
    "axisPage": "page filter",
    "axisRow":  "row filter",
    "axisCol":  "column filter",
}


def _apply_filter_group(
    session: Session,
    tabs: list[str],
    field_name: str,
    allowed: set[str],
    visible: set[str],
    vis_label: str,
) -> int:
    """Apply a pre-validated filter to all matching pivot tables.

    Two-phase approach:
      Phase 1 (per unique cache): prune items NOT in *allowed* that have NO
        records — these are truly stale and disappear from the dropdown.
        Items NOT in *allowed* but WITH records cannot be removed from the cache
        (they're still referenced), so they are left and will get h="1" in Ph 2.
      Phase 2 (per pivot table): set h="1" on items not in *visible* (SELECTED).
        Items in *allowed* but not *visible* remain in the dropdown but unchecked.
        Items not in *allowed* with records stay in the cache but are also h="1".

    *allowed*  — set of values to KEEP in the dropdown (used for cache pruning).
                 Pass an empty set to skip cache pruning.
    *visible*  — set of values to mark as SELECTED/checked (must be ⊆ allowed).

    Returns the count of pivot tables actually modified.
    """
    # Effective allowed set for pruning: if not specified, don't prune anything
    prune_guard = allowed if allowed else None

    # Binary-test flags (set in .env to isolate issues):
    #   PIVOT_FILTER_SKIP_PT_WRITE=true — skip all pivot-table writes (cache h-filter only)
    #   PIVOT_FILTER_SKIP_REMAP=true    — skip _remap_pivot_items_block (keep all items as-is)
    skip_pt_write = os.environ.get("PIVOT_FILTER_SKIP_PT_WRITE", "").strip().lower() == "true"
    skip_remap    = os.environ.get("PIVOT_FILTER_SKIP_REMAP",    "").strip().lower() == "true"
    if skip_pt_write:
        print("  [SKIP_PT_WRITE] Phase 2 pivot-table writes disabled (test mode).")
    if skip_remap:
        print("  [SKIP_REMAP]    Stale-item removal skipped; only h-filter applied (test mode).")

    total_fixed = 0

    for tab_name in tabs:
        sheet_file = _find_sheet_file(session, tab_name)
        if sheet_file is None:
            print(f"  WARNING: Tab '{tab_name}' not found -- skipping.")
            continue
        print(f"\n  Tab '{tab_name}' -> {sheet_file}")

        pt_paths = _find_pivot_tables_on_sheet(session, sheet_file)
        if not pt_paths:
            print("    No pivot tables found on this tab.")
            continue
        print(f"    {len(pt_paths)} pivot table(s): {[p.split('/')[-1] for p in pt_paths]}")

        # --- Phase 1: cache surgery (one pass per unique cache) ---
        # remapping_for_cache[cache_path] = {old_idx: new_idx, ...}
        # idx_to_val_for_cache[cache_path] = {new_idx: value, ...}  (post-prune)
        # pruned_any_for_cache[cache_path] = True if any items were actually removed
        remapping_for_cache: dict[str, dict[int, int]] = {}
        idx_to_val_for_cache: dict[str, dict[int, str]] = {}
        pruned_any_for_cache: dict[str, bool] = {}

        for pt_path in sorted(pt_paths):
            cache_path = _find_pivot_cache_path(session, pt_path)
            if cache_path is None or cache_path in remapping_for_cache:
                continue  # already processed or no cache

            # Skip caches managed by Excel auto-refresh (refreshOnLoad="1").
            # These caches are validated against live source data when the file
            # opens.  Pruning their sharedItems causes a fatal mismatch that
            # crashes Excel.  Leaving them untouched lets Excel's auto-refresh
            # rebuild them correctly; we also skip h-filtering their pivot
            # tables because Excel would override those changes on rebuild anyway.
            cache_raw = session.read(cache_path).decode("utf-8")
            if 'refreshOnLoad="1"' in cache_raw:
                print(f"    [{cache_path.split('/')[-1]}] Skipping "
                      f"(refreshOnLoad='1' — managed by Excel auto-refresh).")
                remapping_for_cache[cache_path] = {}
                idx_to_val_for_cache[cache_path] = {}
                continue

            field_idx, idx_to_val = _find_field_in_cache(session, cache_path, field_name)
            if field_idx is None:
                remapping_for_cache[cache_path] = {}   # sentinel: field not in this cache
                idx_to_val_for_cache[cache_path] = {}
                continue

            # Cache pruning permanently disabled.
            # Pruning sharedItems then remapping pivot-table items blocks triggers
            # STATUS_STACK_BUFFER_OVERRUN (0xc0000409) at mso20win32client.dll
            # +0x3d27f2 regardless of how the items block is formatted.
            # h-filter (h="1" on non-ALLOWED items) achieves the same result for
            # the dropdown without changing any cache index — no remapping needed,
            # no Excel crash.  Stale entries accumulate in the cache over time but
            # are invisible to the user because they are hidden via h-filter.
            remapping = {i: i for i in idx_to_val}
            remapping_for_cache[cache_path] = remapping
            idx_to_val_for_cache[cache_path] = dict(idx_to_val)
            print(f"    [{cache_path.split('/')[-1]}] {len(idx_to_val)} item(s) "
                  f"-> Phase 2 h-filter (cache untouched).")

        # --- Phase 2: h-attribute filter on each pivot table ---
        for pt_path in sorted(pt_paths):
            pt_name = pt_path.split("/")[-1]

            cache_path = _find_pivot_cache_path(session, pt_path)
            if cache_path is None:
                continue

            remapping   = remapping_for_cache.get(cache_path)
            idx_to_val  = idx_to_val_for_cache.get(cache_path)
            if not remapping and not idx_to_val:
                continue  # field not found in this cache

            pt_original = session.read(pt_path).decode("utf-8")
            pt_str = pt_original

            # Need field_idx from the original cache read; it stays the same after pruning
            # (pruning only changes values within sharedItems, not the field position)
            field_idx, _ = _find_field_in_cache(session, cache_path, field_name)
            if field_idx is None:
                continue

            # Always remap x= indices FIRST — even for non-filterable fields.
            # Pivot tables on the configured tab that use the same pruned cache must
            # have their stale x= references cleaned up regardless of axis type.
            # Without this, pivotTableN.xml files with non-filterable axes keep stale
            # x= values that point past the end of the pruned sharedItems → corruption.
            if remapping and not skip_remap:
                pt_str = _remap_pivot_items_block(pt_str, field_idx, remapping)

            axis = _get_field_axis(pt_str, field_idx)
            if axis not in _AXIS_LABELS:
                # Not a filterable axis: no h-filter, but still save remap changes.
                if pt_str != pt_original:
                    if not skip_pt_write:
                        session.write(pt_path, pt_str.encode("utf-8"))
                        print(f"    [{pt_name}] stale items dropped (non-filterable axis, remap only).")
                    else:
                        print(f"    [{pt_name}] [SKIP_PT_WRITE] would drop stale items (non-filterable axis).")
                    total_fixed += 1
                continue
            axis_label = _AXIS_LABELS[axis]

            print(f"    [{pt_name}] field[{field_idx}] '{_s(field_name)}' ({axis_label}, {vis_label}):")
            unmatched = visible - set(idx_to_val.values())
            if unmatched:
                print(f"      WARNING: values not in cache: {[_s(x) for x in sorted(unmatched)]}")
            for i, v in sorted(idx_to_val.items()):
                if v in visible:
                    mark = " [SHOW+check]"
                elif not allowed or v in allowed:
                    mark = " [SHOW-uncheck]"
                else:
                    mark = " [hide]"
                print(f"      [{i}] {_s(v)!r}{mark}")

            # Apply h-filter on top of the (possibly remapped) pt_str.
            # Returns None when no h= changes are needed — but the remap may still
            # have changed pt_str, so we compare against the original either way.
            new_pt_str = _apply_filter_to_pivot(pt_str, field_idx, idx_to_val, visible)
            final_str  = new_pt_str if new_pt_str is not None else pt_str

            if final_str == pt_original:
                print(f"      -> already correct, no change.")
                continue

            n_in_dropdown = sum(1 for v in idx_to_val.values()
                                if not allowed or v in allowed)
            n_checked = sum(1 for v in idx_to_val.values() if v in visible)
            remap_note = " (incl. index remap)" if new_pt_str is None else ""
            if skip_pt_write:
                print(f"      -> [SKIP_PT_WRITE] would stage ({n_in_dropdown} item(s), "
                      f"{n_checked} checked{remap_note}) — not written.")
            else:
                session.write(pt_path, final_str.encode("utf-8"))
                print(f"      -> staged ({n_in_dropdown} item(s) in dropdown, "
                      f"{n_checked} checked{remap_note}).")
            total_fixed += 1

        # --- Phase 2.5: remap x-indices in pivot tables on OTHER sheets ---
        # Skipped in SKIP_REMAP mode (stale items are intentionally kept).
        if skip_remap:
            continue
        # Phase 1 pruned sharedItems in caches; any pivot table referencing those
        # caches must have its x="" indices remapped to match, even if it lives
        # on a sheet not listed in PIVOT_FILTER_N_TABS.  Without this, stale
        # x= values on other sheets cause Excel to corrupt/discard those tables.
        all_pt_paths = sorted(
            n for n in session.namelist()
            if n.startswith("xl/pivotTables/") and n.endswith(".xml")
        )
        for other_pt_path in all_pt_paths:
            if other_pt_path in pt_paths:
                continue  # already handled in Phase 2
            cache_path = _find_pivot_cache_path(session, other_pt_path)
            if cache_path is None:
                continue
            remapping = remapping_for_cache.get(cache_path)
            if not remapping:
                continue  # sentinel (field not in this cache) or nothing pruned
            # Skip only when remapping is identity AND nothing was actually pruned.
            # Even an identity remap (0→0, 1→1, …) must be applied when items were
            # pruned: _remap_pivot_items_block drops items whose x= is no longer in
            # the remapping (the removed ones), which is required even when kept items
            # didn't change their indices.
            if (all(old == new for old, new in remapping.items())
                    and not pruned_any_for_cache.get(cache_path, False)):
                continue
            field_idx, _ = _find_field_in_cache(session, cache_path, field_name)
            if field_idx is None:
                continue
            pt_bytes = session.read(other_pt_path)
            pt_str = pt_bytes.decode("utf-8")
            new_pt_str = _remap_pivot_items_block(pt_str, field_idx, remapping)
            if new_pt_str != pt_str:
                if skip_pt_write:
                    print(f"    [{other_pt_path.split('/')[-1]}] [SKIP_PT_WRITE] would remap "
                          f"(shared cache, other tab) — not written.")
                else:
                    session.write(other_pt_path, new_pt_str.encode("utf-8"))
                    print(f"    [{other_pt_path.split('/')[-1]}] x-values remapped "
                          f"(shared cache, other tab — index-remap only, no h-filter).")

    return total_fixed


def fix_pivot_filter(session: Session) -> bool:
    """Apply all PIVOT_FILTER_N_* groups to matching pivot tables.

    Scans groups N = 1, 2, 3 … stopping when PIVOT_FILTER_N_FIELD is absent.
    Each group reads:
      PIVOT_FILTER_N_TABS      Comma-separated tab names (default: "By Sales")
      PIVOT_FILTER_N_FIELD     Partial field name (case-insensitive, first-line match)
      PIVOT_FILTER_N_ALLOWED   Whitelist — only these values appear in the dropdown;
                               everything else (stale entries) gets h="1"
      PIVOT_FILTER_N_SELECTED  Currently checked values — must be a subset of ALLOWED.
                               ALLOWED-but-not-SELECTED items remain in the dropdown
                               but are unchecked (h="1").

    Validation: if SELECTED contains values absent from ALLOWED, the group is
    skipped entirely (no modification) and an error is displayed.

    Applies to page, row, and column fields — any pivot field with an <items> block.
    """
    def _parse_set(key: str) -> set[str]:
        raw = os.environ.get(key, "").strip()
        return {v.strip() for v in raw.split(",") if v.strip()}

    total_fixed = 0
    n = 1

    while True:
        field_name = os.environ.get(f"PIVOT_FILTER_{n}_FIELD", "").strip()
        if not field_name:
            if n == 1:
                print("  ERROR: PIVOT_FILTER_1_FIELD not set in environment.")
                return False
            break  # no more groups

        tabs_raw = os.environ.get(f"PIVOT_FILTER_{n}_TABS", "By Sales").strip()
        tabs = [t.strip() for t in tabs_raw.split(",") if t.strip()]

        selected = _parse_set(f"PIVOT_FILTER_{n}_SELECTED")
        allowed  = _parse_set(f"PIVOT_FILTER_{n}_ALLOWED")

        print(f"\n  === Filter group {n} ===")
        print(f"  Tabs    : {tabs}")
        print(f"  Field   : {_s(field_name)!r}")
        if allowed:
            print(f"  Allowed : {[_s(x) for x in sorted(allowed)]}")
        if selected:
            print(f"  Selected: {[_s(x) for x in sorted(selected)]}")

        # --- Validate ---
        if not selected and not allowed:
            print(f"  ERROR: PIVOT_FILTER_{n}_SELECTED and ALLOWED are both unset -- skipping.")
            n += 1
            continue

        if selected and allowed:
            invalid = selected - allowed
            if invalid:
                print(f"  ERROR: SELECTED values not present in ALLOWED:")
                for v in sorted(invalid):
                    print(f"    - {_s(v)!r}")
                print(f"  Group {n} skipped -- fix PIVOT_FILTER_{n}_SELECTED first.")
                n += 1
                continue

        # --- Compute visible set ---
        # SELECTED takes priority; fall back to ALLOWED (show all valid entries)
        if selected:
            visible   = selected
            vis_label = f"SELECTED ({len(selected)})"
        else:
            visible   = allowed
            vis_label = f"ALLOWED ({len(allowed)}, SELECTED not set -- all valid entries shown)"

        total_fixed += _apply_filter_group(session, tabs, field_name, allowed, visible, vis_label)
        n += 1

    if total_fixed == 0:
        print("\n  No changes needed -- all filters already match configuration.")
        return False

    print(f"\n  Changes staged in session ({total_fixed} pivot table(s) updated, not yet saved).")
    return True


# ---------------------------------------------------------------------------
# Menu system
# ---------------------------------------------------------------------------

FIXES = [
    ("Fix By Sales Charts",         fix_by_sales_charts),
    ("Pivot page-filter selection",  fix_pivot_filter),
]


def _print_header(env_label: str, input_path: str, output_path: str, session: Session) -> None:
    print("")
    print("=" * 60)
    print("  FixPipe - Excel Pipe File Repair Tool")
    print(f"  Environment : {env_label}")
    print(f"  Input  (ref): {input_path}")
    print(f"  Output (fix): {output_path}")
    if session.applied_fixes:
        print(f"  Pending     : {len(session.applied_fixes)} fix(es) staged, not yet saved")
        for name in session.applied_fixes:
            print(f"                - {name}")
    else:
        print("  Pending     : none")
    print("=" * 60)


def main():
    suffix = load_env_from_config()

    input_path  = os.environ.get("INPUT_SUIVI_RAW",  "").strip().strip('"')
    output_path = os.environ.get("OUTPUT_SUIVI_RAW", "").strip().strip('"')

    env_label = f"{suffix}  (.env.{suffix})" if suffix else "Production  (.env)"

    if not input_path or not output_path:
        print("\nERROR: INPUT_SUIVI_RAW and OUTPUT_SUIVI_RAW must be set in your .env file.")
        return

    if not os.path.exists(input_path):
        print(f"\nERROR: Input file not found: {input_path}")
        return

    session = Session(input_path)

    if os.path.normpath(input_path) == os.path.normpath(output_path):
        print("\n  WARNING: INPUT and OUTPUT point to the same file!")
        print("  The original will be overwritten on save. Proceed with caution.\n")

    while True:
        _print_header(env_label, input_path, output_path, session)

        print("\nAvailable fixes:")
        for i, (label, _) in enumerate(FIXES, 1):
            print(f"  {i}. {label}")
        exit_label = "Save and exit" if session.dirty else "Exit"
        print(f"  0. {exit_label}")

        choice = input("\nSelect fix (0-{}): ".format(len(FIXES))).strip()

        if choice == "0":
            if session.dirty:
                print(f"\nSaving {len(session.applied_fixes)} fix(es) -> {output_path} ...")
                session.save(output_path)
                print("Done.")
            else:
                print("No changes -- nothing to save.")
            print("Bye.")
            break

        try:
            idx = int(choice) - 1
            if idx < 0 or idx >= len(FIXES):
                raise ValueError
        except ValueError:
            print("Invalid selection.")
            continue

        label, func = FIXES[idx]
        print(f"\nRunning: {label}")
        print("-" * 40)
        changed = func(session)
        if changed:
            session.applied_fixes.append(label)
        print("-" * 40)


if __name__ == "__main__":
    main()
