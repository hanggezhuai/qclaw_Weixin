#!/usr/bin/env python3
"""Read-only WeChat ID scanner for Windows.

This tool focuses on extracting:
1. personal account identifiers like ``wxid_xxx``
2. group chat identifiers like ``123456@chatroom``

It intentionally avoids database decryption. Instead it combines:
- account directory discovery under common WeChat data roots
- raw byte scanning of small/high-value files
- process memory scanning of the running WeChat client

That gives a practical "one-click fetch" path, plus a keyword query mode
that searches nearby context captured from the source bytes.
"""

from __future__ import annotations

import argparse
import ctypes as ct
from ctypes import wintypes as wt
import datetime as dt
import json
import os
from pathlib import Path
import re
import sys
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple


ASCII_ID_RE = re.compile(rb"wxid_[0-9A-Za-z_\-]{6,64}|[0-9A-Za-z_\-]{3,100}@chatroom")
UTF16_CHATROOM_RE = re.compile(rb"(?:[0-9A-Za-z_\-]\x00){3,100}@\x00c\x00h\x00a\x00t\x00r\x00o\x00o\x00m\x00")
UTF16_WXID_RE = re.compile(rb"w\x00x\x00i\x00d\x00_\x00(?:[0-9A-Za-z_\-]\x00){6,64}")
ACCOUNT_DIR_RE = re.compile(r"^wxid_[0-9A-Za-z_\-]{6,64}$")
STRICT_WXID_RE = re.compile(r"(?<![0-9A-Za-z_\-])(wxid_[0-9a-z]{6,32})(?![0-9A-Za-z_\-])")
STRICT_CHATROOM_RE = re.compile(r"(?<![0-9A-Za-z_\-])([0-9A-Za-z_\-]{3,50}@chatroom)(?![0-9A-Za-z_\-])")

PRINTABLE_ASCII_RE = re.compile(r"[ -~]{3,}")
TOKEN_RE = re.compile(r"[\u4e00-\u9fffA-Za-z0-9_@\-]{2,}")

MAX_FILE_SCAN_BYTES = 64 * 1024 * 1024
DEFAULT_REGION_LIMIT = 8 * 1024 * 1024
DEFAULT_CONTEXT_BYTES = 160
MAX_CONTEXTS_PER_RECORD = 8
MAX_SOURCES_PER_RECORD = 12

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010
MEM_COMMIT = 0x1000
PAGE_GUARD = 0x100
READABLE_PAGE_PROTECT = {0x02, 0x04, 0x08, 0x20, 0x40, 0x80}


def is_windows() -> bool:
    return os.name == "nt"


class MEMORY_BASIC_INFORMATION(ct.Structure):
    _fields_ = [
        ("BaseAddress", ct.c_void_p),
        ("AllocationBase", ct.c_void_p),
        ("AllocationProtect", wt.DWORD),
        ("RegionSize", ct.c_size_t),
        ("State", wt.DWORD),
        ("Protect", wt.DWORD),
        ("Type", wt.DWORD),
    ]


class ScanError(RuntimeError):
    pass


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")


def normalize_text(text: str) -> str:
    filtered = []
    for ch in text:
        if ch in "\r\n\t":
            filtered.append(" ")
        elif ch.isprintable() or ("\u4e00" <= ch <= "\u9fff"):
            filtered.append(ch)
        else:
            filtered.append(" ")
    return " ".join("".join(filtered).split())


def decode_context_snippets(blob: bytes) -> List[str]:
    snippets: List[str] = []

    ascii_text = normalize_text(blob.decode("utf-8", errors="ignore"))
    for match in PRINTABLE_ASCII_RE.finditer(ascii_text):
        piece = match.group(0).strip()
        if piece:
            snippets.append(piece)

    utf16_len = len(blob) - (len(blob) % 2)
    if utf16_len:
        utf16_text = normalize_text(blob[:utf16_len].decode("utf-16le", errors="ignore"))
        for piece in TOKEN_RE.findall(utf16_text):
            if piece:
                snippets.append(piece)

    deduped: List[str] = []
    seen = set()
    for snippet in snippets:
        snippet = snippet.strip()
        if len(snippet) < 2:
            continue
        if snippet in seen:
            continue
        seen.add(snippet)
        deduped.append(snippet[:240])
        if len(deduped) >= 10:
            break
    return deduped


def classify_identifier(identifier: str) -> str:
    return "chatroom" if identifier.endswith("@chatroom") else "wxid"


def strict_identifiers(text: str) -> List[str]:
    found: List[str] = []
    for regex in (STRICT_WXID_RE, STRICT_CHATROOM_RE):
        for match in regex.finditer(text):
            found.append(match.group(1))
    return unique_preserve_order(found)


def ensure_output_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def unique_preserve_order(values: Iterable[str]) -> List[str]:
    seen = set()
    ordered = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        ordered.append(value)
    return ordered


class ResultStore:
    def __init__(self) -> None:
        self.records: Dict[str, Dict[str, object]] = {}

    def add(
        self,
        identifier: str,
        source_kind: str,
        source_ref: str,
        *,
        location: Optional[str] = None,
        contexts: Optional[Sequence[str]] = None,
        meta: Optional[Dict[str, object]] = None,
    ) -> None:
        identifier = identifier.strip()
        if not identifier:
            return

        record = self.records.get(identifier)
        if record is None:
            record = {
                "id": identifier,
                "kind": classify_identifier(identifier),
                "sources": [],
                "contexts": [],
                "evidence_count": 0,
            }
            self.records[identifier] = record

        source_entry: Dict[str, object] = {
            "kind": source_kind,
            "ref": source_ref,
        }
        if location:
            source_entry["location"] = location
        if meta:
            source_entry.update(meta)

        sources: List[Dict[str, object]] = record["sources"]  # type: ignore[assignment]
        if len(sources) < MAX_SOURCES_PER_RECORD:
            if source_entry not in sources:
                sources.append(source_entry)

        record["evidence_count"] = int(record["evidence_count"]) + 1

        if contexts:
            existing_contexts: List[str] = record["contexts"]  # type: ignore[assignment]
            merged = unique_preserve_order(existing_contexts + [c for c in contexts if c])
            record["contexts"] = merged[:MAX_CONTEXTS_PER_RECORD]

    def to_json_ready(self) -> List[Dict[str, object]]:
        return sorted(
            self.records.values(),
            key=lambda item: (
                0 if item["kind"] == "wxid" else 1,
                -int(item["evidence_count"]),
                str(item["id"]),
            ),
        )


def default_data_roots() -> List[Path]:
    home = Path.home()
    candidates = [
        home / "Documents" / "WeChat Files",
        home / "Documents" / "wechatData" / "WeChat Files",
    ]
    return [path for path in candidates if path.exists()]


def discover_account_dirs(roots: Sequence[Path]) -> List[Path]:
    found: List[Path] = []
    for root in roots:
        if not root.exists():
            continue
        try:
            for child in root.iterdir():
                if child.is_dir() and ACCOUNT_DIR_RE.match(child.name):
                    found.append(child)
        except OSError:
            continue
    return sorted({path.resolve() for path in found})


def iter_candidate_files(account_dir: Path) -> Iterator[Path]:
    direct_names = {"AccInfo.dat", "aconfig.dat", "config01.dat"}
    config_dir = account_dir / "config"
    if config_dir.exists():
        for path in config_dir.rglob("*"):
            if path.is_file() and (path.name in direct_names or path.suffix.lower() in {".dat", ".ini"}):
                yield path

    msg_dir = account_dir / "Msg"
    if msg_dir.exists():
        for path in msg_dir.iterdir():
            if path.is_file():
                yield path

    backup_dir = account_dir / "Backup"
    if backup_dir.exists():
        for path in backup_dir.rglob("*"):
            if path.is_file() and path.stat().st_size <= 8 * 1024 * 1024:
                yield path


def scan_bytes_for_ids(data: bytes) -> Iterator[Tuple[str, int, List[str]]]:
    for match in ASCII_ID_RE.finditer(data):
        try:
            decoded = match.group(0).decode("utf-8", errors="ignore")
        except Exception:
            continue
        identifiers = strict_identifiers(decoded)
        if not identifiers:
            continue
        start = max(0, match.start() - DEFAULT_CONTEXT_BYTES)
        end = min(len(data), match.end() + DEFAULT_CONTEXT_BYTES)
        contexts = decode_context_snippets(data[start:end])
        for identifier in identifiers:
            yield identifier, match.start(), contexts

    for regex in (UTF16_CHATROOM_RE, UTF16_WXID_RE):
        for match in regex.finditer(data):
            try:
                decoded = match.group(0).decode("utf-16le", errors="ignore")
            except Exception:
                continue
            identifiers = strict_identifiers(decoded)
            if not identifiers:
                continue
            start = max(0, match.start() - DEFAULT_CONTEXT_BYTES)
            end = min(len(data), match.end() + DEFAULT_CONTEXT_BYTES)
            contexts = decode_context_snippets(data[start:end])
            for identifier in identifiers:
                yield identifier, match.start(), contexts


def scan_files(account_dirs: Sequence[Path], store: ResultStore) -> Dict[str, object]:
    scanned_files = 0
    matched_files = 0
    errors: List[str] = []

    for account_dir in account_dirs:
        store.add(account_dir.name, "account_directory", str(account_dir))

        for path in iter_candidate_files(account_dir):
            scanned_files += 1
            try:
                size = path.stat().st_size
            except OSError as exc:
                errors.append(f"{path}: {exc}")
                continue

            if size <= 0 or size > MAX_FILE_SCAN_BYTES:
                continue

            try:
                with path.open("rb") as handle:
                    data = handle.read()
            except OSError as exc:
                errors.append(f"{path}: {exc}")
                continue

            file_had_hit = False
            for identifier, offset, contexts in scan_bytes_for_ids(data):
                file_had_hit = True
                store.add(
                    identifier,
                    "file_scan",
                    str(path),
                    location=f"offset:{offset}",
                    contexts=contexts,
                    meta={"size": size},
                )
            if file_had_hit:
                matched_files += 1

    return {
        "scanned_files": scanned_files,
        "matched_files": matched_files,
        "errors": errors[:50],
    }


def enum_wechat_processes(path_filter: Optional[Path] = None) -> List[Dict[str, object]]:
    import subprocess

    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        "Get-Process WeChat,Weixin -ErrorAction SilentlyContinue | "
        "Select-Object Id,ProcessName,Path,MainWindowTitle | ConvertTo-Json -Compress",
    ]
    proc = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=False,
    )
    if not proc.stdout.strip():
        return []

    try:
        payload = json.loads(proc.stdout)
    except json.JSONDecodeError:
        return []

    items = payload if isinstance(payload, list) else [payload]
    normalized: List[Dict[str, object]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        path = item.get("Path")
        if path_filter and path:
            try:
                if Path(path).resolve() != path_filter.resolve():
                    continue
            except OSError:
                continue
        normalized.append(
            {
                "pid": int(item["Id"]),
                "process_name": item.get("ProcessName") or "WeChat",
                "path": path,
                "window_title": item.get("MainWindowTitle"),
            }
        )
    return normalized


def open_process(pid: int):
    k32 = ct.WinDLL("kernel32", use_last_error=True)
    handle = k32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
    if not handle:
        raise OSError(f"OpenProcess({pid}) failed with error {ct.get_last_error()}")
    return k32, handle


def scan_process_memory(
    process: Dict[str, object],
    store: ResultStore,
    *,
    region_size_limit: int,
) -> Dict[str, object]:
    pid = int(process["pid"])
    k32, handle = open_process(pid)

    virtual_query_ex = k32.VirtualQueryEx
    virtual_query_ex.argtypes = [wt.HANDLE, ct.c_void_p, ct.POINTER(MEMORY_BASIC_INFORMATION), ct.c_size_t]
    virtual_query_ex.restype = ct.c_size_t

    read_process_memory = k32.ReadProcessMemory
    read_process_memory.argtypes = [wt.HANDLE, ct.c_void_p, ct.c_void_p, ct.c_size_t, ct.POINTER(ct.c_size_t)]
    read_process_memory.restype = wt.BOOL

    close_handle = k32.CloseHandle

    addr = 0
    regions = 0
    readable_regions = 0
    bytes_read = 0
    errors = 0

    try:
        while True:
            mbi = MEMORY_BASIC_INFORMATION()
            queried = virtual_query_ex(handle, ct.c_void_p(addr), ct.byref(mbi), ct.sizeof(mbi))
            if not queried:
                break

            base = int(mbi.BaseAddress or 0)
            size = int(mbi.RegionSize or 0)
            if size <= 0:
                addr += 0x1000
                continue

            regions += 1
            protect = mbi.Protect & 0xFF
            readable = (
                mbi.State == MEM_COMMIT
                and not (mbi.Protect & PAGE_GUARD)
                and protect in READABLE_PAGE_PROTECT
                and size <= region_size_limit
            )
            if readable:
                readable_regions += 1
                try:
                    buffer = (ct.c_char * size)()
                    read = ct.c_size_t()
                    ok = read_process_memory(handle, ct.c_void_p(base), buffer, size, ct.byref(read))
                    if ok and read.value:
                        chunk = bytes(buffer[: read.value])
                        bytes_read += len(chunk)
                        for identifier, offset, contexts in scan_bytes_for_ids(chunk):
                            store.add(
                                identifier,
                                "process_memory",
                                f"pid:{pid}",
                                location=hex(base + offset),
                                contexts=contexts,
                                meta={
                                    "process_path": process.get("path"),
                                    "window_title": process.get("window_title"),
                                },
                            )
                except Exception:
                    errors += 1

            next_addr = base + size
            if next_addr <= addr:
                break
            addr = next_addr
            if addr >= 0x7FFFFFFFFFFF:
                break
    finally:
        close_handle(handle)

    return {
        "pid": pid,
        "path": process.get("path"),
        "regions_scanned": regions,
        "readable_regions_scanned": readable_regions,
        "bytes_read": bytes_read,
        "read_errors": errors,
    }


def make_output_payload(
    *,
    args: argparse.Namespace,
    account_dirs: Sequence[Path],
    store: ResultStore,
    file_summary: Dict[str, object],
    process_summaries: Sequence[Dict[str, object]],
) -> Dict[str, object]:
    records = store.to_json_ready()
    return {
        "generated_at": now_iso(),
        "mode": "scan",
        "host": os.environ.get("COMPUTERNAME"),
        "scan_options": {
            "include_memory": args.include_memory,
            "region_size_limit": args.region_size_limit,
            "roots": [str(p) for p in args.roots],
            "wechat_path": str(args.wechat_path) if args.wechat_path else None,
        },
        "accounts": [str(path) for path in account_dirs],
        "summary": {
            "wxid_count": sum(1 for item in records if item["kind"] == "wxid"),
            "chatroom_count": sum(1 for item in records if item["kind"] == "chatroom"),
            "total_ids": len(records),
        },
        "file_scan": file_summary,
        "memory_scan": list(process_summaries),
        "records": records,
    }


def write_scan_output(payload: Dict[str, object], output_dir: Path) -> Path:
    ensure_output_dir(output_dir)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    path = output_dir / f"wechat-id-scan-{stamp}.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def latest_scan_file(output_dir: Path) -> Path:
    files = sorted(output_dir.glob("wechat-id-scan-*.json"))
    if not files:
        raise ScanError(f"No scan files found in {output_dir}")
    return files[-1]


def load_scan_payload(path: Path) -> Dict[str, object]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ScanError(str(exc)) from exc
    except json.JSONDecodeError as exc:
        raise ScanError(f"Invalid JSON in {path}: {exc}") from exc


def score_match(keyword: str, record: Dict[str, object]) -> Tuple[int, List[str]]:
    lowered = keyword.lower()
    hits: List[str] = []
    score = 0

    record_id = str(record["id"])
    if lowered in record_id.lower():
        score += 100
        hits.append(f"id:{record_id}")

    for context in record.get("contexts", []):
        text = str(context)
        if lowered in text.lower():
            score += 30
            hits.append(text)

    for source in record.get("sources", []):
        ref = str(source.get("ref", ""))
        if lowered in ref.lower():
            score += 10
            hits.append(f"ref:{ref}")

    return score, unique_preserve_order(hits)[:5]


def cmd_scan(args: argparse.Namespace) -> int:
    default_roots = default_data_roots()
    extra_roots = [Path(p).expanduser() for p in (args.roots or [])]
    roots = unique_preserve_order(str(path) for path in (default_roots + extra_roots))
    roots = [Path(p) for p in roots]
    args.roots = roots
    account_dirs = discover_account_dirs(roots)
    if not account_dirs:
        raise ScanError("No WeChat account directories were found under the configured roots.")

    store = ResultStore()
    file_summary = scan_files(account_dirs, store)

    process_summaries: List[Dict[str, object]] = []
    if args.include_memory:
        processes = enum_wechat_processes(args.wechat_path)
        for process in processes:
            process_summaries.append(
                scan_process_memory(process, store, region_size_limit=args.region_size_limit)
            )

    payload = make_output_payload(
        args=args,
        account_dirs=account_dirs,
        store=store,
        file_summary=file_summary,
        process_summaries=process_summaries,
    )
    output_path = write_scan_output(payload, args.output_dir)

    print(f"scan_file={output_path}")
    print(f"accounts={len(account_dirs)}")
    print(f"wxid_count={payload['summary']['wxid_count']}")
    print(f"chatroom_count={payload['summary']['chatroom_count']}")
    print(f"total_ids={payload['summary']['total_ids']}")

    sample_chatrooms = [item["id"] for item in payload["records"] if item["kind"] == "chatroom"][:10]
    if sample_chatrooms:
        print("sample_chatrooms=")
        for item in sample_chatrooms:
            print(f"  {item}")
    return 0


def cmd_query(args: argparse.Namespace) -> int:
    if args.input:
        path = args.input
    else:
        path = latest_scan_file(args.output_dir)

    payload = load_scan_payload(path)
    records = payload.get("records", [])
    if not isinstance(records, list):
        raise ScanError(f"Unexpected records structure in {path}")

    matches = []
    for item in records:
        if not isinstance(item, dict):
            continue
        if args.kind and item.get("kind") != args.kind:
            continue
        score, hits = score_match(args.keyword, item)
        if score:
            matches.append((score, item, hits))

    matches.sort(key=lambda row: (-row[0], row[1].get("kind", ""), row[1].get("id", "")))

    print(f"scan_file={path}")
    print(f"keyword={args.keyword}")
    print(f"matches={len(matches)}")

    for score, item, hits in matches[: args.limit]:
        print()
        print(f"id={item['id']}")
        print(f"kind={item['kind']}")
        print(f"score={score}")
        print(f"evidence_count={item.get('evidence_count', 0)}")
        contexts = item.get("contexts", [])
        if contexts:
            print("contexts=")
            for text in contexts[:4]:
                print(f"  {text}")
        if hits:
            print("matched_on=")
            for hit in hits[:4]:
                print(f"  {hit}")
        sources = item.get("sources", [])
        if sources:
            print("sources=")
            for source in sources[:3]:
                ref = source.get("ref")
                location = source.get("location")
                if location:
                    print(f"  {source.get('kind')} {ref} {location}")
                else:
                    print(f"  {source.get('kind')} {ref}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read-only WeChat wxid/chatroom scanner")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan local WeChat data and process memory")
    scan_parser.add_argument(
        "--root",
        dest="roots",
        action="append",
        help="Extra data root to scan. May be used multiple times.",
    )
    scan_parser.add_argument(
        "--wechat-path",
        type=Path,
        help="Optional exact WeChat.exe path filter for memory scanning.",
    )
    scan_parser.add_argument(
        "--no-memory",
        dest="include_memory",
        action="store_false",
        help="Skip process memory scanning and only use filesystem evidence.",
    )
    scan_parser.add_argument(
        "--region-size-limit",
        type=int,
        default=DEFAULT_REGION_LIMIT,
        help=f"Maximum readable region size in bytes for process scanning. Default: {DEFAULT_REGION_LIMIT}",
    )
    scan_parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("out"),
        help="Where scan JSON files are written. Default: ./out",
    )
    scan_parser.set_defaults(include_memory=True, func=cmd_scan)

    query_parser = subparsers.add_parser("query", help="Search the latest scan JSON by keyword")
    query_parser.add_argument("keyword", help="Keyword to search in id/context/source")
    query_parser.add_argument(
        "--input",
        type=Path,
        help="Specific scan JSON file. Defaults to the latest file in --output-dir.",
    )
    query_parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("out"),
        help="Directory containing scan JSON files. Default: ./out",
    )
    query_parser.add_argument(
        "--kind",
        choices=("wxid", "chatroom"),
        help="Optional id type filter.",
    )
    query_parser.add_argument(
        "--limit",
        type=int,
        default=20,
        help="Maximum number of matches to print. Default: 20",
    )
    query_parser.set_defaults(func=cmd_query)

    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    if not is_windows():
        raise ScanError("This tool currently supports Windows only.")

    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ScanError as exc:
        print(f"error={exc}", file=sys.stderr)
        raise SystemExit(2)
