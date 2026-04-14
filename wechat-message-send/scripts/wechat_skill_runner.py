#!/usr/bin/env python3
"""Unified workflow wrapper for the WeChat id and sender tools."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import tempfile
from typing import Dict, Optional, Sequence

import wechat_id_tool
import wechat_sender


class WorkflowError(RuntimeError):
    pass


def skill_root() -> Path:
    current = Path(__file__).resolve().parent
    if current.name.lower() == "scripts":
        return current.parent
    return current


def data_root() -> Path:
    return skill_root() / "data"


def out_root() -> Path:
    return data_root() / "out"


def default_mapping_file() -> Path:
    return data_root() / "target_mappings.json"


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def load_mapping_file(path: Path) -> Dict[str, str]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise WorkflowError(f"Invalid mapping file structure: {path}")
    result: Dict[str, str] = {}
    for key, value in payload.items():
        if isinstance(key, str) and isinstance(value, str):
            result[key.strip()] = value.strip()
    return result


def remember_mapping(*, mapping_file: Path, target_id: str, keyword: str) -> Dict[str, str]:
    mapping = load_mapping_file(mapping_file)
    mapping[target_id.strip()] = keyword.strip()
    ensure_parent(mapping_file)
    mapping_file.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    return mapping


def resolve_message(message: Optional[str], message_file: Optional[Path]) -> str:
    if message_file:
        return message_file.read_text(encoding="utf-8").rstrip("\r\n")
    if message is not None:
        return message
    raise WorkflowError("Either --message or --message-file is required.")


def write_temp_message_file(text: str) -> Path:
    data_root().mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".txt",
        prefix="wechat-message-",
        encoding="utf-8",
        delete=False,
        dir=data_root(),
        ) as handle:
        handle.write(text)
        return Path(handle.name)


def write_temp_mapping_file(mapping: Dict[str, str]) -> Path:
    data_root().mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        prefix="wechat-mapping-",
        encoding="utf-8",
        delete=False,
        dir=data_root(),
    ) as handle:
        json.dump(mapping, handle, ensure_ascii=False, indent=2)
        return Path(handle.name)


def prepare_message_file(message: Optional[str], message_file: Optional[Path]) -> tuple[Path, bool]:
    if message_file is not None:
        return message_file, False
    return write_temp_message_file(resolve_message(message, None)), True


def cmd_scan(args: argparse.Namespace) -> int:
    argv = ["scan", "--output-dir", str(out_root())]
    if args.wechat_path:
        argv.extend(["--wechat-path", str(args.wechat_path)])
    if args.no_memory:
        argv.append("--no-memory")
    for root in args.root or []:
        argv.extend(["--root", str(root)])
    return wechat_id_tool.main(argv)


def cmd_find(args: argparse.Namespace) -> int:
    argv = [
        "query",
        args.keyword,
        "--output-dir",
        str(out_root()),
        "--limit",
        str(args.limit),
    ]
    if args.kind:
        argv.extend(["--kind", args.kind])
    if args.input:
        argv.extend(["--input", str(args.input)])
    return wechat_id_tool.main(argv)


def run_sender(argv: Sequence[str]) -> int:
    return wechat_sender.main(["send", *argv])


def cmd_remember(args: argparse.Namespace) -> int:
    mapping = remember_mapping(
        mapping_file=args.mapping_file,
        target_id=args.id,
        keyword=args.keyword,
    )
    print(f"mapping_file={args.mapping_file}")
    print(f"saved_id={args.id}")
    print(f"saved_keyword={mapping[args.id]}")
    return 0


def cmd_send_by_name(args: argparse.Namespace) -> int:
    message_file, should_cleanup_message = prepare_message_file(args.message, args.message_file)
    mapping = load_mapping_file(args.mapping_file)
    mapping[args.keyword.strip()] = args.keyword.strip()
    temp_mapping_file = write_temp_mapping_file(mapping)
    argv = [
        "--ids",
        args.keyword,
        "--mapping-file",
        str(temp_mapping_file),
        "--message-file",
        str(message_file),
    ]

    if args.wechat_path:
        argv.extend(["--wechat-path", str(args.wechat_path)])
    if args.no_enter:
        argv.append("--no-enter")
    if args.dry_run:
        argv.append("--dry-run")
    pick_index = getattr(args, "pick_index", None)
    if pick_index is not None:
        argv.extend(["--pick-index", str(pick_index)])
    debug_dir = getattr(args, "debug_dir", None)
    if debug_dir is not None:
        argv.extend(["--debug-dir", str(debug_dir)])
    try:
        return run_sender(argv)
    finally:
        temp_mapping_file.unlink(missing_ok=True)
        if should_cleanup_message:
            message_file.unlink(missing_ok=True)


def cmd_send_by_id(args: argparse.Namespace) -> int:
    message_file, should_cleanup_message = prepare_message_file(args.message, args.message_file)
    argv = [
        "--ids",
        *args.ids,
        "--mapping-file",
        str(args.mapping_file),
        "--message-file",
        str(message_file),
    ]
    if args.wechat_path:
        argv.extend(["--wechat-path", str(args.wechat_path)])
    if args.no_enter:
        argv.append("--no-enter")
    if args.dry_run:
        argv.append("--dry-run")
    pick_index = getattr(args, "pick_index", None)
    if pick_index is not None:
        argv.extend(["--pick-index", str(pick_index)])
    debug_dir = getattr(args, "debug_dir", None)
    if debug_dir is not None:
        argv.extend(["--debug-dir", str(debug_dir)])
    try:
        return run_sender(argv)
    finally:
        if should_cleanup_message:
            message_file.unlink(missing_ok=True)


def cmd_send_current(args: argparse.Namespace) -> int:
    message_file, should_cleanup_message = prepare_message_file(args.message, args.message_file)
    argv = [
        "--ids",
        *(args.ids or ["current-chat"]),
        "--current",
        "--message-file",
        str(message_file),
    ]
    if args.wechat_path:
        argv.extend(["--wechat-path", str(args.wechat_path)])
    if args.no_enter:
        argv.append("--no-enter")
    if args.dry_run:
        argv.append("--dry-run")
    try:
        return run_sender(argv)
    finally:
        if should_cleanup_message:
            message_file.unlink(missing_ok=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Unified workflow for WeChat skill tasks")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan local WeChat ids and chatrooms")
    scan_parser.add_argument("--wechat-path", type=Path)
    scan_parser.add_argument("--root", action="append", type=Path)
    scan_parser.add_argument("--no-memory", action="store_true")
    scan_parser.set_defaults(func=cmd_scan)

    find_parser = subparsers.add_parser("find", help="Search the latest WeChat scan by keyword")
    find_parser.add_argument("keyword")
    find_parser.add_argument("--kind", choices=("wxid", "chatroom"))
    find_parser.add_argument("--input", type=Path)
    find_parser.add_argument("--limit", type=int, default=20)
    find_parser.set_defaults(func=cmd_find)

    remember_parser = subparsers.add_parser("remember", help="Remember an id -> keyword mapping for stable sending")
    remember_parser.add_argument("--id", required=True)
    remember_parser.add_argument("--keyword", required=True)
    remember_parser.add_argument("--mapping-file", type=Path, default=default_mapping_file())
    remember_parser.set_defaults(func=cmd_remember)

    send_name_parser = subparsers.add_parser("send-by-name", help="Search a conversation by keyword and send")
    send_name_parser.add_argument("--keyword", required=True)
    send_name_parser.add_argument("--message")
    send_name_parser.add_argument("--message-file", type=Path)
    send_name_parser.add_argument("--wechat-path", type=Path)
    send_name_parser.add_argument("--mapping-file", type=Path, default=default_mapping_file())
    send_name_parser.add_argument("--no-enter", action="store_true")
    send_name_parser.add_argument("--dry-run", action="store_true")
    send_name_parser.add_argument(
        "--pick-index",
        type=int,
        default=None,
        help="When search returns multiple rows, pick the Nth OCR candidate (1-based).",
    )
    send_name_parser.add_argument(
        "--debug-dir",
        type=Path,
        default=None,
        help="Write search OCR JSON and region PNGs for troubleshooting.",
    )
    send_name_parser.set_defaults(func=cmd_send_by_name)

    send_id_parser = subparsers.add_parser("send-by-id", help="Send using one or more mapped target ids")
    send_id_parser.add_argument("--ids", nargs="+", required=True)
    send_id_parser.add_argument("--message")
    send_id_parser.add_argument("--message-file", type=Path)
    send_id_parser.add_argument("--wechat-path", type=Path)
    send_id_parser.add_argument("--mapping-file", type=Path, default=default_mapping_file())
    send_id_parser.add_argument("--no-enter", action="store_true")
    send_id_parser.add_argument("--dry-run", action="store_true")
    send_id_parser.add_argument(
        "--pick-index",
        type=int,
        default=None,
        help="Applies to each search-mode target: pick the Nth OCR candidate (1-based).",
    )
    send_id_parser.add_argument(
        "--debug-dir",
        type=Path,
        default=None,
        help="Write search OCR JSON and region PNGs for troubleshooting.",
    )
    send_id_parser.set_defaults(func=cmd_send_by_id)

    send_current_parser = subparsers.add_parser("send-current", help="Send to the currently open chat")
    send_current_parser.add_argument("--ids", nargs="*")
    send_current_parser.add_argument("--message")
    send_current_parser.add_argument("--message-file", type=Path)
    send_current_parser.add_argument("--wechat-path", type=Path)
    send_current_parser.add_argument("--no-enter", action="store_true")
    send_current_parser.add_argument("--dry-run", action="store_true")
    send_current_parser.set_defaults(func=cmd_send_current)

    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except WorkflowError as exc:
        print(f"error={exc}", file=sys.stderr)
        raise SystemExit(2)
