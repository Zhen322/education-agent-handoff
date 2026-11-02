#!/usr/bin/env python3
"""Verify every delivered file against the handoff inventory using the standard library."""
import hashlib
import json
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]
manifest = json.loads((root / "交接文档" / "文件清单.json").read_text(encoding="utf-8"))
failed = []
for entry in manifest["files"]:
    path = root / entry["path"]
    if not path.is_file() or path.is_symlink():
        failed.append(entry["path"] + "：缺失或不是普通文件")
        continue
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    if path.stat().st_size != entry["bytes"] or digest != entry["sha256"]:
        failed.append(entry["path"] + "：大小或 SHA-256 不符")
if failed:
    print("校验未通过：\n" + "\n".join(failed))
    sys.exit(1)
print(f"校验通过：{len(manifest['files'])} 个交付文件的大小与 SHA-256 全部一致。")
print("文件清单自身不参与自校验；后续修改代码或新增数据后不再对应本次交付快照。")
