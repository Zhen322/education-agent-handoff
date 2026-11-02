#!/usr/bin/env python3
"""Restore the sanitized historical D1/R2 snapshot to a NEW local state folder."""
import argparse
import hashlib
import json
import shutil
import sqlite3
import tempfile
from pathlib import Path

bundle = Path(__file__).resolve().parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--target-state", type=Path, default=bundle.parent / "xiaohe-app/.wrangler/state")
args = parser.parse_args()
target = args.target_state.resolve()
if target.exists():
    raise SystemExit(f"拒绝覆盖已有路径：{target}。请关闭服务，将旧 state 单独备份后改名，或指定全新 --target-state。")
manifest = json.loads((bundle / "manifest.json").read_text())
with sqlite3.connect((bundle / manifest["d1File"]).as_uri() + "?mode=ro", uri=True) as db:
    assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok", "D1 integrity check failed"
    assert not db.execute("PRAGMA foreign_key_check").fetchall(), "D1 foreign key check failed"
    for table, count in manifest["counts"].items():
        assert db.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0] == count, table
    assert db.execute("SELECT COUNT(*) FROM workspace_invites WHERE token_hash NOT LIKE 'REDACTED-HANDOFF-INVALID-%'").fetchone()[0] == 0
for obj in manifest["objects"]:
    data = (bundle / "r2-blobs" / obj["blobId"]).read_bytes()
    assert len(data) == obj["bytes"] and hashlib.sha256(data).hexdigest() == obj["sha256"], "R2 content check failed"
target.parent.mkdir(parents=True, exist_ok=True)
staging = Path(tempfile.mkdtemp(prefix="handoff-restore-", dir=target.parent))
try:
    d1 = staging / "v3/d1/miniflare-D1DatabaseObject"
    r2 = staging / "v3/r2/miniflare-R2BucketObject"
    blobs = staging / "v3/r2/site-creator-r2/blobs"
    for directory in (d1, r2, blobs):
        directory.mkdir(parents=True)
    shutil.copy2(bundle / manifest["d1File"], d1 / manifest["d1RuntimeName"])
    shutil.copy2(bundle / manifest["r2File"], r2 / manifest["r2RuntimeName"])
    for obj in manifest["objects"]:
        shutil.copy2(bundle / "r2-blobs" / obj["blobId"], blobs / obj["blobId"])
    # Rename only into a new directory. Existing user data is never removed.
    if target.exists():
        raise RuntimeError("恢复过程中目标路径被创建；已停止，不覆盖。")
    staging.rename(target)
except BaseException:
    shutil.rmtree(staging, ignore_errors=True)
    raise
print(f"历史本地数据已恢复：{target}")
print("已含迁移 0000—0002；不要再次运行这三份建表迁移。此数据不是生产数据库或手机档案。")
