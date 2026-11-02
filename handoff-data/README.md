# 本地历史数据快照与恢复

本目录来自原工程的本地 D1/R2，2026-11-02 通过 SQLite backup API 读取一致性快照。原数据库没有修改。它保存本地演示/验收历史，不是 Sites 生产数据库，也不是手机/模拟器的 SQLite 档案。

- `d1-sanitized.sqlite`：4 个业务表，13 工作区、4 上传登记、12 成员、6 邀请。
- `d1-sanitized.sql`：同一内容的可恢复 SQL；已在内存 SQLite 重建并核对表数量/完整性。
- `r2-metadata.sqlite`、`r2-blobs/`：4 个对象，全部为相同 68 字节、1×1 PNG 测试占位图，与上传登记对应，不包含真实作业原稿。
- `manifest.json`：数量、文件映射、占位图摘要和脱敏说明。
- `restore-local.py`：校验后恢复到全新本地 state，拒绝覆盖已有路径。

6 个邀请原本全部已接受/已撤回，没有待用邀请。本副本仍将 token_hash 全部替换为明确无效的非哈希标记，并使用 secure_delete/VACUUM 清除旧字段残留；没有登录 cookie、session 或有效邀请凭证。为公开交接，全部 actor.name / membership.name 显示名替换为合成演示标签，ID 和关联关系保留；工作区/作答/题库内容保留以便验收追溯。成员/邀请邮箱均为 sites.test / example.test，13 个工作区都包含演示标记，其用途为历史调试，不应误作真实效果数据。

原数据库只执行到旧模式，缺少第三个迁移的 uploaded_files.history_epoch。本副本已按 `0002_history_upload_fence.sql` 增加默认 0 列；同时去掉可再生成的 `_cf_METADATA` 内部表。这些改变只发生在导出副本。

## 使用交接版本恢复

先关闭所有 Web/Wrangler 服务；在仓库根目录：

```sh
python3 handoff-data/restore-local.py
```

目标是 `xiaohe-app/.wrangler/state`。若已经存在，脚本停止；先单独备份并改名旧目录再重试，勿直接删除。恢复后无需再次运行 0000—0002 迁移，安装依赖、构建并启动 Web 即可。

测试恢复可指定新的临时目录：

```sh
python3 handoff-data/restore-local.py --target-state /tmp/xiaohe-handoff-test-state
```

恢复目录结构与当前 package-lock 中的 Miniflare/Wrangler 对应；升级这些依赖前先保留原快照并验证迁移。普通 SQL 也可用 SQLite 恢复到新数据库，例如 `sqlite3 new.sqlite < handoff-data/d1-sanitized.sql`。SQL 内含 BEGIN/COMMIT，作为 SQLite 恢复用途，不应直接当作远程 Cloudflare D1 自动迁移脚本。

不要把这些历史身份和演示数据直接导入正式生产数据库。生产数据与资源权限需要另外从原平台交接；此目录没有代替账号交接。
