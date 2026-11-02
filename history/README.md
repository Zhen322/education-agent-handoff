# 教育项目历史源码

这两个 Git bundle 保留原工程可达提交历史、分支和文件内容。为公开交接，提交作者和提交者的邮箱统一为 GitHub noreply 地址，因此提交 SHA 与原记录不同；主分支文件树、提交数量已与原仓库核对一致。原始历史仍保存在交接人的本机，不放入公开仓库。

在仓库根目录执行，可恢复为两个独立历史仓库（另建文件夹，勿覆盖当前交接工程）：

```sh
git clone history/xiaohe-app.bundle ../xiaohe-app-history
git clone history/xiaohe-mobile.bundle ../xiaohe-mobile-history
```

最新可运行交接版本以本仓库的 `xiaohe-app/` 和 `xiaohe-mobile/` 为准；其中另加了新电脑入口、恢复说明和交接验证记录。
