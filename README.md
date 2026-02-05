# fund-docgen (Vercel)

这是一个「前端表单 + Flask API + docx 模板替换 + zip 下载」的小工具，已调整为适配 Vercel 的无状态运行环境：

- 静态页面：`public/index.html`
- API：`app.py`（Flask）
- docx 模板：`templates/<产品类型>/*.docx`

## 本地运行

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

然后访问：http://127.0.0.1:5000

## 在 Vercel 部署（通过 GitHub）

1. 把本项目推到 GitHub
2. 在 Vercel 里 Import Git Repository
3. 直接 Deploy（无需 build command）

> 提示：Vercel 建议把静态资源放到 `public/`，并且函数环境的磁盘是临时的，所以本项目改成 **/api/generate 直接返回 zip**。

另外，仓库包含：
- `.python-version`（固定 Python 3.12，避免构建日志里的版本探测提示）；
- `vercel.json` 使用 `functions + routes`（避免 `builds` 配置导致的 Vercel 警告）。

## 常见问题：Vercel 部署后 404

404 常见原因通常有两类：

1. **入口未声明**：仓库根目录有 `app.py`，但没有在 `vercel.json` 里声明 Python Function。
2. **路由顺序不当**：把所有请求都强制转发到 `app.py`，导致 `public/` 下的静态文件（如 `index.html`、`/static/*`）无法由 Vercel 文件系统优先命中。

本仓库现在的解决方案：
- `functions.app.py.runtime = python3.12`，明确函数入口；
- `routes` 第一条使用 `{"handle": "filesystem"}`，让 Vercel 先匹配 `public/` 静态资源；
- 未命中静态文件时，再由 `/(.*) -> /app.py` 兜底，确保 `/api/*` 与动态路由都进入 Flask。

如果你仍看到 404，请优先检查：
- Vercel Project 的 **Root Directory** 是否指向当前仓库根目录；
- 部署日志里是否出现 `app.py` 被识别为 Python Function；
- 访问路径是否正确（首页 `/`、接口 `/api/...`）。

