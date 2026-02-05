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

## 常见问题：Vercel 部署后 404

如果仓库根目录只有 `app.py` 而没有 `vercel.json`，Vercel 可能无法正确识别 Flask 入口并把请求转发到 Python 函数，访问首页或 `/api/*` 会出现 404。

本仓库现在通过 `vercel.json` 显式声明：
- `app.py` 使用 `@vercel/python` 构建；
- `public/**` 作为静态资源；
- `/static/*` 指向 `public/static/*`；
- 其余路由统一转发到 `app.py`。

