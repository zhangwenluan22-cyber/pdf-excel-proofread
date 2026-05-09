# 部署到公网（Render）

## 1. 准备代码仓库
把以下文件推到 GitHub 仓库：

- `pdf_excel_proofread_server.py`
- `pdf_excel_proofread_web.html`

## 2. 在 Render 创建 Web Service
1. 打开 Render，点击 `New +` -> `Web Service`
2. 连接你的 GitHub 仓库
3. 配置：
   - Runtime: `Python 3`
   - Build Command: 留空
   - Start Command:
     ```bash
     python3 pdf_excel_proofread_server.py
     ```

## 3. 环境变量（可选）
Render 会自动注入 `PORT`。  
本项目已支持：

- `HOST`（默认 `0.0.0.0`）
- `PORT`（默认 `8765`，在 Render 上会被覆盖）

## 4. 部署后访问
Render 会给你一个公网地址，例如：

`https://your-app-name.onrender.com`

把这个地址发给别人即可访问。

## 5. 注意事项
1. 当前没有登录/权限控制，任何拿到链接的人都能上传文件。
2. 如涉及敏感 PDF/Excel，建议后续加账号系统或私有部署。
3. 免费套餐可能休眠，首次打开会有冷启动等待。
