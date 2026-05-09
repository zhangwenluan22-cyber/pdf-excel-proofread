# PDF × Excel 人工校对台

一个用于 **PDF 预览 + Excel 层级校对** 的小工具。

## 功能
- 上传 Excel 与 PDF
- 按 `大条目 -> 中条目 -> 小条目` 层级展示内容
- 每条支持 `待校对 / 校对无误 / 有错误` 与备注
- 导出“有错误行（CSV）”
- 保留 `Excel原行号` 便于回原表定位

## 本地运行
```bash
cd /Users/user/Desktop/句型辞典html创建
python3 pdf_excel_proofread_server.py
```

浏览器打开：
`http://127.0.0.1:8765`

## 公网部署（Render）
本项目已包含 `render.yaml`，可直接部署：

1. 把项目推到 GitHub
2. 在 Render 选择该仓库创建 Web Service
3. Render 会读取 `render.yaml` 自动配置并启动

部署后会得到一个公网 URL，可分享给他人访问。

## 关键文件
- `pdf_excel_proofread_server.py`：后端服务
- `pdf_excel_proofread_web.html`：前端页面
- `render.yaml`：Render 部署配置
