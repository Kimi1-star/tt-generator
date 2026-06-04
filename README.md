# TT 文件生成器

自动生成 Commercial Invoice / Packing List / Shipment Advice（TT 文件）的 Web 工具。

## 功能

- 上传合同 PDF + 明细/ERP Excel
- 自动提取合同号、买方、品名、单价、贸易条款
- 自动解析货物明细（ERP 格式 `.xlsx` 或人工明细 `.xls`）
- 手动填写日期、船名、港口、运费等
- 一键生成含 CI / PL / SA 三个 Sheet 的 Excel 文件

## 本地运行

```bash
pip install -r requirements.txt
python app.py
# 访问 http://localhost:5000
```

## 部署到 Render

1. 将本目录推送到 GitHub 仓库
2. 在 [render.com](https://render.com) 创建 New Web Service
3. 连接 GitHub 仓库，Render 会自动使用 `render.yaml` 配置

## 文件说明

| 文件 | 说明 |
|------|------|
| `app.py` | Flask Web 应用入口 |
| `parser.py` | PDF / Excel 解析 |
| `generator.py` | Excel TT 文件生成 |
| `templates/index.html` | 前端页面 |
