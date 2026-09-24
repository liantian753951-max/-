# 长图转报价表

在 Codex 对话中识别需求长图，生成带制作图的报价表，并自动保存 XLSX 到 Windows 桌面。

## 工作方式

- Codex 负责识别项目名称、物件名称和图片位置。
- Python 固定程序负责裁图、套用缓存模板、嵌入图片与校验。
- Google Drive 连接器负责导入 Google Sheets。
- 工时、各项合计和总价留空；模板单价保留。

这不是独立客户端，不包含独立 OCR 服务、API Key 或 Google 登录凭据。

## 本地构建

安装 Python 依赖：`python -m pip install -r requirements.txt`。

准备 JSON：

```json
{"project":"项目名","supplier":"供应商","expected_count":1,"items":[{"name":"物件名称","box":[100,800,1300,1800]}]}
```

坐标以 EXIF 校正后的原图像素为准。

```sh
python skills/quote-image-to-sheet/scripts/build_quote.py build image.jpg input.json output
python skills/quote-image-to-sheet/scripts/save_desktop.py output/result.json
```

第二条命令仅用于 Windows。Google Sheets 导入由 Codex 内已连接的 Google Drive 完成。

## 文件

- `.codex-plugin/plugin.json`：插件清单。
- `skills/quote-image-to-sheet/SKILL.md`：对话工作流程。
- `scripts/build_quote.py`：固定转换程序。
- `scripts/save_desktop.py`：保存到实际桌面目录，同名文件不覆盖。
- `assets/template.xlsx`：去除旧项目内容和图片的模板缓存。

原始需求图、生成的客户报价表、临时文件和凭据不纳入仓库。
