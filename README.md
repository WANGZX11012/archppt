# archppt

## 翻译工作流（Translate PPT to zh-CN，离线）

此仓库提供一个离线工作流，将 `01计算机设计导论.pptx` 中所有"非中文"的文本翻译为简体中文，并生成 `01计算机设计导论_zh-CN.pptx`。不需要任何 API 密钥；默认使用 Argos Translate（离线），可选启用图片 OCR（pytesseract）。

### 使用步骤

1. 合并本 PR 后，仓库会包含以下文件：
   - `scripts/translate_ppt.py`
   - `.github/workflows/translate-ppt.yml`

2. 触发工作流：
   - 在 GitHub 仓库页面进入 **Actions**，选择 **Translate PPT to zh-CN (offline)**。
   - 点击 **Run workflow**，确认或修改参数：
     - `source_path`（默认 `01计算机设计导论.pptx`）
     - `output_path`（默认 `01计算机设计导论_zh-CN.pptx`）
     - `enable_ocr`（是否启用图片 OCR，默认 `false`）

3. 运行完成后，工作流会将翻译后的 `output_path` 文件提交到当前分支。

### 说明

- 若文本原本为中文，将不会被翻译；仅翻译非中文为简体中文。
- 首次运行会下载并安装 Argos Translate 对应语言模型（如 en→zh），下载完成后可离线使用。
- 启用 OCR 时，脚本会尽力识别图片中的文字，并将识别结果及其翻译写入该图片的 `alternative_text` 以供参考。复杂公式/图表可能无法完整识别。
- 直接设置文本可能影响段落内的精细格式（如不同 run 的样式）；该方案优先保障翻译正确性。