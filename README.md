# archppt

## 翻译工作流（Translate PPT to zh-CN）

此仓库包含一个 GitHub Actions 工作流，可将 `01计算机设计导论.pptx` 中所有"非中文"的文本翻译为简体中文，并生成 `01计算机设计导论_zh-CN.pptx`。

### 使用步骤

1. 在仓库设置中添加一个 Repository Secret：
   - 名称：`DEEPL_API_KEY`
   - 值：你的 [DeepL API](https://www.deepl.com/docs-api/) 密钥

2. 触发工作流：
   - 在 GitHub 仓库页面，进入 **Actions** 标签页，选择 **Translate PPT to zh-CN** 工作流。
   - 点击 **Run workflow**，确认/修改以下参数：
     - `source_path`（默认 `01计算机设计导论.pptx`）
     - `output_path`（默认 `01计算机设计导论_zh-CN.pptx`）
     - `enable_ocr`（是否启用图片 OCR，默认 `false`）

3. 运行完成后，工作流会把翻译后的 `output_path` 文件提交到当前分支。

### 说明

- 翻译服务使用 DeepL API，目标语言为简体中文（ZH）。若文本原本为中文，将不会被翻译。
- 启用 OCR 时，脚本会尝试识别图片中的文字，并将识别结果及其翻译写入图片的 `alternative_text` 以供参考。复杂公式/图表可能无法完整识别。
- 为了保证稳定性，脚本在翻译时可能会简化文本框的格式（例如 run 级格式）。若需保留精细格式，请在合并后根据需要调整脚本。