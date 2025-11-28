# Gamma API 演示文稿生成工具

这个项目用于调用 Gamma API，将文本内容或 PDF 文件转换为演示文稿（PPT），并导出为 PDF 格式保存到 `output` 目录。

## 功能特性

- 支持从 `.docx` 文件提取文本内容和图片链接
- 支持从 `.pdf` 文件提取文本内容（Gamma API 不支持直接 PDF 输入）
- 调用 Gamma API 生成演示文稿
- 自动等待生成完成
- **智能导出功能**：
  - 优先尝试通过 API 导出 PDF/PPTX
  - 如果 API 失败，自动使用浏览器自动化导出
  - 支持导出为 PDF 和 PPTX 格式
- 生成记录管理：自动记录已生成的文档，避免重复生成
- 导出并下载到 `output` 目录

## 安装依赖

### 使用 Conda（推荐）

```bash
# 创建并激活 conda 环境
conda env create -f environment.yml
conda activate gamma_local
```

### 使用 Pip（备选）

```bash
pip install -r requirements.txt
```

## 配置

1. 复制 `env_example.txt` 为 `.env`
2. 在 `.env` 文件中设置您的 Gamma API Key：

```
GAMMA_API_KEY=your_gamma_api_key_here
```

您可以从 [Gamma 开发者平台](https://developers.gamma.app) 获取 API Key。

## 使用方法

1. 确保已激活 conda 环境（如果使用 conda）：
   ```bash
   conda activate gamma_local
   ```

2. 将需要处理的文件（`.docx` 或 `.pdf`）放入 `dataset` 目录

3. 运行主程序：
   ```bash
   # 正常处理（自动跳过已生成的文档）
   python main.py
   
   # 强制重新生成（即使已存在记录）
   python main.py --force
   
   # 查看所有生成记录
   python main.py --list
   
   # 使用浏览器自动化导出（如果 API 导出失败）
   python main.py --browser
   
   # 组合使用：强制重新生成并使用浏览器导出
   python main.py --force --browser
   
   # 使用无头模式（后台运行，不显示浏览器窗口）
   # 在 .env 文件中设置: BROWSER_HEADLESS=true
   python main.py --browser
   ```

程序会：
- 自动查找 `dataset` 目录下的文件
- **检查是否已生成过**（通过文件路径和哈希值）
- 如果已生成，直接使用已有记录，避免重复生成
- 优先处理包含"原始邮件数据"的 `.docx` 文件
- 提取文本内容和图片链接
- 调用 Gamma API 生成演示文稿
- 等待生成完成
- 记录生成信息到 `generation_records.json`
- **智能导出 PDF**：
  - 优先尝试通过 API 导出
  - 如果 API 失败，自动使用浏览器自动化导出（如果安装了 Selenium）
  - 支持使用 `--browser` 参数强制使用浏览器自动化

## 项目结构

```
gamma_local/
├── dataset/              # 输入文件目录
│   ├── *.docx           # Word 文档
│   └── *.pdf            # PDF 文档
├── output/              # 输出目录（自动创建）
│   └── *.pdf            # 生成的 PDF 文件
├── generation_records.json  # 生成记录文件（自动创建）
├── main.py              # 主程序
├── environment.yml       # Conda 环境配置（推荐）
├── requirements.txt     # Pip 依赖包（备选）
├── env_example.txt      # 环境变量配置示例
├── .env                 # 环境变量配置（需要创建）
└── README.md            # 说明文档
```

## API 参数说明

程序严格按照 [Gamma API 文档](https://developers.gamma.app/docs/generate-api-parameters-explained) 的规范实现：

- `inputText`: 输入文本内容
- `textMode`: 设置为 `"generate"`
- `format`: 设置为 `"presentation"`
- `themeId`: 主题 ID（默认 `"Oasis"`）
- `cardSplit`: 设置为 `"auto"` 自动分割幻灯片
- `exportAs`: 设置为 `"pdf"` 导出为 PDF
- `textOptions`: 文本选项（详细程度、语气、目标受众、语言）
- `imageOptions`: 图像选项（来源、模型、风格）
- `cardOptions`: 卡片选项（尺寸、页眉页脚）

## 生成记录功能

程序会自动记录所有生成的文档信息，包括：
- 文件路径和哈希值（用于唯一标识）
- 生成任务 ID (generationId)
- Gamma 演示文稿 URL
- 生成状态和时间戳
- PDF 下载状态

**功能特性：**
- ✅ **自动去重**: 处理文件前会检查是否已生成过，避免重复调用 API
- ✅ **记录查询**: 使用 `python main.py --list` 查看所有生成记录
- ✅ **强制重新生成**: 使用 `python main.py --force` 强制重新生成
- ✅ **记录持久化**: 所有记录保存在 `generation_records.json` 文件中

**记录文件格式：**
```json
{
  "generation_id": {
    "file_path": "/path/to/file.docx",
    "file_name": "file.docx",
    "file_hash": "md5_hash",
    "generation_id": "xxx",
    "gamma_url": "https://gamma.app/docs/xxx",
    "status": "completed",
    "created_at": "2025-11-24T10:00:00",
    "updated_at": "2025-11-24T10:05:00",
    "pdf_downloaded": true,
    "pdf_path": "output/file_gamma_presentation.pdf"
  }
}
```

## 注意事项

1. **API 限制**: Gamma API 的 `inputText` 参数有最大令牌限制（约 100,000 tokens），程序会自动截断过长的文本
2. **PDF 输入**: Gamma API 不支持直接 PDF 输入，程序会先提取 PDF 的文本内容
3. **生成时间**: 演示文稿生成可能需要一些时间，程序会定期检查状态直到完成
4. **API Key**: 请妥善保管您的 API Key，不要提交到版本控制系统
5. **记录文件**: `generation_records.json` 包含生成记录，已添加到 `.gitignore`，不会被提交到版本控制
6. **浏览器自动化**: 如果 API 导出失败，程序会自动尝试使用 Selenium 进行浏览器自动化导出。需要安装 Chrome 浏览器。
7. **导出格式**: 根据[官方文档](https://developers.gamma.app/docs/getting-started)，支持导出为 PDF 和 PPTX 格式

## 错误处理

- 如果 API 调用失败，程序会显示错误信息
- 如果生成超时（默认 5 分钟），程序会停止等待
- 如果文件读取失败，程序会跳过该文件

## 许可证

MIT License

