# FileFormat

Word文档格式化工具，支持自定义样式和格式设置。

## 功能特点

- 支持多级标题格式设置（最多6级）
- 自定义字体和字号
- 段落格式设置（行间距、缩进）
- 保持原有表格样式
- 保持原有图片格式
- 智能段落编号

## 项目结构

```plaintext
FileFormat/
├── src/
│   ├── document/        # 文档处理模块
│   ├── utils/          # 工具模块
│   └── main.py         # 主程序
├── config.ini          # 配置文件
└── README.md
```

## 使用说明

1. 准备配置文件：
   ```ini
   [Typography]
   body_font = 仿宋
   body_size = 18
   line_spacing = 1.5
   ```

2. 运行程序：
   ```
   python src/main.py --input 源文件.docx --output 结果文件.docx
   ```

## 环境要求

- Python 3.7+
- Windows 操作系统
```

这次重构主要改进：
1. 将表格和图片处理逻辑分离
2. 提高代码的可维护性和复用性
3. 简化了配置文件的结构
4. 优化了文档说明

需要我详细解释任何部分吗？