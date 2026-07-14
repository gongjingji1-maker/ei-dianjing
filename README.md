# EI期刊智能筛选工具

一个集成了 EI 期刊列表筛选、元数据分析的 Web 工具，支持多维条件组合筛选和结果导出。

## 功能特性

- **多维筛选**：支持 Source title、Source type、Publisher、Country/Region、Subject 等多维度组合筛选
- **预设方案**：内置多种常用筛选预设（计算机与AI方向、能源工程方向等）
- **Web界面**：基于 Flask 的交互式筛选界面，支持勾选/下拉多选
- **数据统计**：自动统计出版商、国家、语言分布，支持出版商规模分级筛选
- **Excel导出**：筛选结果导出为带统计摘要的多 Sheet Excel 文件
- **期刊跳转**：点击期刊标题可直接搜索对应出版商官网

## 项目结构

```
EI期刊筛选工具/
├── ei_toolkit.py              # 核心筛选引擎
├── create_repo.py             # GitHub仓库创建脚本（需配置Token）
├── web/
│   ├── app.py                 # Flask Web应用
│   └── templates/
│       └── index.html         # 主页面模板
├── presets/                   # 筛选预设配置
│   ├── cs_ai_focus.json       # 计算机与AI方向
│   ├── energy_engineering.json # 能源工程方向
│   └── default_english.json   # 英文期刊默认方案
├── data/
│   ├── CPXSourceList_072026.xlsx  # EI原始数据（需定期更新）
│   └── metadata.json          # 筛选项元数据（自动生成）
└── README.md                  # 本文件
```

## 快速开始

### 1. 环境要求
- Python 3.8+
- pip 包管理器

### 2. 安装依赖
```bash
pip install flask openpyxl pandas
```

### 3. 启动Web界面
```bash
cd web
python app.py
```
打开浏览器访问：`http://localhost:5000`

### 4. 命令行使用
```bash
# 使用预设筛选
python ei_toolkit.py filter CPXSourceList_072026.xlsx --preset presets/cs_ai_focus.json

# 交互式筛选
python ei_toolkit.py
```

## 数据更新

当 EI 发布新的 CPXSourceList 时：
1. 下载最新 Excel 文件到 `data/` 目录
2. 运行 `python ei_toolkit.py metadata` 更新元数据
3. 重启 Web 服务即可使用最新数据

## 筛选预设配置

预设文件为 JSON 格式，支持以下操作：
- `include`: 包含匹配（精确或模糊）
- `exclude`: 排除匹配
- `regex`: 正则表达式匹配

示例 `presets/cs_ai_focus.json`：
```json
{
  "name": "计算机与AI方向",
  "description": "筛选计算机科学、人工智能、机器学习相关期刊",
  "filters": {
    "subject": {
      "include": ["COMPUTER SCIENCE", "ARTIFICIAL INTELLIGENCE", "MACHINE LEARNING"]
    },
    "language": {
      "exclude": ["CHINESE"]
    }
  }
}
```

## 注意事项

- 原始 EI 数据文件较大（约 5.4MB），GitHub 网页上传可能受限，建议使用 Git 命令行推送
- `create_repo.py` 中的 GitHub Token 请通过环境变量传入，不要硬编码
- 翻译缓存文件 `title_translations_cache.json` 会自动生成，无需手动管理

## 技术栈

- **后端**：Python + Flask
- **数据处理**：pandas, openpyxl
- **前端**：原生 HTML + JavaScript + Tailwind CSS
- **数据格式**：Excel (.xlsx), JSON

## 许可证

MIT License

## 作者

gongjingji1-maker

---

> 本项目为 EI 期刊筛选的个人研究工具，数据来源于 EI Compendex 官方列表，仅供学术参考使用。
