# Office Preview

[![Release](https://img.shields.io/github/v/release/dunnice/obsidian-office-preview?style=flat-square)](https://github.com/dunnice/obsidian-office-preview/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

**Office Preview** is an Obsidian plugin for native, smooth, and privacy-first previewing of Word (`.docx`), Excel (`.xlsx`, `.xls`, `.csv`), PowerPoint (`.pptx`, `.ppt`), JSON, SQL, and YAML files directly within Obsidian.

Zero external software dependencies, zero network requests, 100% locally rendered in Obsidian.

---

## English

### ✨ Key Features

- 📄 **Word Documents (`.docx`) Preview**
  - **Fluid Reader Layout**: Adaptive viewport centering with smooth reading flow.
  - **Preserved Formatting**: Keeps original typography, headings, font sizes, and paragraph structure.
  - **🔍 Keyword Search (`Cmd+F` / `Ctrl+F`)**: Real-time highlighting with smooth auto-scroll to match.
  - **Modification Timestamp**: Displays file last modified time and size unobtrusively at the top.

- 📊 **Excel Spreadsheets (`.xlsx`, `.xls`, `.csv`) Preview**
  - **Native Responsive DOM Table Engine**: Clean, performant, and crisp table rendering.
  - **Multi-Sheet Tabs**: Instant switching between workbook sheets with responsive tabs.
  - **🔍 Cell Search & Highlight**: Search matching cell contents with count indicators and focus navigation.
  - **Sticky Headers & Line Numbers**: Sticky column headers (A, B, C...) and row numbers with smooth scrolling.

- 📽️ **PowerPoint Presentations (`.pptx`, `.ppt`) Preview**
  - **Fit-Width Presentation Layout**: Slides fit viewport width automatically without horizontal scrollbars.
  - **Multi-Slide Cascading Stream**: Renders every slide sequentially and reliably.
  - **Legacy `.ppt` Support**: Friendly detection and one-click prompt to open with default system app.
  - **🖥️ Fullscreen Presentation**: Enter distraction-free fullscreen presentation mode anytime.
  - **🔍 In-Slide Full-Text Search**: Full-text search across all slide text boxes with keyword jumps.

- 💻 **Structured Code & Config Files (`.json`, `.sql`, `.yaml`, `.yml`) Preview**
  - **{ } JSON Formatting**: Auto indentation and beautification, with quick toggle between formatted and compact view.
  - **↩️ Smart Word Wrap**: Automatic line breaking according to viewport width to prevent horizontal overflow.
  - **🗂️ 7:3 Golden Ratio Tree Outline**: Displays hierarchical keys for JSON/YAML on the right side with collapsible parent nodes and instant code line navigation.
  - **🎨 Syntax Highlighting**: Prism.js syntax coloring adapted for both dark and light Obsidian themes.
  - **Line Numbers & One-Click Copy**: Clean gutter numbers and one-click clipboard copying.
  - **🔍 Full-Text Search**: `Cmd+F` search with hit counting and Enter navigation.

- 🔒 **100% Local & Privacy-First**
  - Runs entirely within Obsidian's local Electron process. No internet connection required, keeping your confidential notes and files completely secure.

---

## 中文说明

### ✨ 核心特性

- 📄 **Word 文档 (`.docx`) 预览**
  - **自适应优雅排版**：视口居中呈现，排版顺畅自然。
  - **保留字号与格式**：完美继承原始文档的标题层级、字体与段落样式。
  - **🔍 关键词搜索与定位**：支持快捷键 `Cmd+F` / `Ctrl+F` 实时检索高亮与居中定位。
  - **修改时间与体积**：文件名旁低调呈现修改时间与文件大小。

- 📊 **Excel 表格 (`.xlsx` / `.xls` / `.csv`) 预览**
  - **原生响应式 DOM 表格引擎**：告别遮挡，100% 高清表格呈现。
  - **多 Sheet 切换**：顶部选项卡毫秒级流畅切换 Sheet。
  - **🔍 单元格搜寻与定位**：支持全局检索匹配单元格并定位跳转。
  - **固定行列头**：列标与行号吸顶/吸左，支持双向平滑滚动。

- 📽️ **PowerPoint 演示文稿 (`.pptx` / `.ppt`) 预览**
  - **自适应视口宽度 (Fit Width)**：幻灯片自适应视口宽度，彻底消除横向滚动条。
  - **全页连续流式平铺**：逐页解析并平铺展示所有幻灯片。
  - **🖥️ 一键全屏演示**：沉浸式深黑背景全屏放映。
  - **🔍 幻灯片全文检索**：全局搜寻幻灯片内文本并定位。

- 💻 **结构化代码与配置文件 (`.json` / `.sql` / `.yaml` / `.yml`) 预览**
  - **{ } JSON 智能格式化**：自动缩进美化排版，支持一键切换【✨ 格式化 / 原始紧凑】。
  - **↩️ 智能自动换行 (Word Wrap)**：内容根据视口宽度自动折行，杜绝横向滚动条。
  - **🗂️ 7:3 黄金比例树形大纲**：右侧展示 JSON/YAML 层级键名，支持逐级折叠与展开，点击快速定位行。
  - **🎨 语法高亮**：完整支持 JSON、SQL、YAML 语法着色（适配暗黑/明亮主题）。
  - **对齐行号与一键复制**：独立行号列与顶部一键复制代码全文。

- 🔒 **全本地 & 私密安全**
  - 纯本地运行，不产生任何外部网络请求，保障数据隐私安全。

---

## 🛠️ 安装与使用 (Installation)

### 方式一：社区插件市场安装（推荐）
在 Obsidian 官方插件市场上线后，打开 Obsidian **设置** -> **社区插件** -> **浏览**，搜索 **`Office Preview`** 并安装启用。

### 方式二：手动安装
1. 前往 [Releases](https://github.com/dunnice/obsidian-office-preview/releases) 下载最新的 `main.js`、`styles.css` 和 `manifest.json`；
2. 将这三个文件放置在你的库目录：`.obsidian/plugins/office-preview/`；
3. 打开 Obsidian 设置并启用插件。

---

## 📄 开源协议 (License)

[MIT License](LICENSE)
