> **Obsidian Office Preview** 是一个用于在 Obsidian 内纯本地、高颜值的 Word (`.docx`)、Excel (`.xlsx`, `.xls`, `.csv`) 与 PowerPoint (`.pptx`, `.ppt`) 文件预览插件。零外部软件依赖、零网络请求，全本地 Electron 进程渲染。

---

## ✨ 核心特性

- 📄 **Word 文档 (`.docx`) 预览**
  - **优雅阅读器排版**：自适应视口居中呈现，排版顺畅自然。
  - **忠实保留字号与格式**：完美继承原始文档的标题层级、字体与段落样式。
  - **🔍 关键词搜索与定位**：支持 `Cmd+F` / `Ctrl+F` 唤起搜索面板，实时文本高亮与平滑滚动居中定位。
  - **修改时间展示**：顶部文件名右侧浅色低调展示最后修改时间。

- 📊 **Excel 表格 (`.xlsx` / `.xls` / `.csv`) 预览**
  - **原生响应式 DOM 表格引擎**：彻底告别可输入框 overlay 遮挡问题，100% 清晰渲染。
  - **多 Sheet 页签切换**：顶部美观的 Sheet 选项卡，点击即可毫秒级无缝切换目标 Sheet。
  - **🔍 单元格搜索与定位**：支持 `Cmd+F` 全局搜寻匹配单元格，匹配计数与平滑聚焦定位。
  - **固定表头与行号**：列标号 (A, B, C...) 自动吸顶，行号 (1, 2, 3...) 自动左侧固定，支持 Mac 硬件加速双向平滑滚动。

- 📽️ **PowerPoint 演示文稿 (`.pptx` / `.ppt`) 预览**
  - **自适应视口铺满 (Fit Width)**：幻灯片宽度自动自适应视口，**彻底消除横向滚动条**，仅保留纵向流畅滚动。
  - **逐页独立安全渲染**：避免单页图元异常中断全局，多页顺畅垂直向下铺展。
  - **旧版 `.ppt` 智能检测**：友好引导并提供一键调用外部应用打开。
  - **🖥️ 一键全屏演示**：点击全屏按钮即可进入沉浸式全黑背景演示模式。
  - **🔍 幻灯片全文检索**：支持快捷键 `Cmd+F` 检索所有幻灯片文本内容，实时高亮与平滑定位跳转。

- 💻 **结构化代码与配置文件 (`.json` / `.sql` / `.yaml` / `.yml`) 预览**
  - **{ } JSON 智能格式化**：自动美化排版与缩进，支持一键切换【✨ 格式化 / 紧凑模式】。
  - **↩️ 根据内容智能自动换行 (Word Wrap)**：内容根据视口宽度自动折行，**杜绝横向滚动条**，顶部支持一键开关【↩️ 换行】。
  - **🗂️ 7:3 黄金比例树形大纲**：在右侧窗展示 JSON/YAML 层级键名，**支持父级节点逐层折叠/展开及顶部【🔼 全部折叠 / 🔽 全部展开】**，点击任意节点平滑跳转定位代码行！
  - **🎨 精美语法高亮**：完整支持 JSON、SQL、YAML 语法着色（适配 Obsidian 深色/浅色主题）。
  - **🔢 对齐行号与一键复制**：左侧独立行号列，顶部工具栏支持一键快速复制代码全文。
  - **🔍 关键词搜索定位**：支持 `Cmd+F` 检索代码内容，实时高亮与回车跳转定位。

- 🔒 **全本地 & 高性能**
  - 跑在 Obsidian/Electron 渲染进程中，对 Apple Silicon (M1/M2/M3) 芯片极佳适配，绝不发起网络请求，数据 100% 私密安全。

---

## 🛠️ 安装说明

### 方式一：复制构建产物安装
1. 下载或编译项目生成的 `main.js`、`styles.css` 和 `manifest.json`。
2. 将这三个文件拷贝至你的 Vault 插件目录：
   `YOUR_VAULT/.obsidian/plugins/obsidian-office-preview/`
3. 打开 Obsidian **设置** -> **第三方插件**，点击刷新并**启用 Office Preview** 插件。

### 方式二：开发环境软链接 (macOS / Linux)
在终端运行：
```bash
ln -s /path/to/obsidian-office-preview "YOUR_VAULT/.obsidian/plugins/obsidian-office-preview"
```

---

## 💻 开发者指南

如果你想要修改代码或进行二次开发：

```bash
# 1. 克隆代码库
git clone git@github.com:dunnice/obsidian-office-preview.git
cd obsidian-office-preview

# 2. 安装依赖
npm install

# 3. 编译构建
npm run build

# 4. 增量开发模式 (自动监听修改)
npm run dev
```

---

## 📄 开源协议

[MIT License](LICENSE)
