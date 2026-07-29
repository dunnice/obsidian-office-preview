> **Obsidian Office Preview** 是一个用于在 Obsidian 内纯本地、高颜值的 Word (`.docx`) 和 Excel (`.xlsx`, `.xls`, `.csv`) 文件预览插件。零外部软件依赖、零网络请求，全本地 Electron 进程渲染。

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
