import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import Prism from 'prismjs';
import 'prismjs/components/prism-json';
import 'prismjs/components/prism-sql';
import 'prismjs/components/prism-yaml';

export const VIEW_TYPE_CODE = 'office-preview-code-view';

interface OutlineNode {
    key: string;
    path: string;
    type: 'object' | 'array' | 'value';
    preview?: string;
    level: number;
    lineNumber: number;
}

export class CodeView extends FileView {
    private currentFile: TFile | null = null;
    private rawContent: string = '';
    private formattedContent: string = '';
    private isJsonFormatted: boolean = true;
    private language: string = 'plaintext';
    private isOutlineVisible: boolean = true;

    // 搜索状态管理
    private searchMatches: HTMLElement[] = [];
    private currentMatchIndex: number = -1;
    private isSearchVisible: boolean = false;
    private searchInputEl: HTMLInputElement | null = null;
    private searchCountEl: HTMLElement | null = null;
    private searchContainerEl: HTMLElement | null = null;

    constructor(leaf: WorkspaceLeaf) {
        super(leaf);
    }

    getViewType(): string {
        return VIEW_TYPE_CODE;
    }

    getDisplayText(): string {
        return this.currentFile ? this.currentFile.name : 'Code Preview';
    }

    getIcon(): string {
        const ext = this.currentFile?.extension?.toLowerCase();
        switch (ext) {
            case 'json':
                return 'code-glyph';
            case 'sql':
                return 'database';
            case 'yaml':
            case 'yml':
                return 'file-text';
            default:
                return 'code';
        }
    }

    async onLoadFile(file: TFile): Promise<void> {
        this.currentFile = file;
        this.searchMatches = [];
        this.currentMatchIndex = -1;
        this.isSearchVisible = false;
        this.determineLanguage(file.extension);
        await this.renderCode();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.currentFile = null;
        this.rawContent = '';
        this.formattedContent = '';
        this.searchMatches = [];
        this.contentEl.empty();
    }

    private determineLanguage(ext: string): void {
        const lower = ext.toLowerCase();
        switch (lower) {
            case 'json':
                this.language = 'json';
                break;
            case 'sql':
                this.language = 'sql';
                break;
            case 'yaml':
            case 'yml':
                this.language = 'yaml';
                break;
            default:
                this.language = 'plaintext';
                break;
        }
    }

    private getFileIconEmoji(): string {
        switch (this.language) {
            case 'json':
                return '{ }';
            case 'sql':
                return '🗄️';
            case 'yaml':
                return '📋';
            default:
                return '📄';
        }
    }

    private formatDate(mtime: number): string {
        const d = new Date(mtime);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        const hours = String(d.getHours()).padStart(2, '0');
        const minutes = String(d.getMinutes()).padStart(2, '0');
        return `${year}-${month}-${day} ${hours}:${minutes}`;
    }

    private formatFileSize(bytes: number): string {
        if (bytes < 1024) return `${bytes} B`;
        if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
        return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    }

    private async renderCode(): Promise<void> {
        if (!this.currentFile) return;

        const container = this.contentEl;
        container.empty();
        container.addClass('office-preview-code-container');

        // 1. 顶部主工具栏
        const toolbar = container.createEl('div', { cls: 'office-preview-toolbar' });

        const titleContainer = toolbar.createEl('div', { cls: 'office-preview-title-container' });
        titleContainer.createEl('span', {
            cls: 'office-preview-title',
            text: `${this.getFileIconEmoji()} ${this.currentFile.name}`
        });

        if (this.currentFile.stat) {
            if (this.currentFile.stat.mtime) {
                const timeStr = this.formatDate(this.currentFile.stat.mtime);
                titleContainer.createEl('span', {
                    cls: 'office-preview-mtime',
                    text: `修改时间: ${timeStr}`
                });
            }
            if (this.currentFile.stat.size) {
                const sizeStr = this.formatFileSize(this.currentFile.stat.size);
                titleContainer.createEl('span', {
                    cls: 'office-preview-mtime',
                    text: `(${sizeStr})`
                });
            }
        }

        const controls = toolbar.createEl('div', { cls: 'office-preview-controls' });

        // JSON 专属：格式化 / 紧凑切换按钮
        let jsonToggleBtn: HTMLButtonElement | null = null;
        if (this.language === 'json') {
            jsonToggleBtn = controls.createEl('button', {
                cls: 'office-preview-btn',
                text: this.isJsonFormatted ? '✨ 原始紧凑' : '✨ 格式化'
            });
        }

        // 大纲面板切换按钮 (仅对 json, yaml 生效)
        let toggleOutlineBtn: HTMLButtonElement | null = null;
        const supportsOutline = this.language === 'json' || this.language === 'yaml';
        if (supportsOutline) {
            toggleOutlineBtn = controls.createEl('button', {
                cls: `office-preview-btn ${this.isOutlineVisible ? 'is-active' : ''}`,
                text: '🗂️ 大纲'
            });
        }

        // 一键复制代码按钮
        const copyBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '📋 复制' });

        // 搜索触发按钮
        const toggleSearchBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '🔍 查找' });

        // 外部应用打开
        const openExternalBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '↗ 外部打开' });
        openExternalBtn.onclick = () => this.openWithDefaultApp();

        // 刷新按钮
        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderCode();

        // 2. 悬浮搜索控制面板
        this.searchContainerEl = container.createEl('div', { cls: 'office-preview-search-bar is-hidden' });

        this.searchInputEl = this.searchContainerEl.createEl('input', {
            cls: 'office-preview-search-input',
            type: 'text',
            placeholder: `搜索 ${this.language.toUpperCase()} 内容 (Cmd+F)...`
        });

        this.searchCountEl = this.searchContainerEl.createEl('span', { cls: 'office-preview-search-count', text: '0 / 0' });

        const prevMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔼' });
        const nextMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔽' });
        const closeSearchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '✖' });

        // 3. 代码双栏主视口布局 (左侧内容 60% : 右侧大纲 40%)
        const mainBodyEl = container.createEl('div', { cls: 'office-preview-code-body' });
        
        const showOutlineInitially = supportsOutline && this.isOutlineVisible;
        const renderWrapper = mainBodyEl.createEl('div', {
            cls: `office-preview-code-wrapper ${!showOutlineInitially ? 'is-full-width' : ''}`
        });
        const outlinePane = mainBodyEl.createEl('div', {
            cls: `office-preview-outline-pane ${!showOutlineInitially ? 'is-hidden' : ''}`
        });

        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: `正在读取 ${this.language.toUpperCase()} 内容...` });

        try {
            this.rawContent = await this.app.vault.read(this.currentFile);
            loadingEl.remove();

            // 如果是 JSON，尝试解析并格式化
            if (this.language === 'json') {
                try {
                    const parsed = JSON.parse(this.rawContent);
                    this.formattedContent = JSON.stringify(parsed, null, 2);
                } catch (e) {
                    console.warn('JSON parse error (displaying raw content):', e);
                    this.formattedContent = this.rawContent;
                }
            }

            // 渲染代码 DOM 树
            const codeMount = renderWrapper.createEl('div', { cls: 'code-render-mount' });
            
            const displayCode = (this.language === 'json' && this.isJsonFormatted)
                ? this.formattedContent
                : this.rawContent;

            this.renderSyntaxTree(codeMount, displayCode);

            // 渲染右侧 40% 层级大纲
            if (supportsOutline) {
                this.renderOutlineTree(outlinePane, renderWrapper, displayCode);
            }

            // 大纲按钮切换事件 (保持 6:4 比例或一键 100% 全宽)
            if (toggleOutlineBtn) {
                toggleOutlineBtn.onclick = () => {
                    this.isOutlineVisible = !this.isOutlineVisible;
                    toggleOutlineBtn!.toggleClass('is-active', this.isOutlineVisible);
                    outlinePane.toggleClass('is-hidden', !this.isOutlineVisible);
                    renderWrapper.toggleClass('is-full-width', !this.isOutlineVisible);
                };
            }

            // JSON 切换点击事件
            if (jsonToggleBtn) {
                jsonToggleBtn.onclick = () => {
                    this.isJsonFormatted = !this.isJsonFormatted;
                    jsonToggleBtn!.setText(this.isJsonFormatted ? '✨ 原始紧凑' : '✨ 格式化');
                    codeMount.empty();
                    const codeToDisplay = this.isJsonFormatted ? this.formattedContent : this.rawContent;
                    this.renderSyntaxTree(codeMount, codeToDisplay);
                    if (supportsOutline) {
                        this.renderOutlineTree(outlinePane, renderWrapper, codeToDisplay);
                    }
                    if (this.isSearchVisible && this.searchInputEl?.value) {
                        performSearch(this.searchInputEl.value);
                    }
                };
            }

            // 一键复制功能
            copyBtn.onclick = async () => {
                const codeToCopy = (this.language === 'json' && this.isJsonFormatted)
                    ? this.formattedContent
                    : this.rawContent;
                try {
                    await navigator.clipboard.writeText(codeToCopy);
                    copyBtn.setText('✓ 已复制');
                    setTimeout(() => copyBtn.setText('📋 复制'), 1800);
                } catch (err) {
                    console.error('Failed to copy text:', err);
                }
            };

            // 搜索逻辑注册
            const performSearch = (query: string) => {
                this.clearSearchHighlights(codeMount);
                if (!query || query.trim() === '') {
                    this.searchMatches = [];
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                    return;
                }

                this.highlightText(codeMount, query.trim().toLowerCase());
                this.searchMatches = Array.from(codeMount.querySelectorAll('.code-search-highlight'));

                if (this.searchMatches.length > 0) {
                    this.currentMatchIndex = 0;
                    this.focusMatch(0);
                } else {
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                }
            };

            let searchTimer: any = null;
            this.searchInputEl.oninput = () => {
                if (searchTimer) clearTimeout(searchTimer);
                searchTimer = setTimeout(() => {
                    performSearch(this.searchInputEl?.value || '');
                }, 200);
            };

            this.searchInputEl.onkeydown = (e: KeyboardEvent) => {
                if (e.key === 'Enter') {
                    if (e.shiftKey) this.navigateMatch(-1);
                    else this.navigateMatch(1);
                } else if (e.key === 'Escape') {
                    this.toggleSearch(false, codeMount);
                }
            };

            prevMatchBtn.onclick = () => this.navigateMatch(-1);
            nextMatchBtn.onclick = () => this.navigateMatch(1);
            closeSearchBtn.onclick = () => this.toggleSearch(false, codeMount);
            toggleSearchBtn.onclick = () => this.toggleSearch(!this.isSearchVisible, codeMount);

            // 快捷键监听
            container.tabIndex = 0;
            container.onkeydown = (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, codeMount);
                }
            };

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createEl('div', {
                cls: 'office-preview-error',
                text: `读取文件失败: ${error instanceof Error ? error.message : String(error)}`
            });
        }
    }

    /**
     * 渲染带行号与语法高亮的 DOM
     */
    private renderSyntaxTree(mount: HTMLElement, codeText: string): void {
        const grammar = Prism.languages[this.language] || Prism.languages.plaintext;
        const highlightedHtml = Prism.highlight(codeText, grammar, this.language);

        const lines = highlightedHtml.split('\n');

        const codeContainer = mount.createEl('div', { cls: 'code-viewer-container' });
        const table = codeContainer.createEl('table', { cls: 'code-table' });
        const tbody = table.createEl('tbody');

        lines.forEach((lineHtml, idx) => {
            const tr = tbody.createEl('tr', { cls: 'code-line-tr' });
            tr.setAttribute('data-line-number', String(idx + 1));
            
            // 行号单元格
            const lineNumTd = tr.createEl('td', {
                cls: 'code-line-number',
                text: String(idx + 1)
            });

            // 代码单元格
            const codeTd = tr.createEl('td', { cls: 'code-line-content' });
            const pre = codeTd.createEl('pre', { cls: `language-${this.language}` });
            const code = pre.createEl('code', { cls: `language-${this.language}` });
            
            // 空行处理
            code.innerHTML = lineHtml || '&nbsp;';
        });
    }

    /**
     * 构建右侧 40% 宽度的结构大纲面板
     */
    private renderOutlineTree(outlinePane: HTMLElement, codeWrapper: HTMLElement, codeText: string): void {
        outlinePane.empty();

        const nodes = this.language === 'json'
            ? this.extractJsonOutline(codeText)
            : this.extractYamlOutline(codeText);

        const header = outlinePane.createEl('div', { cls: 'outline-header' });
        header.createEl('span', { cls: 'outline-title', text: '🗂️ 层级大纲 (40%)' });
        header.createEl('span', { cls: 'outline-count', text: `${nodes.length} 项` });

        if (nodes.length === 0) {
            const emptyEl = outlinePane.createEl('div', { cls: 'outline-empty' });
            emptyEl.setText('未检测到层级结构');
            return;
        }

        const listEl = outlinePane.createEl('div', { cls: 'outline-list' });

        nodes.forEach((node) => {
            const itemEl = listEl.createEl('div', {
                cls: `outline-item level-${Math.min(node.level, 8)}`
            });
            itemEl.style.paddingLeft = `${node.level * 16 + 12}px`;

            // 类型图标
            const iconEl = itemEl.createEl('span', { cls: 'outline-item-icon' });
            if (node.type === 'object') iconEl.setText('{ }');
            else if (node.type === 'array') iconEl.setText('[ ]');
            else iconEl.setText('•');

            // 键名
            itemEl.createEl('span', { cls: 'outline-item-key', text: node.key });

            // 预览值/提示 (40% 宽度下支持显示更丰富的内容)
            if (node.preview) {
                itemEl.createEl('span', { cls: 'outline-item-preview', text: node.preview });
            }

            // 行号标签
            itemEl.createEl('span', { cls: 'outline-item-line', text: `:${node.lineNumber}` });

            // 点击平滑滚动定位至对应代码行并聚焦高亮
            itemEl.onclick = () => {
                const targetTr = codeWrapper.querySelector(`tr[data-line-number="${node.lineNumber}"]`);
                if (targetTr) {
                    targetTr.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    // 添加短暂聚焦动画
                    targetTr.addClass('is-focused-line');
                    setTimeout(() => targetTr.removeClass('is-focused-line'), 1600);
                }
            };
        });
    }

    /**
     * 从格式化 JSON 中提取层级大纲
     */
    private extractJsonOutline(jsonText: string): OutlineNode[] {
        const nodes: OutlineNode[] = [];
        const lines = jsonText.split('\n');

        lines.forEach((line, idx) => {
            const trimmed = line.trim();
            // 匹配类似于 "key": { 或 "key": [ 或 "key": "value"
            const match = line.match(/^(\s*)"([^"]+)"\s*:\s*(.*)$/);
            if (match) {
                const indent = match[1].length;
                const key = match[2];
                const rest = match[3].trim();
                const level = Math.floor(indent / 2);

                let type: 'object' | 'array' | 'value' = 'value';
                let preview: string | undefined = undefined;

                if (rest.startsWith('{')) {
                    type = 'object';
                } else if (rest.startsWith('[')) {
                    type = 'array';
                } else {
                    type = 'value';
                    preview = rest.replace(/,$/, '');
                    if (preview.length > 36) preview = preview.substring(0, 34) + '...';
                }

                nodes.push({
                    key,
                    path: key,
                    type,
                    preview,
                    level,
                    lineNumber: idx + 1
                });
            }
        });

        return nodes;
    }

    /**
     * 从 YAML 文本中提取层级大纲
     */
    private extractYamlOutline(yamlText: string): OutlineNode[] {
        const nodes: OutlineNode[] = [];
        const lines = yamlText.split('\n');

        lines.forEach((line, idx) => {
            // 忽略注释与空行
            const trimmed = line.trim();
            if (!trimmed || trimmed.startsWith('#') || trimmed.startsWith('---')) return;

            // 匹配键行: "  key:" 或 "  - key: value" 或 "  key: value"
            const match = line.match(/^(\s*)(?:-\s+)?([a-zA-Z0-9_\-\.\/]+)\s*:\s*(.*)$/);
            if (match) {
                const indent = match[1].length;
                const key = match[2];
                const rest = match[3].trim();
                const level = Math.floor(indent / 2);

                let type: 'object' | 'array' | 'value' = 'value';
                let preview: string | undefined = undefined;

                if (!rest || rest.startsWith('#')) {
                    type = 'object';
                } else if (rest.startsWith('[')) {
                    type = 'array';
                } else {
                    type = 'value';
                    preview = rest.split('#')[0].trim();
                    if (preview.length > 36) preview = preview.substring(0, 34) + '...';
                }

                nodes.push({
                    key,
                    path: key,
                    type,
                    preview,
                    level,
                    lineNumber: idx + 1
                });
            }
        });

        return nodes;
    }

    private openWithDefaultApp(): void {
        if (!this.currentFile) return;
        try {
            if ((this.app as any).openWithDefaultApp) {
                (this.app as any).openWithDefaultApp(this.currentFile.path);
            } else {
                const electron = (window as any).require?.('electron');
                if (electron?.shell) {
                    const basePath = (this.app.vault.adapter as any).basePath || '';
                    const fullPath = basePath ? `${basePath}/${this.currentFile.path}` : this.currentFile.path;
                    electron.shell.openPath(fullPath);
                }
            }
        } catch (e) {
            console.error('Failed to open file with default app:', e);
        }
    }

    private toggleSearch(show: boolean, root: HTMLElement): void {
        this.isSearchVisible = show;
        if (this.searchContainerEl) {
            if (show) {
                this.searchContainerEl.removeClass('is-hidden');
                if (this.searchInputEl) {
                    this.searchInputEl.focus();
                    this.searchInputEl.select();
                }
            } else {
                this.searchContainerEl.addClass('is-hidden');
                this.clearSearchHighlights(root);
            }
        }
    }

    private navigateMatch(direction: number): void {
        if (this.searchMatches.length === 0) return;
        this.currentMatchIndex += direction;
        if (this.currentMatchIndex < 0) this.currentMatchIndex = this.searchMatches.length - 1;
        if (this.currentMatchIndex >= this.searchMatches.length) this.currentMatchIndex = 0;
        this.focusMatch(this.currentMatchIndex);
    }

    private focusMatch(index: number): void {
        this.searchMatches.forEach((el, i) => {
            if (i === index) el.addClass('is-active');
            else el.removeClass('is-active');
        });

        const activeEl = this.searchMatches[index];
        if (activeEl) {
            activeEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }

        if (this.searchCountEl) {
            this.searchCountEl.setText(`${index + 1} / ${this.searchMatches.length}`);
        }
    }

    private clearSearchHighlights(root: HTMLElement): void {
        const highlights = root.querySelectorAll('.code-search-highlight');
        highlights.forEach((mark) => {
            const parent = mark.parentNode;
            if (parent) {
                parent.replaceChild(document.createTextNode(mark.textContent || ''), mark);
                parent.normalize();
            }
        });
    }

    private highlightText(root: HTMLElement, query: string): void {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
        const nodesToReplace: { node: Text; matches: { index: number; length: number }[] }[] = [];

        let currentNode: Node | null = walker.nextNode();
        while (currentNode) {
            const textNode = currentNode as Text;
            const text = textNode.nodeValue || '';
            const lowerText = text.toLowerCase();

            // 忽略行号列
            if (textNode.parentElement && textNode.parentElement.hasClass('code-line-number')) {
                currentNode = walker.nextNode();
                continue;
            }

            if (lowerText.includes(query) && textNode.parentElement && !textNode.parentElement.hasClass('code-search-highlight')) {
                const matches: { index: number; length: number }[] = [];
                let startIndex = 0;
                let foundIndex = lowerText.indexOf(query, startIndex);
                while (foundIndex !== -1) {
                    matches.push({ index: foundIndex, length: query.length });
                    startIndex = foundIndex + query.length;
                    foundIndex = lowerText.indexOf(query, startIndex);
                }
                nodesToReplace.push({ node: textNode, matches });
            }
            currentNode = walker.nextNode();
        }

        nodesToReplace.forEach(({ node, matches }) => {
            const text = node.nodeValue || '';
            const fragment = document.createDocumentFragment();
            let lastIndex = 0;

            matches.forEach(({ index, length }) => {
                if (index > lastIndex) {
                    fragment.appendChild(document.createTextNode(text.substring(lastIndex, index)));
                }
                const mark = document.createElement('mark');
                mark.className = 'code-search-highlight';
                mark.textContent = text.substring(index, index + length);
                fragment.appendChild(mark);
                lastIndex = index + length;
            });

            if (lastIndex < text.length) {
                fragment.appendChild(document.createTextNode(text.substring(lastIndex)));
            }

            if (node.parentNode) {
                node.parentNode.replaceChild(fragment, node);
            }
        });
    }
}
