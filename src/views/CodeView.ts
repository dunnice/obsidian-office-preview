import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import Prism from 'prismjs';
import 'prismjs/components/prism-json';
import 'prismjs/components/prism-sql';
import 'prismjs/components/prism-yaml';
import 'prismjs/components/prism-python';
import 'prismjs/components/prism-properties';
import 'prismjs/components/prism-bash';
import 'prismjs/components/prism-ini';

export const VIEW_TYPE_CODE = 'office-preview-code-view';

interface OutlineTreeNode {
    id: string;
    key: string;
    path: string;
    type: 'object' | 'array' | 'value';
    preview?: string;
    level: number;
    lineNumber: number;
    children: OutlineTreeNode[];
    isCollapsed: boolean;
}

export class CodeView extends FileView {
    private currentFile: TFile | null = null;
    private rawContent: string = '';
    private formattedContent: string = '';
    private isJsonFormatted: boolean = true;
    private isWordWrap: boolean = true; // 默认开启自动换行
    private language: string = 'plaintext';
    private isOutlineVisible: boolean = true;
    private outlineTree: OutlineTreeNode[] = [];

    // 编辑模式状态管理
    private isEditMode: boolean = false;
    private hasUnsavedChanges: boolean = false;
    private editorTextareaEl: HTMLTextAreaElement | null = null;
    private editorLineNumbersEl: HTMLElement | null = null;
    private editStatusBadgeEl: HTMLElement | null = null;
    private cursorInfoEl: HTMLElement | null = null;
    private modeToggleBtn: HTMLButtonElement | null = null;
    private saveBtn: HTMLButtonElement | null = null;
    private mtimeEl: HTMLElement | null = null;
    private fileSizeEl: HTMLElement | null = null;

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
        const ext = this.currentFile?.extension?.toLowerCase() || '';
        const name = this.currentFile?.name?.toLowerCase() || '';

        if (name.includes('.env') || ext === 'env' || name === '.env') return 'sliders-horizontal';
        if (ext === 'py' || ext === 'pyw') return 'code-glyph';
        if (ext === 'sql') return 'database';
        if (ext === 'json') return 'code-glyph';
        if (ext === 'yaml' || ext === 'yml') return 'file-text';
        if (ext === 'txt') return 'document';
        return 'code';
    }

    async onLoadFile(file: TFile): Promise<void> {
        this.currentFile = file;
        this.searchMatches = [];
        this.currentMatchIndex = -1;
        this.isSearchVisible = false;
        this.isEditMode = false;
        this.hasUnsavedChanges = false;
        this.determineLanguage(file.extension, file.name);
        await this.renderCode();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        // 如果离开时有未保存的编辑内容，自动静默保存
        if (this.hasUnsavedChanges && this.editorTextareaEl && this.currentFile) {
            try {
                await this.app.vault.modify(this.currentFile, this.editorTextareaEl.value);
            } catch {
                // ignore
            }
        }
        this.currentFile = null;
        this.rawContent = '';
        this.formattedContent = '';
        this.searchMatches = [];
        this.outlineTree = [];
        this.isEditMode = false;
        this.hasUnsavedChanges = false;
        this.editorTextareaEl = null;
        this.editorLineNumbersEl = null;
        this.editStatusBadgeEl = null;
        this.cursorInfoEl = null;
        this.modeToggleBtn = null;
        this.saveBtn = null;
        this.mtimeEl = null;
        this.fileSizeEl = null;
        this.contentEl.empty();
    }

    private determineLanguage(ext: string, filename: string = ''): void {
        const lowerExt = (ext || '').toLowerCase();
        const lowerName = (filename || '').toLowerCase();

        // 1. 优先根据文件名特征判断（环境配置与各类点文件）
        if (lowerName.includes('.env') || lowerName === '.env' || lowerExt === 'env') {
            this.language = 'properties';
            return;
        }
        if (lowerName.endsWith('.gitignore') || lowerName.endsWith('.npmignore') || lowerName.endsWith('.dockerignore')) {
            this.language = 'bash';
            return;
        }
        if (lowerName === '.editorconfig' || lowerName.endsWith('.ini') || lowerName.endsWith('.cfg')) {
            this.language = 'ini';
            return;
        }

        // 2. 根据扩展名判断
        switch (lowerExt) {
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
            case 'py':
            case 'pyw':
                this.language = 'python';
                break;
            case 'conf':
            case 'properties':
                this.language = 'properties';
                break;
            case 'sh':
            case 'bash':
            case 'zsh':
                this.language = 'bash';
                break;
            case 'txt':
            default:
                this.language = 'plaintext';
                break;
        }
    }

    private getFileIconEmoji(): string {
        const name = this.currentFile?.name?.toLowerCase() || '';
        if (name.includes('.env') || this.language === 'properties') return '⚙️';
        if (this.language === 'python') return '🐍';
        if (this.language === 'sql') return '🗄️';
        if (this.language === 'json') return '{ }';
        if (this.language === 'yaml') return '📋';
        if (this.language === 'bash') return '💻';
        return '📄';
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

    private refreshFileMeta(): void {
        if (!this.currentFile?.stat) return;
        if (this.mtimeEl && this.currentFile.stat.mtime) {
            this.mtimeEl.setText(`修改时间: ${this.formatDate(this.currentFile.stat.mtime)}`);
        }
        if (this.fileSizeEl && this.currentFile.stat.size) {
            this.fileSizeEl.setText(`(${this.formatFileSize(this.currentFile.stat.size)})`);
        }
    }

    private updateEditStatus(): void {
        if (!this.editStatusBadgeEl) return;
        if (this.hasUnsavedChanges) {
            this.editStatusBadgeEl.setText('● 未保存修改');
            this.editStatusBadgeEl.addClass('is-dirty');
            this.editStatusBadgeEl.removeClass('is-saved');
        } else {
            this.editStatusBadgeEl.setText('✓ 已保存');
            this.editStatusBadgeEl.addClass('is-saved');
            this.editStatusBadgeEl.removeClass('is-dirty');
        }
    }

    private async saveContent(): Promise<void> {
        if (!this.currentFile || !this.editorTextareaEl) return;
        const newContent = this.editorTextareaEl.value;
        try {
            await this.app.vault.modify(this.currentFile, newContent);
            this.rawContent = newContent;
            this.hasUnsavedChanges = false;
            this.updateEditStatus();

            if (this.saveBtn) {
                this.saveBtn.setText('✓ 已保存');
                window.setTimeout(() => {
                    if (this.saveBtn) this.saveBtn.setText('💾 保存');
                }, 1500);
            }
            this.refreshFileMeta();
        } catch (err) {
            console.error('Failed to save file:', err);
            if (this.saveBtn) {
                this.saveBtn.setText('❌ 保存失败');
                window.setTimeout(() => {
                    if (this.saveBtn) this.saveBtn.setText('💾 保存');
                }, 2000);
            }
        }
    }

    private async renderCode(): Promise<void> {
        if (!this.currentFile) return;

        const container = this.contentEl;
        container.empty();
        container.addClass('office-preview-code-container');

        // 1. 顶部主工具栏
        const toolbar = container.createDiv({ cls: 'office-preview-toolbar' });

        const titleContainer = toolbar.createDiv({ cls: 'office-preview-title-container' });
        titleContainer.createSpan({
            cls: 'office-preview-title',
            text: `${this.getFileIconEmoji()} ${this.currentFile.name}`
        });

        // 编辑状态徽章 (未保存 / 已保存)
        this.editStatusBadgeEl = titleContainer.createSpan({
            cls: `office-preview-status-badge ${this.isEditMode ? '' : 'is-hidden'}`,
            text: '✓ 已保存'
        });

        // 光标与行列统计信息 (编辑模式展示)
        this.cursorInfoEl = titleContainer.createSpan({
            cls: `office-preview-cursor-info ${this.isEditMode ? '' : 'is-hidden'}`,
            text: ''
        });

        if (this.currentFile.stat) {
            if (this.currentFile.stat.mtime) {
                this.mtimeEl = titleContainer.createSpan({
                    cls: 'office-preview-mtime',
                    text: `修改时间: ${this.formatDate(this.currentFile.stat.mtime)}`
                });
            }
            if (this.currentFile.stat.size) {
                this.fileSizeEl = titleContainer.createSpan({
                    cls: 'office-preview-mtime',
                    text: `(${this.formatFileSize(this.currentFile.stat.size)})`
                });
            }
        }

        const controls = toolbar.createDiv({ cls: 'office-preview-controls' });

        // 模式切换按钮 (✏️ 编辑 / 👁️ 预览)
        this.modeToggleBtn = controls.createEl('button', {
            cls: `office-preview-btn ${this.isEditMode ? 'is-active' : ''}`,
            text: this.isEditMode ? '👁️ 预览' : '✏️ 编辑'
        });

        // 快速保存按钮 (💾 保存)
        this.saveBtn = controls.createEl('button', {
            cls: `office-preview-btn office-preview-save-btn ${this.isEditMode ? '' : 'is-hidden'}`,
            text: '💾 保存',
            title: '保存修改 (Cmd+S / Ctrl+S)'
        });
        this.saveBtn.onclick = async () => await this.saveContent();

        // JSON 专属：格式化 / 紧凑切换按钮
        let jsonToggleBtn: HTMLButtonElement | null = null;
        if (this.language === 'json') {
            jsonToggleBtn = controls.createEl('button', {
                cls: 'office-preview-btn',
                text: this.isJsonFormatted ? '✨ 原始紧凑' : '✨ 格式化'
            });
        }

        // 自动换行切换按钮
        const wrapToggleBtn = controls.createEl('button', {
            cls: `office-preview-btn ${this.isWordWrap ? 'is-active' : ''}`,
            text: this.isWordWrap ? '↩️ 换行开启' : '↩️ 换行关闭'
        });

        // 大纲面板切换按钮 (对 json, yaml, sql, python, properties 生效)
        let toggleOutlineBtn: HTMLButtonElement | null = null;
        const supportsOutline = ['json', 'yaml', 'sql', 'python', 'properties', 'bash'].includes(this.language);
        if (supportsOutline) {
            toggleOutlineBtn = controls.createEl('button', {
                cls: `office-preview-btn ${this.isOutlineVisible ? 'is-active' : ''} ${this.isEditMode ? 'is-hidden' : ''}`,
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
        reloadBtn.onclick = async () => {
            if (this.hasUnsavedChanges) {
                await this.saveContent();
            }
            await this.renderCode();
        };

        // 2. 悬浮搜索控制面板
        this.searchContainerEl = container.createDiv({ cls: 'office-preview-search-bar is-hidden' });

        this.searchInputEl = this.searchContainerEl.createEl('input', {
            cls: 'office-preview-search-input',
            type: 'text',
            placeholder: `搜索 ${this.language.toUpperCase()} 内容 (Cmd+F)...`
        });

        this.searchCountEl = this.searchContainerEl.createSpan({ cls: 'office-preview-search-count', text: '0 / 0' });

        const prevMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔼' });
        const nextMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔽' });
        const closeSearchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '✖' });

        // 3. 主视口容器
        const mainBodyEl = container.createDiv({ cls: 'office-preview-code-body' });
        
        const showOutlineInitially = supportsOutline && this.isOutlineVisible && !this.isEditMode;
        
        // 预览视口与大纲面板
        const renderWrapper = mainBodyEl.createDiv({
            cls: `office-preview-code-wrapper ${!showOutlineInitially ? 'is-full-width' : ''} ${this.isWordWrap ? 'is-word-wrap' : ''} ${this.isEditMode ? 'is-hidden' : ''}`
        });
        const outlinePane = mainBodyEl.createDiv({
            cls: `office-preview-outline-pane ${!showOutlineInitially ? 'is-hidden' : ''}`
        });

        // 编辑视口
        const editorWrapper = mainBodyEl.createDiv({
            cls: `office-preview-editor-wrapper ${this.isWordWrap ? 'is-word-wrap' : ''} ${!this.isEditMode ? 'is-hidden' : ''}`
        });

        const loadingEl = renderWrapper.createDiv({ cls: 'office-preview-loading', text: `正在读取 ${this.currentFile.name}...` });

        try {
            this.rawContent = await this.app.vault.read(this.currentFile);
            loadingEl.remove();

            // 如果是 JSON，尝试解析并格式化
            if (this.language === 'json') {
                try {
                    const parsed: unknown = JSON.parse(this.rawContent);
                    this.formattedContent = JSON.stringify(parsed, null, 2);
                } catch {
                    this.formattedContent = this.rawContent;
                }
            }

            // 渲染只读代码 DOM 树
            const codeMount = renderWrapper.createDiv({ cls: 'code-render-mount' });
            
            const displayCode = (this.language === 'json' && this.isJsonFormatted)
                ? this.formattedContent
                : this.rawContent;

            this.renderSyntaxTree(codeMount, displayCode);

            // 渲染右侧层级大纲
            if (supportsOutline) {
                this.renderOutlineTree(outlinePane, renderWrapper, displayCode);
            }

            // 构建编辑器 DOM
            this.buildEditorDOM(editorWrapper, this.rawContent);

            // 模式切换事件 (✏️ 编辑 / 👁️ 预览)
            this.modeToggleBtn.onclick = async () => {
                await this.toggleEditMode(
                    renderWrapper,
                    outlinePane,
                    editorWrapper,
                    codeMount,
                    supportsOutline,
                    toggleOutlineBtn
                );
            };

            // 换行模式切换事件
            wrapToggleBtn.onclick = () => {
                this.isWordWrap = !this.isWordWrap;
                wrapToggleBtn.toggleClass('is-active', this.isWordWrap);
                wrapToggleBtn.setText(this.isWordWrap ? '↩️ 换行开启' : '↩️ 换行关闭');
                renderWrapper.toggleClass('is-word-wrap', this.isWordWrap);
                editorWrapper.toggleClass('is-word-wrap', this.isWordWrap);
            };

            // 大纲按钮切换事件 (保持 7:3 比例或一键 100% 全宽)
            if (toggleOutlineBtn) {
                toggleOutlineBtn.onclick = () => {
                    this.isOutlineVisible = !this.isOutlineVisible;
                    toggleOutlineBtn.toggleClass('is-active', this.isOutlineVisible);
                    outlinePane.toggleClass('is-hidden', !this.isOutlineVisible);
                    renderWrapper.toggleClass('is-full-width', !this.isOutlineVisible);
                };
            }

            // JSON 切换点击事件
            if (jsonToggleBtn) {
                jsonToggleBtn.onclick = () => {
                    this.isJsonFormatted = !this.isJsonFormatted;
                    jsonToggleBtn.setText(this.isJsonFormatted ? '✨ 原始紧凑' : '✨ 格式化');
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
                const codeToCopy = this.isEditMode
                    ? (this.editorTextareaEl?.value || this.rawContent)
                    : ((this.language === 'json' && this.isJsonFormatted)
                        ? this.formattedContent
                        : this.rawContent);
                try {
                    await navigator.clipboard.writeText(codeToCopy);
                    copyBtn.setText('✓ 已复制');
                    window.setTimeout(() => copyBtn.setText('📋 复制'), 1800);
                } catch {
                    // ignore
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

            let searchTimer: number | null = null;
            this.searchInputEl.oninput = () => {
                if (searchTimer !== null) window.clearTimeout(searchTimer);
                searchTimer = window.setTimeout(() => {
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

            // 全局快捷键监听 (Cmd+F 搜索，Cmd+S 保存)
            container.tabIndex = 0;
            container.onkeydown = async (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, codeMount);
                } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
                    if (this.isEditMode) {
                        e.preventDefault();
                        await this.saveContent();
                    }
                }
            };

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createDiv({
                cls: 'office-preview-error',
                text: `读取文件失败: ${error instanceof Error ? error.message : String(error)}`
            });
        }
    }

    /**
     * 构建内置沉浸式文本/代码编辑器 DOM
     */
    private buildEditorDOM(editorWrapper: HTMLElement, initialText: string): void {
        editorWrapper.empty();

        const gutter = editorWrapper.createDiv({ cls: 'editor-gutter' });
        this.editorLineNumbersEl = gutter.createDiv({ cls: 'editor-line-numbers' });

        const textareaContainer = editorWrapper.createDiv({ cls: 'editor-textarea-container' });
        this.editorTextareaEl = textareaContainer.createEl('textarea', {
            cls: 'office-preview-textarea',
            attr: {
                spellcheck: 'false',
                placeholder: `开始编辑 ${this.currentFile?.name || ''}...`
            }
        });
        this.editorTextareaEl.value = initialText;

        const updateLineNumbers = () => {
            if (!this.editorTextareaEl || !this.editorLineNumbersEl) return;
            const lineCount = (this.editorTextareaEl.value.match(/\n/g) || []).length + 1;
            let numbers = '';
            for (let i = 1; i <= lineCount; i++) {
                numbers += i + '\n';
            }
            this.editorLineNumbersEl.setText(numbers);
        };

        const updateCursorInfo = () => {
            if (!this.editorTextareaEl || !this.cursorInfoEl) return;
            const val = this.editorTextareaEl.value;
            const selStart = this.editorTextareaEl.selectionStart;
            const linesUpToCursor = val.substring(0, selStart).split('\n');
            const currentLine = linesUpToCursor.length;
            const currentCol = linesUpToCursor[linesUpToCursor.length - 1].length + 1;
            const totalLines = (val.match(/\n/g) || []).length + 1;
            this.cursorInfoEl.setText(`第 ${currentLine} 行, 第 ${currentCol} 列 | 共 ${totalLines} 行, ${val.length} 字符`);
        };

        updateLineNumbers();
        updateCursorInfo();

        // 联动滚动：textarea 滚动与行号严格像素对齐
        this.editorTextareaEl.addEventListener('scroll', () => {
            if (this.editorLineNumbersEl && this.editorTextareaEl) {
                this.editorLineNumbersEl.scrollTop = this.editorTextareaEl.scrollTop;
            }
        });

        // 输入监听
        this.editorTextareaEl.addEventListener('input', () => {
            this.hasUnsavedChanges = true;
            this.updateEditStatus();
            updateLineNumbers();
            updateCursorInfo();
        });

        // 光标与选择位置追踪
        this.editorTextareaEl.addEventListener('click', updateCursorInfo);
        this.editorTextareaEl.addEventListener('keyup', (e: KeyboardEvent) => {
            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown'].includes(e.key)) {
                updateCursorInfo();
            }
        });

        // 快捷键支持：Tab 缩进与 Cmd+S 保存
        this.editorTextareaEl.addEventListener('keydown', async (e: KeyboardEvent) => {
            if (e.key === 'Tab') {
                e.preventDefault();
                const start = this.editorTextareaEl!.selectionStart;
                const end = this.editorTextareaEl!.selectionEnd;
                const val = this.editorTextareaEl!.value;
                this.editorTextareaEl!.value = val.substring(0, start) + '    ' + val.substring(end);
                this.editorTextareaEl!.selectionStart = this.editorTextareaEl!.selectionEnd = start + 4;
                this.hasUnsavedChanges = true;
                this.updateEditStatus();
                updateLineNumbers();
                updateCursorInfo();
            } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
                e.preventDefault();
                await this.saveContent();
            }
        });
    }

    /**
     * 切换【编辑模式】与【预览模式】
     */
    private async toggleEditMode(
        renderWrapper: HTMLElement,
        outlinePane: HTMLElement,
        editorWrapper: HTMLElement,
        codeMount: HTMLElement,
        supportsOutline: boolean,
        toggleOutlineBtn: HTMLButtonElement | null
    ): Promise<void> {
        if (this.isEditMode) {
            // 切回预览模式：若有改动，先自动保存，保证内容无缝同步
            if (this.hasUnsavedChanges) {
                await this.saveContent();
            }
            this.isEditMode = false;
            if (this.modeToggleBtn) {
                this.modeToggleBtn.setText('✏️ 编辑');
                this.modeToggleBtn.removeClass('is-active');
            }
            if (this.saveBtn) this.saveBtn.addClass('is-hidden');
            if (this.editStatusBadgeEl) this.editStatusBadgeEl.addClass('is-hidden');
            if (this.cursorInfoEl) this.cursorInfoEl.addClass('is-hidden');

            editorWrapper.addClass('is-hidden');
            renderWrapper.removeClass('is-hidden');

            if (supportsOutline && toggleOutlineBtn) {
                toggleOutlineBtn.removeClass('is-hidden');
                if (this.isOutlineVisible) {
                    outlinePane.removeClass('is-hidden');
                    renderWrapper.removeClass('is-full-width');
                }
            }

            // 更新高亮渲染
            codeMount.empty();
            const displayCode = (this.language === 'json' && this.isJsonFormatted)
                ? this.formattedContent
                : this.rawContent;
            this.renderSyntaxTree(codeMount, displayCode);
            if (supportsOutline) {
                this.renderOutlineTree(outlinePane, renderWrapper, displayCode);
            }
        } else {
            // 进入编辑模式
            this.isEditMode = true;
            if (this.modeToggleBtn) {
                this.modeToggleBtn.setText('👁️ 预览');
                this.modeToggleBtn.addClass('is-active');
            }
            if (this.saveBtn) this.saveBtn.removeClass('is-hidden');
            if (this.editStatusBadgeEl) this.editStatusBadgeEl.removeClass('is-hidden');
            if (this.cursorInfoEl) this.cursorInfoEl.removeClass('is-hidden');

            renderWrapper.addClass('is-hidden');
            outlinePane.addClass('is-hidden');
            editorWrapper.removeClass('is-hidden');

            if (toggleOutlineBtn) {
                toggleOutlineBtn.addClass('is-hidden');
            }

            if (this.editorTextareaEl) {
                this.editorTextareaEl.value = this.rawContent;
                this.editorTextareaEl.focus();
            }
        }
    }

    /**
     * 渲染带行号与语法高亮的 DOM (完全杜绝 unsafe innerHTML 赋值)
     */
    private renderSyntaxTree(mount: HTMLElement, codeText: string): void {
        const grammar = Prism.languages[this.language] || Prism.languages.plaintext;
        const highlightedHtml = Prism.highlight(codeText, grammar, this.language);
        const normalizedHtml = highlightedHtml.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const lines = normalizedHtml.split('\n');

        const codeContainer = mount.createDiv({ cls: 'code-viewer-container' });
        const table = codeContainer.createEl('table', { cls: 'code-table' });
        const tbody = table.createEl('tbody');

        const domParser = new DOMParser();

        lines.forEach((lineHtml, idx) => {
            const tr = tbody.createEl('tr', { cls: 'code-line-tr' });
            tr.setAttribute('data-line-number', String(idx + 1));
            
            // 行号单元格
            tr.createEl('td', {
                cls: 'code-line-number',
                text: String(idx + 1)
            });

            // 代码单元格
            const codeTd = tr.createEl('td', { cls: 'code-line-content' });
            const pre = codeTd.createEl('pre', { cls: `language-${this.language}` });
            const code = pre.createEl('code', { cls: `language-${this.language}` });
            
            if (!lineHtml) {
                code.setText('\u00A0');
            } else {
                // 1. 显式提取行首所有前导缩进空白（空格与制表符），用纯文本节点稳固承载，绝不被任何 HTML 解析器过滤
                const indentMatch = lineHtml.match(/^([ \t]+)/);
                const indent = indentMatch ? indentMatch[1] : '';
                const restHtml = lineHtml.slice(indent.length);

                if (indent) {
                    const indentSpan = code.createSpan({ cls: 'code-indent-spaces' });
                    indentSpan.setText(indent);
                }

                // 2. 将剩余高亮内容安全解析为子节点
                if (restHtml) {
                    const parsedDoc = domParser.parseFromString(restHtml, 'text/html');
                    while (parsedDoc.body.firstChild) {
                        code.appendChild(parsedDoc.body.firstChild);
                    }
                } else if (!indent) {
                    code.setText('\u00A0');
                }
            }
        });
    }

    /**
     * 构建右侧结构大纲面板 (支持多级折叠与一键全部折叠/展开)
     */
    private renderOutlineTree(outlinePane: HTMLElement, codeWrapper: HTMLElement, codeText: string): void {
        outlinePane.empty();

        if (this.language === 'json') {
            this.outlineTree = this.buildJsonOutlineTree(codeText);
        } else if (this.language === 'sql') {
            this.outlineTree = this.buildSqlOutlineTree(codeText);
        } else if (this.language === 'python') {
            this.outlineTree = this.buildPythonOutlineTree(codeText);
        } else if (this.language === 'properties' || this.language === 'bash') {
            this.outlineTree = this.buildEnvOutlineTree(codeText);
        } else {
            this.outlineTree = this.buildYamlOutlineTree(codeText);
        }

        const totalNodes = this.countTreeNodes(this.outlineTree);

        // 大纲顶部工具栏
        const header = outlinePane.createDiv({ cls: 'outline-header' });
        const headerLeft = header.createDiv({ cls: 'outline-header-left' });
        headerLeft.createSpan({ cls: 'outline-title', text: '🗂️ 层级大纲' });
        headerLeft.createSpan({ cls: 'outline-count', text: `${totalNodes} 项` });

        // 快速折叠与展开按钮组
        const headerActions = header.createDiv({ cls: 'outline-header-actions' });
        const collapseAllBtn = headerActions.createEl('button', {
            cls: 'outline-action-icon-btn',
            text: '🔼',
            title: '折叠全部层级'
        });
        const expandAllBtn = headerActions.createEl('button', {
            cls: 'outline-action-icon-btn',
            text: '🔽',
            title: '展开全部层级'
        });

        if (this.outlineTree.length === 0) {
            const emptyEl = outlinePane.createDiv({ cls: 'outline-empty' });
            emptyEl.setText('未检测到层级结构');
            return;
        }

        const listEl = outlinePane.createDiv({ cls: 'outline-list' });

        // 递归渲染节点 DOM
        const renderNodesRecursive = (nodes: OutlineTreeNode[], parentContainer: HTMLElement) => {
            nodes.forEach((node) => {
                const nodeRowEl = parentContainer.createDiv({
                    cls: `outline-item level-${Math.min(node.level, 8)}`
                });
                nodeRowEl.style.paddingLeft = `${node.level * 14 + 8}px`;

                const hasChildren = node.children && node.children.length > 0;

                // 折叠/展开小箭头 (父节点可点击折叠)
                const arrowEl = nodeRowEl.createSpan({
                    cls: `outline-toggle-arrow ${hasChildren ? 'has-children' : 'is-leaf'} ${node.isCollapsed ? 'is-collapsed' : ''}`,
                    text: hasChildren ? (node.isCollapsed ? '▶' : '▼') : ' '
                });

                // 类型图标
                const iconEl = nodeRowEl.createSpan({ cls: 'outline-item-icon' });
                if (node.type === 'object') iconEl.setText('{ }');
                else if (node.type === 'array') iconEl.setText('[ ]');
                else iconEl.setText('•');

                // 键名
                nodeRowEl.createSpan({ cls: 'outline-item-key', text: node.key });

                // 折叠时显示包含的子节点数量提示
                const collapsedBadgeEl = nodeRowEl.createSpan({
                    cls: `outline-collapsed-badge ${node.isCollapsed ? '' : 'is-hidden'}`,
                    text: `{ ${node.children.length} }`
                });

                // 预览值/提示
                if (node.preview && !hasChildren) {
                    nodeRowEl.createSpan({ cls: 'outline-item-preview', text: node.preview });
                }

                // 行号标签
                nodeRowEl.createSpan({ cls: 'outline-item-line', text: `:${node.lineNumber}` });

                // 子节点容器
                let childrenContainer: HTMLElement | null = null;
                if (hasChildren) {
                    childrenContainer = parentContainer.createDiv({
                        cls: `outline-children-group ${node.isCollapsed ? 'is-hidden' : ''}`
                    });
                    renderNodesRecursive(node.children, childrenContainer);
                }

                // 点击小箭头折叠/展开当前节点
                if (hasChildren && childrenContainer) {
                    arrowEl.onclick = (e) => {
                        e.stopPropagation();
                        node.isCollapsed = !node.isCollapsed;
                        arrowEl.setText(node.isCollapsed ? '▶' : '▼');
                        arrowEl.toggleClass('is-collapsed', node.isCollapsed);
                        childrenContainer.toggleClass('is-hidden', node.isCollapsed);
                        collapsedBadgeEl.toggleClass('is-hidden', !node.isCollapsed);
                    };
                }

                // 点击节点行平滑定位并聚焦代码行
                nodeRowEl.onclick = () => {
                    const targetTr = codeWrapper.querySelector(`tr[data-line-number="${node.lineNumber}"]`);
                    if (targetTr) {
                        targetTr.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        targetTr.addClass('is-focused-line');
                        window.setTimeout(() => targetTr.removeClass('is-focused-line'), 1600);
                    }
                };
            });
        };

        renderNodesRecursive(this.outlineTree, listEl);

        // 全部折叠
        collapseAllBtn.onclick = () => {
            const setAllCollapsed = (nodes: OutlineTreeNode[], collapsed: boolean) => {
                nodes.forEach((n) => {
                    if (n.children.length > 0) {
                        n.isCollapsed = collapsed;
                        setAllCollapsed(n.children, collapsed);
                    }
                });
            };
            setAllCollapsed(this.outlineTree, true);
            this.renderOutlineTree(outlinePane, codeWrapper, codeText);
        };

        // 全部展开
        expandAllBtn.onclick = () => {
            const setAllCollapsed = (nodes: OutlineTreeNode[], collapsed: boolean) => {
                nodes.forEach((n) => {
                    if (n.children.length > 0) {
                        n.isCollapsed = collapsed;
                        setAllCollapsed(n.children, collapsed);
                    }
                });
            };
            setAllCollapsed(this.outlineTree, false);
            this.renderOutlineTree(outlinePane, codeWrapper, codeText);
        };
    }

    private countTreeNodes(nodes: OutlineTreeNode[]): number {
        let count = 0;
        nodes.forEach((n) => {
            count += 1;
            if (n.children.length > 0) {
                count += this.countTreeNodes(n.children);
            }
        });
        return count;
    }

    /**
     * 将扁平的 JSON 键行构建为多级树形结构
     */
    private buildJsonOutlineTree(jsonText: string): OutlineTreeNode[] {
        const flatNodes: OutlineTreeNode[] = [];
        const lines = jsonText.split('\n');

        lines.forEach((line, idx) => {
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

                flatNodes.push({
                    id: `json-${idx + 1}`,
                    key,
                    path: key,
                    type,
                    preview,
                    level,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
            }
        });

        return this.constructTreeFromFlatNodes(flatNodes);
    }

    /**
     * 将扁平的 YAML 键行构建为多级树形结构
     */
    private buildYamlOutlineTree(yamlText: string): OutlineTreeNode[] {
        const flatNodes: OutlineTreeNode[] = [];
        const lines = yamlText.split('\n');

        lines.forEach((line, idx) => {
            const trimmed = line.trim();
            if (!trimmed || trimmed.startsWith('#') || trimmed.startsWith('---')) return;

            const match = line.match(/^(\s*)(?:-\s+)?([a-zA-Z0-9_./-]+)\s*:\s*(.*)$/);
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

                flatNodes.push({
                    id: `yaml-${idx + 1}`,
                    key,
                    path: key,
                    type,
                    preview,
                    level,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
            }
        });

        return this.constructTreeFromFlatNodes(flatNodes);
    }

    /**
     * 构建 SQL 结构大纲 (提取章节注释、表、视图、索引 DDL 结构)
     */
    private buildSqlOutlineTree(sqlText: string): OutlineTreeNode[] {
        const flatNodes: OutlineTreeNode[] = [];
        const lines = sqlText.split('\n');

        lines.forEach((line, idx) => {
            const trimmed = line.trim();
            if (!trimmed) return;

            // 1. 注释章节标题: 以 -- 开头，后跟文字
            const commentMatch = trimmed.match(/^--\s*([#=*~-]+\s*)?([^\s=*~#-].*)$/);
            if (commentMatch && !trimmed.startsWith('-- =') && !trimmed.startsWith('-- -') && !trimmed.startsWith('-- *')) {
                const commentTitle = commentMatch[2].trim();
                if (commentTitle.length > 1 && !commentTitle.startsWith('Nacos') && !commentTitle.startsWith('http')) {
                    flatNodes.push({
                        id: `sql-sec-${idx + 1}`,
                        key: commentTitle,
                        path: commentTitle,
                        type: 'value',
                        preview: '章节',
                        level: 0,
                        lineNumber: idx + 1,
                        children: [],
                        isCollapsed: false
                    });
                    return;
                }
            }

            // 2. DDL 语句结构: CREATE TABLE, CREATE VIEW, CREATE INDEX 等
            const ddlMatch = trimmed.match(/^(?:CREATE\s+(?:OR\s+REPLACE\s+)?(?:TEMPORARY\s+)?(TABLE|VIEW|INDEX|PROCEDURE|FUNCTION)|ALTER\s+TABLE|DROP\s+TABLE)\s+(?:IF\s+NOT\s+EXISTS\s+)?(?:`?([a-zA-Z0-9_.-]+)`?|\"?([a-zA-Z0-9_.-]+)\"?)/i);
            if (ddlMatch) {
                const operation = (ddlMatch[1] || 'TABLE').toUpperCase();
                const targetName = ddlMatch[2] || ddlMatch[3] || '';
                const lastNodeIsSection = flatNodes.length > 0 && flatNodes[flatNodes.length - 1].type === 'value';
                flatNodes.push({
                    id: `sql-ddl-${idx + 1}`,
                    key: targetName || trimmed.substring(0, 30),
                    path: targetName || '',
                    type: 'object',
                    preview: operation,
                    level: lastNodeIsSection ? 1 : 0,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
            }
        });

        return this.constructTreeFromFlatNodes(flatNodes);
    }

    /**
     * 构建 Python 结构大纲 (提取类、方法、独立函数、模块注释章节)
     */
    private buildPythonOutlineTree(pyText: string): OutlineTreeNode[] {
        const flatNodes: OutlineTreeNode[] = [];
        const lines = pyText.split('\n');

        lines.forEach((line, idx) => {
            const trimmed = line.trim();
            if (!trimmed) return;

            // 1. 注释章节
            const commentMatch = trimmed.match(/^#\s*([#=*~-]+\s*)?([^\s=*~#-].*)$/);
            if (commentMatch && !trimmed.startsWith('# =') && !trimmed.startsWith('# -') && !trimmed.startsWith('# *')) {
                const commentTitle = commentMatch[2].trim();
                if (commentTitle.length > 1) {
                    flatNodes.push({
                        id: `py-sec-${idx + 1}`,
                        key: commentTitle,
                        path: commentTitle,
                        type: 'value',
                        preview: '章节',
                        level: 0,
                        lineNumber: idx + 1,
                        children: [],
                        isCollapsed: false
                    });
                    return;
                }
            }

            // 2. class ClassName(...)
            const classMatch = line.match(/^(\s*)class\s+([a-zA-Z0-9_]+)/);
            if (classMatch) {
                const indent = classMatch[1].length;
                const className = classMatch[2];
                const level = Math.floor(indent / 4);
                flatNodes.push({
                    id: `py-cls-${idx + 1}`,
                    key: className,
                    path: className,
                    type: 'object',
                    preview: 'class',
                    level,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
                return;
            }

            // 3. def func_name(...) / async def func_name(...)
            const defMatch = line.match(/^(\s*)(async\s+def|def)\s+([a-zA-Z0-9_]+)/);
            if (defMatch) {
                const indent = defMatch[1].length;
                const isAsync = defMatch[2].startsWith('async');
                const funcName = defMatch[3];
                const level = Math.floor(indent / 4);
                flatNodes.push({
                    id: `py-def-${idx + 1}`,
                    key: funcName,
                    path: funcName,
                    type: 'value',
                    preview: isAsync ? 'async def' : 'def',
                    level,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
            }
        });

        return this.constructTreeFromFlatNodes(flatNodes);
    }

    /**
     * 构建环境配置与点文件结构大纲 (.env, properties, conf 等)
     */
    private buildEnvOutlineTree(envText: string): OutlineTreeNode[] {
        const flatNodes: OutlineTreeNode[] = [];
        const lines = envText.split('\n');
        let currentSectionLevel = 0;

        lines.forEach((line, idx) => {
            const trimmed = line.trim();
            if (!trimmed) return;

            // 1. 注释章节分组
            const commentMatch = trimmed.match(/^#\s*([#=*~-]+\s*)?([^\s=*~#-].*)$/);
            if (commentMatch && !trimmed.startsWith('# =') && !trimmed.startsWith('# -') && !trimmed.startsWith('# *')) {
                const commentTitle = commentMatch[2].trim();
                if (commentTitle.length > 1) {
                    currentSectionLevel = 1;
                    flatNodes.push({
                        id: `env-sec-${idx + 1}`,
                        key: commentTitle,
                        path: commentTitle,
                        type: 'value',
                        preview: '分组',
                        level: 0,
                        lineNumber: idx + 1,
                        children: [],
                        isCollapsed: false
                    });
                    return;
                }
            }

            // 2. 键值对 KEY=VALUE 或 KEY: VALUE
            const kvMatch = trimmed.match(/^([a-zA-Z0-9_.-]+)\s*[:=]\s*(.*)$/);
            if (kvMatch) {
                const key = kvMatch[1];
                let val = kvMatch[2].trim().replace(/^["']|["']$/g, '');
                if (val.length > 24) val = val.substring(0, 22) + '...';
                flatNodes.push({
                    id: `env-kv-${idx + 1}`,
                    key,
                    path: key,
                    type: 'object',
                    preview: val || '=',
                    level: currentSectionLevel,
                    lineNumber: idx + 1,
                    children: [],
                    isCollapsed: false
                });
            }
        });

        return this.constructTreeFromFlatNodes(flatNodes);
    }

    /**
     * 基于层级 level 使用栈算法将平铺节点转换为标准嵌套树
     */
    private constructTreeFromFlatNodes(flatNodes: OutlineTreeNode[]): OutlineTreeNode[] {
        const rootNodes: OutlineTreeNode[] = [];
        const stack: OutlineTreeNode[] = [];

        flatNodes.forEach((node) => {
            while (stack.length > 0 && stack[stack.length - 1].level >= node.level) {
                stack.pop();
            }

            if (stack.length === 0) {
                rootNodes.push(node);
            } else {
                stack[stack.length - 1].children.push(node);
            }

            stack.push(node);
        });

        return rootNodes;
    }

    private openWithDefaultApp(): void {
        if (!this.currentFile) return;
        try {
            if ('openWithDefaultApp' in this.app) {
                (this.app as { openWithDefaultApp: (path: string) => void }).openWithDefaultApp(this.currentFile.path);
            }
        } catch {
            // ignore
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
            const fragment = createFragment();
            let lastIndex = 0;

            matches.forEach(({ index, length }) => {
                if (index > lastIndex) {
                    fragment.appendChild(document.createTextNode(text.substring(lastIndex, index)));
                }
                const mark = createEl('mark', {
                    cls: 'code-search-highlight',
                    text: text.substring(index, index + length)
                });
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
