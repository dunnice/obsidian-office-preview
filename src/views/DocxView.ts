import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import { renderAsync } from 'docx-preview';

export const VIEW_TYPE_DOCX = 'docx-preview-view';

export class DocxView extends FileView {
    private currentFile: TFile | null = null;
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
        return VIEW_TYPE_DOCX;
    }

    getDisplayText(): string {
        return this.currentFile ? this.currentFile.name : 'Word Preview';
    }

    getIcon(): string {
        return 'file-text';
    }

    async onLoadFile(file: TFile): Promise<void> {
        this.currentFile = file;
        this.searchMatches = [];
        this.currentMatchIndex = -1;
        this.isSearchVisible = false;
        await this.renderDocx();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.currentFile = null;
        this.searchMatches = [];
        this.contentEl.empty();
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

    private async renderDocx(): Promise<void> {
        if (!this.currentFile) return;

        const container = this.contentEl;
        container.empty();
        container.addClass('office-preview-docx-container');

        // 1. 顶部主工具栏
        const toolbar = container.createEl('div', { cls: 'office-preview-toolbar' });
        
        const titleContainer = toolbar.createEl('div', { cls: 'office-preview-title-container' });
        titleContainer.createEl('span', { 
            cls: 'office-preview-title', 
            text: `📄 ${this.currentFile.name}` 
        });

        // 浅色小字修改时间
        if (this.currentFile.stat && this.currentFile.stat.mtime) {
            const timeStr = this.formatDate(this.currentFile.stat.mtime);
            titleContainer.createEl('span', {
                cls: 'office-preview-mtime',
                text: `修改时间: ${timeStr}`
            });
        }

        const controls = toolbar.createEl('div', { cls: 'office-preview-controls' });

        // 搜索触发按钮
        const toggleSearchBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '🔍 查找' });

        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderDocx();

        // 2. 悬浮搜索控制面板
        this.searchContainerEl = container.createEl('div', { cls: 'office-preview-search-bar is-hidden' });
        
        this.searchInputEl = this.searchContainerEl.createEl('input', {
            cls: 'office-preview-search-input',
            type: 'text',
            placeholder: '输入关键词搜索 (Cmd+F)...'
        });

        this.searchCountEl = this.searchContainerEl.createEl('span', { cls: 'office-preview-search-count', text: '0 / 0' });

        const prevMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔼' });
        const nextMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔽' });
        const closeSearchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '✖' });

        // 3. 自然连续的 Word 阅读视图包裹容器
        const renderWrapper = container.createEl('div', { cls: 'office-preview-docx-wrapper' });
        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: '正在渲染 Word 文档...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            const docxContainer = renderWrapper.createEl('div', { cls: 'docx-render-body' });

            await renderAsync(arrayBuffer, docxContainer, undefined, {
                className: 'docx',
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreFonts: false,
                breakPages: false,
                experimental: true,
                useBase64URL: true,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true,
            });

            // 搜索逻辑注册
            const performSearch = (query: string) => {
                this.clearSearchHighlights(docxContainer);
                if (!query || query.trim() === '') {
                    this.searchMatches = [];
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                    return;
                }

                // 遍历搜寻文本节点
                this.highlightText(docxContainer, query.trim().toLowerCase());
                this.searchMatches = Array.from(docxContainer.querySelectorAll('.docx-search-highlight'));
                
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
                    this.toggleSearch(false, docxContainer);
                }
            };

            prevMatchBtn.onclick = () => this.navigateMatch(-1);
            nextMatchBtn.onclick = () => this.navigateMatch(1);
            closeSearchBtn.onclick = () => this.toggleSearch(false, docxContainer);
            toggleSearchBtn.onclick = () => this.toggleSearch(!this.isSearchVisible, docxContainer);

            // 快捷键 Cmd+F / Ctrl+F 触发
            container.onkeydown = (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, docxContainer);
                }
            };

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createEl('div', { 
                cls: 'office-preview-error', 
                text: `解析 Word 文档失败: ${error instanceof Error ? error.message : String(error)}` 
            });
        }
    }

    private toggleSearch(show: boolean, docxContainer: HTMLElement): void {
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
                this.clearSearchHighlights(docxContainer);
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
        const highlights = root.querySelectorAll('.docx-search-highlight');
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
            
            if (lowerText.includes(query) && textNode.parentElement && !textNode.parentElement.hasClass('docx-search-highlight')) {
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
                const mark = createEl('mark', {
                    cls: 'office-preview-search-highlight',
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
