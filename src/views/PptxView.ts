import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import { init } from 'pptx-preview';

export const VIEW_TYPE_PPTX = 'pptx-preview-view';

export class PptxView extends FileView {
    private currentFile: TFile | null = null;
    private previewer: any = null;

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
        return VIEW_TYPE_PPTX;
    }

    getDisplayText(): string {
        return this.currentFile ? this.currentFile.name : 'PowerPoint Preview';
    }

    getIcon(): string {
        return 'presentation';
    }

    async onLoadFile(file: TFile): Promise<void> {
        this.currentFile = file;
        this.searchMatches = [];
        this.currentMatchIndex = -1;
        this.isSearchVisible = false;
        await this.renderPptx();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.cleanup();
        this.currentFile = null;
        this.contentEl.empty();
    }

    private cleanup(): void {
        if (this.previewer && typeof this.previewer.destroy === 'function') {
            try {
                this.previewer.destroy();
            } catch (e) {
                console.warn('Error destroying pptx previewer:', e);
            }
        }
        this.previewer = null;
        this.searchMatches = [];
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

    /**
     * 判断是否为旧版 OLE2 复合二进制 .ppt 文件
     */
    private isLegacyPpt(arrayBuffer: ArrayBuffer): boolean {
        if (arrayBuffer.byteLength < 8) return false;
        const uint8 = new Uint8Array(arrayBuffer.slice(0, 8));
        const ole2Magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
        return ole2Magic.every((byte, idx) => uint8[idx] === byte);
    }

    private async renderPptx(): Promise<void> {
        if (!this.currentFile) return;

        this.cleanup();
        const container = this.contentEl;
        container.empty();
        container.addClass('office-preview-pptx-container');

        // 1. 顶部主工具栏
        const toolbar = container.createEl('div', { cls: 'office-preview-toolbar' });

        const titleContainer = toolbar.createEl('div', { cls: 'office-preview-title-container' });
        titleContainer.createEl('span', {
            cls: 'office-preview-title',
            text: `📽️ ${this.currentFile.name}`
        });

        if (this.currentFile.stat && this.currentFile.stat.mtime) {
            const timeStr = this.formatDate(this.currentFile.stat.mtime);
            titleContainer.createEl('span', {
                cls: 'office-preview-mtime',
                text: `修改时间: ${timeStr}`
            });
        }

        const controls = toolbar.createEl('div', { cls: 'office-preview-controls' });

        // 全屏演示按钮
        const fullscreenBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '🖥️ 全屏' });

        // 搜索触发按钮
        const toggleSearchBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '🔍 查找' });

        // 外部应用打开
        const openExternalBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '↗ 外部打开' });
        openExternalBtn.onclick = () => this.openWithDefaultApp();

        // 刷新按钮
        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderPptx();

        // 2. 悬浮搜索控制面板
        this.searchContainerEl = container.createEl('div', { cls: 'office-preview-search-bar is-hidden' });

        this.searchInputEl = this.searchContainerEl.createEl('input', {
            cls: 'office-preview-search-input',
            type: 'text',
            placeholder: '搜索幻灯片内容 (Cmd+F)...'
        });

        this.searchCountEl = this.searchContainerEl.createEl('span', { cls: 'office-preview-search-count', text: '0 / 0' });

        const prevMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔼' });
        const nextMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔽' });
        const closeSearchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '✖' });

        // 3. 类似 Word 的自然长卷滚动视口
        const renderWrapper = container.createEl('div', { cls: 'office-preview-pptx-wrapper' });
        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: '正在解析并流式平铺 PowerPoint 幻灯片...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            if (this.isLegacyPpt(arrayBuffer)) {
                this.renderLegacyPptNotice(renderWrapper);
                return;
            }

            // 内容挂载容器
            const pptxRenderEl = renderWrapper.createEl('div', { cls: 'pptx-render-mount' });

            // 动态计算适合视口宽度的单页渲染基准宽（不设置 height，彻底禁用定高，让高度按幻灯片张数自然向下铺开）
            const availableWidth = Math.max(renderWrapper.clientWidth - 64, 600);
            const renderWidth = Math.min(availableWidth, 980);

            // 初始化预览器：不传递 height，这样 pptx-preview 不会定高，每张幻灯片一页页平铺向下！
            this.previewer = init(pptxRenderEl, {
                width: renderWidth,
                mode: 'list'
            });

            await this.previewer.preview(arrayBuffer);

            // 全屏演示逻辑
            fullscreenBtn.onclick = () => {
                if (!document.fullscreenElement) {
                    renderWrapper.requestFullscreen().catch(err => {
                        console.warn('Fullscreen error:', err);
                    });
                } else {
                    document.exitFullscreen();
                }
            };

            // 搜索逻辑注册
            const performPptxSearch = (query: string) => {
                this.clearSearchHighlights(pptxRenderEl);
                if (!query || query.trim() === '') {
                    this.searchMatches = [];
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                    return;
                }

                this.highlightText(pptxRenderEl, query.trim().toLowerCase());
                this.searchMatches = Array.from(pptxRenderEl.querySelectorAll('.pptx-search-highlight'));

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
                    performPptxSearch(this.searchInputEl?.value || '');
                }, 200);
            };

            this.searchInputEl.onkeydown = (e: KeyboardEvent) => {
                if (e.key === 'Enter') {
                    if (e.shiftKey) this.navigateMatch(-1);
                    else this.navigateMatch(1);
                } else if (e.key === 'Escape') {
                    this.toggleSearch(false, pptxRenderEl);
                }
            };

            prevMatchBtn.onclick = () => this.navigateMatch(-1);
            nextMatchBtn.onclick = () => this.navigateMatch(1);
            closeSearchBtn.onclick = () => this.toggleSearch(false, pptxRenderEl);
            toggleSearchBtn.onclick = () => this.toggleSearch(!this.isSearchVisible, pptxRenderEl);

            // 快捷键监听
            container.tabIndex = 0;
            container.onkeydown = (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, pptxRenderEl);
                }
            };

        } catch (error) {
            loadingEl.remove();
            if (this.currentFile.extension.toLowerCase() === 'ppt') {
                this.renderLegacyPptNotice(renderWrapper);
            } else {
                renderWrapper.createEl('div', {
                    cls: 'office-preview-error',
                    text: `解析 PowerPoint 演示文稿失败: ${error instanceof Error ? error.message : String(error)}`
                });
            }
        }
    }

    /**
     * 旧版二进制 .ppt 文件的友好提示与引导界面
     */
    private renderLegacyPptNotice(wrapper: HTMLElement): void {
        wrapper.empty();
        const card = wrapper.createEl('div', { cls: 'office-preview-legacy-card' });

        card.createEl('div', { cls: 'office-preview-legacy-icon', text: '📽️' });
        card.createEl('h3', { cls: 'office-preview-legacy-title', text: 'PowerPoint 97-2003 演示文稿 (.ppt)' });
        card.createEl('p', {
            cls: 'office-preview-legacy-desc',
            text: '当前文件为早期 Office 二进制复合格式 (.ppt)。由于该格式属于专有二进制结构，建议使用系统默认应用（Microsoft PowerPoint / WPS / Keynote）打开，或将其另存为现代 .pptx 格式以获得完整的内嵌原生预览支持。'
        });

        const btnGroup = card.createEl('div', { cls: 'office-preview-legacy-actions' });
        const openBtn = btnGroup.createEl('button', {
            cls: 'office-preview-btn is-primary',
            text: '🚀 使用系统默认应用打开'
        });
        openBtn.onclick = () => this.openWithDefaultApp();

        const reloadBtn = btnGroup.createEl('button', {
            cls: 'office-preview-btn',
            text: '🔄 重新检测'
        });
        reloadBtn.onclick = () => this.renderPptx();
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

    private toggleSearch(show: boolean, pptxContainer: HTMLElement): void {
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
                this.clearSearchHighlights(pptxContainer);
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
        const highlights = root.querySelectorAll('.pptx-search-highlight');
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

            if (lowerText.includes(query) && textNode.parentElement && !textNode.parentElement.hasClass('pptx-search-highlight')) {
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
                mark.className = 'pptx-search-highlight';
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
