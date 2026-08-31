import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import { init } from 'pptx-preview';

export const VIEW_TYPE_PPTX = 'pptx-preview-view';

export class PptxView extends FileView {
    private currentFile: TFile | null = null;
    private previewer: any = null;
    private viewMode: 'list' | 'slide' = 'list';
    private currentSlideIndex: number = 0;
    private totalSlides: number = 0;
    private zoomLevel: number = 1.0;
    private resizeObserver: ResizeObserver | null = null;

    // 搜索状态管理
    private searchMatches: HTMLElement[] = [];
    private currentMatchIndex: number = -1;
    private isSearchVisible: boolean = false;
    private searchInputEl: HTMLInputElement | null = null;
    private searchCountEl: HTMLElement | null = null;
    private searchContainerEl: HTMLElement | null = null;

    // 工具栏元素引用
    private pageIndicatorEl: HTMLElement | null = null;
    private prevSlideBtn: HTMLButtonElement | null = null;
    private nextSlideBtn: HTMLButtonElement | null = null;
    private zoomTextEl: HTMLElement | null = null;
    private pptxScaleWrapperEl: HTMLElement | null = null;

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
        this.currentSlideIndex = 0;
        this.totalSlides = 0;
        this.zoomLevel = 1.0;
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
        if (this.resizeObserver) {
            this.resizeObserver.disconnect();
            this.resizeObserver = null;
        }
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
        // OLE2 复合文件特征魔数: D0 CF 11 E0 A1 B1 1A E1
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

        // 视图模式切换：列表 / 单页
        const modeListBtn = controls.createEl('button', {
            cls: `office-preview-btn ${this.viewMode === 'list' ? 'is-active' : ''}`,
            text: '📜 列表'
        });
        const modeSlideBtn = controls.createEl('button', {
            cls: `office-preview-btn ${this.viewMode === 'slide' ? 'is-active' : ''}`,
            text: '📽️ 单页'
        });

        // 翻页控件组（单页模式专属）
        const navGroup = controls.createEl('div', {
            cls: `office-preview-pptx-nav-group ${this.viewMode === 'slide' ? '' : 'is-hidden'}`
        });
        this.prevSlideBtn = navGroup.createEl('button', { cls: 'office-preview-btn', text: '◀ 上一页' });
        this.pageIndicatorEl = navGroup.createEl('span', { cls: 'office-preview-page-indicator', text: '1 / 1' });
        this.nextSlideBtn = navGroup.createEl('button', { cls: 'office-preview-btn', text: '下一页 ▶' });

        // 缩放控件组
        const zoomGroup = controls.createEl('div', { cls: 'office-preview-zoom-group' });
        const zoomOutBtn = zoomGroup.createEl('button', { cls: 'office-preview-btn', text: '➖' });
        this.zoomTextEl = zoomGroup.createEl('span', { cls: 'office-preview-zoom-text', text: '100%' });
        const zoomInBtn = zoomGroup.createEl('button', { cls: 'office-preview-btn', text: '➕' });
        const zoomResetBtn = zoomGroup.createEl('button', { cls: 'office-preview-btn', text: '适应' });

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

        // 3. PPT 渲染主视口与滚动包裹容器
        const renderWrapper = container.createEl('div', { cls: 'office-preview-pptx-wrapper' });
        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: '正在解析 PowerPoint 幻灯片...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            // 检测是否为旧版二进制 .ppt 文件
            if (this.isLegacyPpt(arrayBuffer)) {
                this.renderLegacyPptNotice(renderWrapper);
                return;
            }

            // 创建缩放层与内容挂载容器
            this.pptxScaleWrapperEl = renderWrapper.createEl('div', { cls: 'pptx-scale-wrapper' });
            const pptxRenderEl = this.pptxScaleWrapperEl.createEl('div', { cls: 'pptx-render-mount' });

            // 初始化 PPTX 预览器
            const baseWidth = 960;
            const baseHeight = 540;
            this.previewer = init(pptxRenderEl, {
                width: baseWidth,
                height: baseHeight,
                mode: this.viewMode
            });

            await this.previewer.preview(arrayBuffer);

            // 获取总幻灯片数
            this.totalSlides = this.previewer.slideCount || 1;
            this.updatePaginationDisplay();

            // 模式切换事件
            modeListBtn.onclick = () => {
                if (this.viewMode === 'list') return;
                this.viewMode = 'list';
                this.renderPptx();
            };

            modeSlideBtn.onclick = () => {
                if (this.viewMode === 'slide') return;
                this.viewMode = 'slide';
                this.renderPptx();
            };

            // 单页模式下的翻页逻辑
            this.prevSlideBtn.onclick = () => this.goToPrevSlide();
            this.nextSlideBtn.onclick = () => this.goToNextSlide();

            // 缩放逻辑
            zoomInBtn.onclick = () => this.setZoom(this.zoomLevel + 0.1);
            zoomOutBtn.onclick = () => this.setZoom(this.zoomLevel - 0.1);
            zoomResetBtn.onclick = () => this.autoFitZoom(renderWrapper);

            // 全屏演示
            fullscreenBtn.onclick = () => {
                if (!document.fullscreenElement) {
                    renderWrapper.requestFullscreen().catch(err => {
                        console.warn('Fullscreen error:', err);
                    });
                } else {
                    document.exitFullscreen();
                }
            };

            // 自动适配初始缩放比例
            this.autoFitZoom(renderWrapper);

            // 监听视口大小变化
            this.resizeObserver = new ResizeObserver(() => {
                this.autoFitZoom(renderWrapper);
            });
            this.resizeObserver.observe(renderWrapper);

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

            // 键盘事件监听 (快捷键)
            container.tabIndex = 0;
            container.onkeydown = (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, pptxRenderEl);
                } else if (this.viewMode === 'slide' && !this.isSearchVisible) {
                    if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ') {
                        e.preventDefault();
                        this.goToNextSlide();
                    } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
                        e.preventDefault();
                        this.goToPrevSlide();
                    }
                }
            };

        } catch (error) {
            loadingEl.remove();
            // 如果解析失败，检测是否为 .ppt 给予智能引导
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

    private setZoom(newZoom: number): void {
        this.zoomLevel = Math.max(0.3, Math.min(2.5, newZoom));
        if (this.pptxScaleWrapperEl) {
            this.pptxScaleWrapperEl.style.transform = `scale(${this.zoomLevel})`;
        }
        if (this.zoomTextEl) {
            this.zoomTextEl.setText(`${Math.round(this.zoomLevel * 100)}%`);
        }
    }

    private autoFitZoom(wrapper: HTMLElement): void {
        if (!this.pptxScaleWrapperEl) return;
        const containerWidth = wrapper.clientWidth - 48;
        const targetWidth = 960;
        if (containerWidth > 0) {
            const fitScale = Math.min(1.2, Math.max(0.4, containerWidth / targetWidth));
            this.setZoom(fitScale);
        }
    }

    private goToPrevSlide(): void {
        if (this.previewer && this.currentSlideIndex > 0) {
            this.currentSlideIndex--;
            if (typeof this.previewer.renderPreSlide === 'function') {
                this.previewer.renderPreSlide();
            } else if (typeof this.previewer.renderSingleSlide === 'function') {
                this.previewer.renderSingleSlide(this.currentSlideIndex);
            }
            this.updatePaginationDisplay();
        }
    }

    private goToNextSlide(): void {
        if (this.previewer && this.currentSlideIndex < this.totalSlides - 1) {
            this.currentSlideIndex++;
            if (typeof this.previewer.renderNextSlide === 'function') {
                this.previewer.renderNextSlide();
            } else if (typeof this.previewer.renderSingleSlide === 'function') {
                this.previewer.renderSingleSlide(this.currentSlideIndex);
            }
            this.updatePaginationDisplay();
        }
    }

    private updatePaginationDisplay(): void {
        if (this.pageIndicatorEl) {
            this.pageIndicatorEl.setText(`${this.currentSlideIndex + 1} / ${this.totalSlides}`);
        }
        if (this.prevSlideBtn) {
            this.prevSlideBtn.disabled = this.currentSlideIndex <= 0;
        }
        if (this.nextSlideBtn) {
            this.nextSlideBtn.disabled = this.currentSlideIndex >= this.totalSlides - 1;
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
