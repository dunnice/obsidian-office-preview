import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import { renderAsync } from 'docx-preview';

export const VIEW_TYPE_DOCX = 'docx-preview-view';

export class DocxView extends FileView {
    private currentFile: TFile | null = null;

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
        await this.renderDocx();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.currentFile = null;
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
        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderDocx();

        // 2. 自然连续的 Word 阅读视图包裹容器
        const renderWrapper = container.createEl('div', { cls: 'office-preview-docx-wrapper' });
        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: '正在渲染 Word 文档...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            const docxContainer = renderWrapper.createEl('div', { cls: 'docx-render-body' });

            // 渲染配置：自然连续排版
            await renderAsync(arrayBuffer, docxContainer, undefined, {
                className: 'docx',
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreFonts: false,
                breakPages: false, // 禁用强制分页，回归自然连续的文档流
                experimental: true,
                useBase64URL: true,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true,
            });

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createEl('div', { 
                cls: 'office-preview-error', 
                text: `解析 Word 文档失败: ${error instanceof Error ? error.message : String(error)}` 
            });
        }
    }
}
