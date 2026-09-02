import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import * as XLSX from 'xlsx';

export const VIEW_TYPE_XLSX = 'xlsx-preview-view';

export class XlsxView extends FileView {
    private currentFile: TFile | null = null;
    private workbook: XLSX.WorkBook | null = null;
    private currentSheetIndex: number = 0;

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
        return VIEW_TYPE_XLSX;
    }

    getDisplayText(): string {
        return this.currentFile ? this.currentFile.name : 'Excel Preview';
    }

    getIcon(): string {
        return 'table';
    }

    async onLoadFile(file: TFile): Promise<void> {
        this.currentFile = file;
        this.currentSheetIndex = 0;
        this.searchMatches = [];
        this.currentMatchIndex = -1;
        this.isSearchVisible = false;
        await this.renderXlsx();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.currentFile = null;
        this.workbook = null;
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

    private async renderXlsx(): Promise<void> {
        if (!this.currentFile) return;

        const container = this.contentEl;
        container.empty();
        container.addClass('office-preview-xlsx-container');

        // 1. 顶部主工具栏
        const toolbar = container.createDiv({ cls: 'office-preview-toolbar' });
        
        const titleContainer = toolbar.createDiv({ cls: 'office-preview-title-container' });
        titleContainer.createSpan({ 
            cls: 'office-preview-title', 
            text: `📊 ${this.currentFile.name}` 
        });

        if (this.currentFile.stat && this.currentFile.stat.mtime) {
            const timeStr = this.formatDate(this.currentFile.stat.mtime);
            titleContainer.createSpan({
                cls: 'office-preview-mtime',
                text: `修改时间: ${timeStr}`
            });
        }

        const controls = toolbar.createDiv({ cls: 'office-preview-controls' });

        // 搜索触发按钮
        const toggleSearchBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '🔍 查找' });

        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderXlsx();

        // 2. 悬浮搜索控制面板
        this.searchContainerEl = container.createDiv({ cls: 'office-preview-search-bar is-hidden' });
        
        this.searchInputEl = this.searchContainerEl.createEl('input', {
            cls: 'office-preview-search-input',
            type: 'text',
            placeholder: '搜索单元格内容 (Cmd+F)...'
        });

        this.searchCountEl = this.searchContainerEl.createSpan({ cls: 'office-preview-search-count', text: '0 / 0' });

        const prevMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔼' });
        const nextMatchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '🔽' });
        const closeSearchBtn = this.searchContainerEl.createEl('button', { cls: 'office-preview-search-btn', text: '✖' });

        // 3. 多 Sheet 标签页切换栏
        const tabsBar = container.createDiv({ cls: 'office-preview-sheet-tabs-bar' });

        // 4. 表格内容包裹容器
        const renderWrapper = container.createDiv({ cls: 'office-preview-xlsx-wrapper' });
        const loadingEl = renderWrapper.createDiv({ cls: 'office-preview-loading', text: '正在解析 Excel 工作簿...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            this.workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true });

            if (!this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
                renderWrapper.createDiv({ cls: 'office-preview-error', text: 'Excel 文件未包含任何工作表 (Sheet)' });
                return;
            }

            // 渲染 Sheet 切换 Tab 按钮
            tabsBar.empty();
            this.workbook.SheetNames.forEach((sheetName, index) => {
                const tabBtn = tabsBar.createEl('button', {
                    cls: `office-preview-sheet-tab ${index === this.currentSheetIndex ? 'is-active' : ''}`,
                    text: sheetName
                });

                tabBtn.onclick = () => {
                    if (this.currentSheetIndex === index) return;
                    this.currentSheetIndex = index;
                    
                    tabsBar.querySelectorAll('.office-preview-sheet-tab').forEach((el, i) => {
                        if (i === index) el.addClass('is-active');
                        else el.removeClass('is-active');
                    });

                    this.loadCurrentSheetHtml(renderWrapper);
                };
            });

            // 初始加载当前激活的 Sheet
            this.loadCurrentSheetHtml(renderWrapper);

            // 搜索逻辑注册
            const performExcelSearch = (query: string) => {
                this.clearTableHighlights(renderWrapper);
                if (!query || query.trim() === '') {
                    this.searchMatches = [];
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                    return;
                }

                const lowerQuery = query.trim().toLowerCase();
                const cells = Array.from(renderWrapper.querySelectorAll<HTMLElement>('.excel-cell'));
                this.searchMatches = cells.filter(cell => cell.textContent?.toLowerCase().includes(lowerQuery));

                this.searchMatches.forEach(cell => cell.addClass('excel-search-highlight'));

                if (this.searchMatches.length > 0) {
                    this.currentMatchIndex = 0;
                    this.focusExcelMatch(0);
                } else {
                    this.currentMatchIndex = -1;
                    if (this.searchCountEl) this.searchCountEl.setText('0 / 0');
                }
            };

            let searchTimer: number | null = null;
            this.searchInputEl.oninput = () => {
                if (searchTimer !== null) window.clearTimeout(searchTimer);
                searchTimer = window.setTimeout(() => {
                    performExcelSearch(this.searchInputEl?.value || '');
                }, 150);
            };

            this.searchInputEl.onkeydown = (e: KeyboardEvent) => {
                if (e.key === 'Enter') {
                    if (e.shiftKey) this.navigateExcelMatch(-1);
                    else this.navigateExcelMatch(1);
                } else if (e.key === 'Escape') {
                    this.toggleSearch(false, renderWrapper);
                }
            };

            prevMatchBtn.onclick = () => this.navigateExcelMatch(-1);
            nextMatchBtn.onclick = () => this.navigateExcelMatch(1);
            closeSearchBtn.onclick = () => this.toggleSearch(false, renderWrapper);
            toggleSearchBtn.onclick = () => this.toggleSearch(!this.isSearchVisible, renderWrapper);

            // 快捷键监听
            container.tabIndex = 0;
            container.onkeydown = (e: KeyboardEvent) => {
                if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
                    e.preventDefault();
                    this.toggleSearch(true, renderWrapper);
                }
            };

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createDiv({ 
                cls: 'office-preview-error', 
                text: `解析 Excel 表格失败: ${error instanceof Error ? error.message : String(error)}` 
            });
        }
    }

    private toggleSearch(show: boolean, renderWrapper: HTMLElement): void {
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
                this.clearTableHighlights(renderWrapper);
            }
        }
    }

    private navigateExcelMatch(direction: number): void {
        if (this.searchMatches.length === 0) return;
        this.currentMatchIndex += direction;
        if (this.currentMatchIndex < 0) this.currentMatchIndex = this.searchMatches.length - 1;
        if (this.currentMatchIndex >= this.searchMatches.length) this.currentMatchIndex = 0;
        this.focusExcelMatch(this.currentMatchIndex);
    }

    private focusExcelMatch(index: number): void {
        this.searchMatches.forEach((el, i) => {
            if (i === index) el.addClass('is-active-match');
            else el.removeClass('is-active-match');
        });

        const activeCell = this.searchMatches[index];
        if (activeCell) {
            activeCell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        }

        if (this.searchCountEl) {
            this.searchCountEl.setText(`${index + 1} / ${this.searchMatches.length}`);
        }
    }

    private clearTableHighlights(renderWrapper: HTMLElement): void {
        const matches = renderWrapper.querySelectorAll('.excel-search-highlight');
        matches.forEach(el => {
            el.removeClass('excel-search-highlight');
            el.removeClass('is-active-match');
        });
    }

    private loadCurrentSheetHtml(renderWrapper: HTMLElement): void {
        if (!this.workbook || !this.workbook.SheetNames) return;

        const sheetName = this.workbook.SheetNames[this.currentSheetIndex] || this.workbook.SheetNames[0];
        const ws = this.workbook.Sheets[sheetName];

        renderWrapper.empty();

        if (!ws) {
            renderWrapper.createDiv({ cls: 'office-preview-error', text: '该工作表无有效数据' });
            return;
        }

        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

        const tableContainer = renderWrapper.createDiv({ cls: 'excel-table-container' });
        const table = tableContainer.createEl('table', { cls: 'excel-table' });

        const thead = table.createEl('thead');
        const headerRow = thead.createEl('tr');
        
        headerRow.createEl('th', { cls: 'excel-corner-header', text: '' });

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const colName = XLSX.utils.encode_col(C);
            headerRow.createEl('th', { cls: 'excel-col-header', text: colName });
        }

        const tbody = table.createEl('tbody');

        for (let R = range.s.r; R <= range.e.r; ++R) {
            const tr = tbody.createEl('tr');
            
            tr.createEl('td', { cls: 'excel-row-header', text: String(R + 1) });

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellRef] as XLSX.CellObject | undefined;
                
                let cellText = '';
                if (cell !== undefined && cell.v !== undefined && cell.v !== null) {
                    cellText = XLSX.utils.format_cell(cell);
                    if (cellText === undefined) {
                        cellText = String(cell.v);
                    }
                }

                const td = tr.createEl('td', { cls: 'excel-cell', text: cellText });
                
                if (!cellText) {
                    td.addClass('is-empty');
                }
            }
        }

        // 如果搜索框处于开启状态，重新应用搜索高亮
        if (this.isSearchVisible && this.searchInputEl && this.searchInputEl.value) {
            const lowerQuery = this.searchInputEl.value.trim().toLowerCase();
            if (lowerQuery) {
                const cells = Array.from(renderWrapper.querySelectorAll<HTMLElement>('.excel-cell'));
                this.searchMatches = cells.filter(c => c.textContent?.toLowerCase().includes(lowerQuery));
                this.searchMatches.forEach(c => c.addClass('excel-search-highlight'));
                if (this.searchMatches.length > 0) {
                    this.currentMatchIndex = 0;
                    this.focusExcelMatch(0);
                }
            }
        }
    }
}
