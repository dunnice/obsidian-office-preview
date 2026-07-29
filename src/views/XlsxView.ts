import { FileView, WorkspaceLeaf, TFile } from 'obsidian';
import * as XLSX from 'xlsx';

export const VIEW_TYPE_XLSX = 'xlsx-preview-view';

export class XlsxView extends FileView {
    private currentFile: TFile | null = null;
    private workbook: XLSX.WorkBook | null = null;
    private currentSheetIndex: number = 0;

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
        await this.renderXlsx();
    }

    async onUnloadFile(file: TFile): Promise<void> {
        this.currentFile = null;
        this.workbook = null;
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
        const toolbar = container.createEl('div', { cls: 'office-preview-toolbar' });
        
        const titleContainer = toolbar.createEl('div', { cls: 'office-preview-title-container' });
        titleContainer.createEl('span', { 
            cls: 'office-preview-title', 
            text: `📊 ${this.currentFile.name}` 
        });

        // 格式化并展示修改时间
        if (this.currentFile.stat && this.currentFile.stat.mtime) {
            const timeStr = this.formatDate(this.currentFile.stat.mtime);
            titleContainer.createEl('span', {
                cls: 'office-preview-mtime',
                text: `修改时间: ${timeStr}`
            });
        }

        const controls = toolbar.createEl('div', { cls: 'office-preview-controls' });
        const reloadBtn = controls.createEl('button', { cls: 'office-preview-btn', text: '刷新' });
        reloadBtn.onclick = () => this.renderXlsx();

        // 2. 多 Sheet 标签页切换栏
        const tabsBar = container.createEl('div', { cls: 'office-preview-sheet-tabs-bar' });

        // 3. 表格内容包裹容器（带原生纵横滚动条）
        const renderWrapper = container.createEl('div', { cls: 'office-preview-xlsx-wrapper' });
        const loadingEl = renderWrapper.createEl('div', { cls: 'office-preview-loading', text: '正在解析 Excel 工作簿...' });

        try {
            const arrayBuffer = await this.app.vault.readBinary(this.currentFile);
            loadingEl.remove();

            this.workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true });

            if (!this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
                renderWrapper.createEl('div', { cls: 'office-preview-error', text: 'Excel 文件未包含任何工作表 (Sheet)' });
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

            // 首次展示第 0 个 Sheet
            this.loadCurrentSheetHtml(renderWrapper);

        } catch (error) {
            loadingEl.remove();
            renderWrapper.createEl('div', { 
                cls: 'office-preview-error', 
                text: `解析 Excel 表格失败: ${error instanceof Error ? error.message : String(error)}` 
            });
        }
    }

    private loadCurrentSheetHtml(renderWrapper: HTMLElement): void {
        if (!this.workbook || !this.workbook.SheetNames) return;

        const sheetName = this.workbook.SheetNames[this.currentSheetIndex] || this.workbook.SheetNames[0];
        const ws = this.workbook.Sheets[sheetName];

        renderWrapper.empty();

        if (!ws) {
            renderWrapper.createEl('div', { cls: 'office-preview-error', text: '该工作表无有效数据' });
            return;
        }

        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

        const tableContainer = renderWrapper.createEl('div', { cls: 'excel-table-container' });
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
                const cell = ws[cellRef];
                
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
    }
}
