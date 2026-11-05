import { Plugin } from 'obsidian';
import { DocxView, VIEW_TYPE_DOCX } from './views/DocxView';
import { XlsxView, VIEW_TYPE_XLSX } from './views/XlsxView';
import { PptxView, VIEW_TYPE_PPTX } from './views/PptxView';
import { CodeView, VIEW_TYPE_CODE } from './views/CodeView';
import { DotfileSuggestModal } from './modals/DotfileSuggestModal';

export default class OfficePreviewPlugin extends Plugin {
    async onload() {
        // 注册 Word 视图与扩展名绑定
        this.registerView(
            VIEW_TYPE_DOCX,
            (leaf) => new DocxView(leaf)
        );

        try {
            this.registerExtensions(['docx'], VIEW_TYPE_DOCX);
        } catch (e) {
            console.warn('Docx extension already registered or conflict:', e);
        }

        // 注册 Excel 视图与扩展名绑定
        this.registerView(
            VIEW_TYPE_XLSX,
            (leaf) => new XlsxView(leaf)
        );

        try {
            this.registerExtensions(['xlsx', 'xls', 'csv'], VIEW_TYPE_XLSX);
        } catch (e) {
            console.warn('Xlsx/Xls/Csv extension already registered or conflict:', e);
        }

        // 注册 PowerPoint 视图与扩展名绑定
        this.registerView(
            VIEW_TYPE_PPTX,
            (leaf) => new PptxView(leaf)
        );

        try {
            this.registerExtensions(['pptx', 'ppt'], VIEW_TYPE_PPTX);
        } catch (e) {
            console.warn('Pptx/Ppt extension already registered or conflict:', e);
        }

        // 注册结构化代码、Python 与文本/隐藏配置文件视图
        this.registerView(
            VIEW_TYPE_CODE,
            (leaf) => new CodeView(leaf)
        );

        const codeExtensions = [
            'json', 'sql', 'yaml', 'yml', 'txt',
            // Python 支持
            'py', 'pyw',
            // 隐藏环境配置与点文件扩展名
            'env', 'local', 'development', 'production', 'staging', 'test',
            'gitignore', 'dockerignore', 'editorconfig', 'npmignore', 'npmrc',
            'conf', 'ini', 'properties', 'toml', 'cfg', 'sh', 'bash'
        ];

        for (const ext of codeExtensions) {
            try {
                this.registerExtensions([ext], VIEW_TYPE_CODE);
            } catch (e) {
                console.warn(`Code extension (.${ext}) registration warning:`, e);
            }
        }

        // 注册快捷命令：快速浏览、打开或创建隐藏文件 (.env, .gitignore, .py 等)
        this.addCommand({
            id: 'open-dotfile',
            name: '浏览或创建隐藏与配置文件 (.env, .gitignore, .py 等)',
            callback: () => {
                new DotfileSuggestModal(this.app).open();
            }
        });
    }

    onunload() {
    }
}
