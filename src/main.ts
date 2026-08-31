import { Plugin } from 'obsidian';
import { DocxView, VIEW_TYPE_DOCX } from './views/DocxView';
import { XlsxView, VIEW_TYPE_XLSX } from './views/XlsxView';
import { PptxView, VIEW_TYPE_PPTX } from './views/PptxView';
import { CodeView, VIEW_TYPE_CODE } from './views/CodeView';

export default class OfficePreviewPlugin extends Plugin {
    async onload() {
        console.log('加载 Obsidian Office Preview 插件...');

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

        // 注册结构化代码与配置文件视图 (json, sql, yaml, yml)
        this.registerView(
            VIEW_TYPE_CODE,
            (leaf) => new CodeView(leaf)
        );

        try {
            this.registerExtensions(['json', 'sql', 'yaml', 'yml'], VIEW_TYPE_CODE);
        } catch (e) {
            console.warn('Code extensions (json, sql, yaml, yml) already registered or conflict:', e);
        }
    }

    onunload() {
        console.log('卸载 Obsidian Office Preview 插件');
    }
}
