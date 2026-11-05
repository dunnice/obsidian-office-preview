import { App, FuzzySuggestModal, TFile, Notice } from 'obsidian';
import { VIEW_TYPE_CODE } from '../views/CodeView';

export interface FileItem {
    path: string;
    isNew?: boolean;
}

export class DotfileSuggestModal extends FuzzySuggestModal<FileItem> {
    private items: FileItem[] = [];

    constructor(app: App) {
        super(app);
        this.setPlaceholder('输入文件名搜索或创建隐藏文件 (例如: .env, .env.local, .gitignore, script.py)...');
    }

    async onOpen() {
        super.onOpen();
        await this.scanDotfiles();
    }

    private async scanDotfiles(): Promise<void> {
        const foundPaths: string[] = [];
        const adapter = this.app.vault.adapter;

        const scanDir = async (dir: string) => {
            try {
                const list = await adapter.list(dir);
                for (const file of list.files) {
                    const filename = file.split('/').pop() || '';
                    if (
                        filename.startsWith('.') ||
                        filename.endsWith('.py') ||
                        filename.endsWith('.pyw') ||
                        filename.endsWith('.env') ||
                        filename.includes('.env.') ||
                        filename.endsWith('.conf') ||
                        filename.endsWith('.ini') ||
                        filename.endsWith('.properties') ||
                        filename.endsWith('.toml') ||
                        filename.endsWith('.cfg') ||
                        filename === 'Dockerfile'
                    ) {
                        foundPaths.push(file);
                    }
                }
                for (const folder of list.folders) {
                    const foldername = folder.split('/').pop() || '';
                    // 忽略大型或内部隐藏目录
                    if (foldername !== '.git' && foldername !== '.obsidian' && foldername !== 'node_modules' && foldername !== '.trash') {
                        await scanDir(folder);
                    }
                }
            } catch {
                // 忽略权限或访问异常
            }
        };

        await scanDir('');
        this.items = foundPaths.map((p) => ({ path: p }));
    }

    getItems(): FileItem[] {
        const query = this.inputEl.value.trim();
        if (query && !this.items.some((i) => i.path === query || i.path.endsWith('/' + query))) {
            return [{ path: query, isNew: true }, ...this.items];
        }
        return this.items;
    }

    getItemText(item: FileItem): string {
        return item.isNew ? `[创建并编辑新文件] ${item.path}` : item.path;
    }

    async onChooseItem(item: FileItem, evt: MouseEvent | KeyboardEvent): Promise<void> {
        let file = this.app.vault.getAbstractFileByPath(item.path);
        if (!file && item.isNew) {
            try {
                file = await this.app.vault.create(item.path, '');
                new Notice(`已创建文件: ${item.path}`);
            } catch (err) {
                new Notice(`创建文件失败: ${err instanceof Error ? err.message : String(err)}`);
                return;
            }
        }

        if (file instanceof TFile) {
            const leaf = this.app.workspace.getLeaf(false);
            await leaf.setViewState({
                type: VIEW_TYPE_CODE,
                state: { file: file.path }
            });
            this.app.workspace.setActiveLeaf(leaf, { focus: true });
        } else {
            new Notice(`无法以视图打开: ${item.path}`);
        }
    }
}
