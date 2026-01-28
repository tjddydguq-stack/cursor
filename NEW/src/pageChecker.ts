import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

export class PageChecker {
    private static instance: PageChecker | undefined;
    private watcher: vscode.FileSystemWatcher | undefined;
    private checkInterval: NodeJS.Timeout | undefined;
    private targetUrl: string = '';
    private targetFile: string = '';
    private lastContent: string = '';
    private statusBarItem: vscode.StatusBarItem;

    private constructor() {
        this.statusBarItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Right,
            100
        );
        this.statusBarItem.command = 'pageChecker.showStatus';
        this.statusBarItem.show();
    }

    public static getInstance(): PageChecker {
        if (!PageChecker.instance) {
            PageChecker.instance = new PageChecker();
        }
        return PageChecker.instance;
    }

    public async startWatching(filePath: string, url: string) {
        this.targetFile = filePath;
        this.targetUrl = url;
        
        // 파일 읽기
        try {
            this.lastContent = fs.readFileSync(filePath, 'utf-8');
            this.updateStatus('감시 중...', '$(sync~spin)');
        } catch (error) {
            vscode.window.showErrorMessage(`파일을 읽을 수 없습니다: ${filePath}`);
            return;
        }

        // 파일 감시 시작
        const pattern = new vscode.RelativePattern(
            vscode.workspace.workspaceFolders![0],
            path.basename(filePath)
        );
        
        this.watcher = vscode.workspace.createFileSystemWatcher(pattern);
        
        this.watcher.onDidChange(async (uri) => {
            await this.onFileChanged(uri);
        });

        // 주기적으로 체크 (30초마다)
        this.checkInterval = setInterval(() => {
            this.checkPage();
        }, 30000);

        // 즉시 한 번 체크
        this.checkPage();

        vscode.window.showInformationMessage(
            `자사몰 페이지 감시 시작: ${url}`
        );
    }

    private async onFileChanged(uri: vscode.Uri) {
        try {
            const newContent = fs.readFileSync(uri.fsPath, 'utf-8');
            
            if (newContent !== this.lastContent) {
                this.lastContent = newContent;
                this.updateStatus('파일 변경됨', '$(file-code)');
                
                // 파일이 변경되면 즉시 체크
                setTimeout(() => {
                    this.checkPage();
                }, 2000); // 2초 후 체크 (서버 반영 시간 고려)
            }
        } catch (error) {
            console.error('파일 변경 감지 오류:', error);
        }
    }

    private async checkPage() {
        if (!this.targetUrl) return;

        try {
            this.updateStatus('체크 중...', '$(sync~spin)');
            
            // 웹뷰로 메시지 전송하여 페이지 체크 요청
            vscode.commands.executeCommand('pageChecker.checkPage', {
                url: this.targetUrl,
                localContent: this.lastContent
            });

        } catch (error) {
            this.updateStatus('체크 실패', '$(error)');
            console.error('페이지 체크 오류:', error);
        }
    }

    private updateStatus(text: string, icon: string = '') {
        this.statusBarItem.text = `${icon} ${text}`;
        this.statusBarItem.tooltip = `자사몰 페이지 체크: ${this.targetUrl || '설정되지 않음'}`;
    }

    public stop() {
        if (this.watcher) {
            this.watcher.dispose();
            this.watcher = undefined;
        }
        if (this.checkInterval) {
            clearInterval(this.checkInterval);
            this.checkInterval = undefined;
        }
        this.statusBarItem.hide();
        this.updateStatus('중지됨', '$(stop)');
    }

    public dispose() {
        this.stop();
        this.statusBarItem.dispose();
        PageChecker.instance = undefined;
    }
}
