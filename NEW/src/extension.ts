import * as vscode from 'vscode';
import { ChatPanel } from './chatPanel';
import { PageChecker } from './pageChecker';
import { PageMonitor } from './pageMonitor';
import * as path from 'path';
import * as fs from 'fs';

export function activate(context: vscode.ExtensionContext) {
    console.log('Claude Chat Extension이 활성화되었습니다.');

    // 채팅 패널 열기 명령어
    const openChatCommand = vscode.commands.registerCommand('claudeChat.openChat', () => {
        ChatPanel.createOrShow(context.extensionUri);
    });

    // 채팅 패널 닫기 명령어
    const closeChatCommand = vscode.commands.registerCommand('claudeChat.closeChat', () => {
        ChatPanel.disposeAll();
    });

    context.subscriptions.push(openChatCommand);
    context.subscriptions.push(closeChatCommand);

    // 페이지 모니터 열기 명령어
    const openMonitorCommand = vscode.commands.registerCommand('pageMonitor.open', async () => {
        const url = await vscode.window.showInputBox({
            prompt: '자사몰 상세페이지 URL을 입력하세요',
            placeHolder: 'https://yourstore.com/product/detail.html?product_no=123'
        });
        
        if (url) {
            const activeEditor = vscode.window.activeTextEditor;
            let localContent = '';
            
            if (activeEditor && activeEditor.document.fileName.endsWith('.html')) {
                localContent = activeEditor.document.getText();
            }
            
            PageMonitor.createOrShow(context.extensionUri, url, localContent);
        }
    });

    // 페이지 체크 시작 명령어
    const startCheckCommand = vscode.commands.registerCommand('pageChecker.start', async () => {
        const activeEditor = vscode.window.activeTextEditor;
        if (!activeEditor || !activeEditor.document.fileName.endsWith('.html')) {
            vscode.window.showWarningMessage('HTML 파일을 열어주세요.');
            return;
        }

        const url = await vscode.window.showInputBox({
            prompt: '자사몰 상세페이지 URL을 입력하세요',
            placeHolder: 'https://yourstore.com/product/detail.html?product_no=123'
        });

        if (url) {
            const filePath = activeEditor.document.uri.fsPath;
            const checker = PageChecker.getInstance();
            await checker.startWatching(filePath, url);
            
            // 모니터 패널도 열기
            const localContent = activeEditor.document.getText();
            PageMonitor.createOrShow(context.extensionUri, url, localContent);
        }
    });

    // 페이지 체크 중지 명령어
    const stopCheckCommand = vscode.commands.registerCommand('pageChecker.stop', () => {
        const checker = PageChecker.getInstance();
        checker.stop();
        vscode.window.showInformationMessage('페이지 체크가 중지되었습니다.');
    });

    // 페이지 체크 명령어 (수동)
    const checkPageCommand = vscode.commands.registerCommand('pageChecker.checkPage', async (data?: { url: string, localContent: string }) => {
        if (data) {
            PageMonitor.createOrShow(context.extensionUri, data.url, data.localContent);
        } else {
            const url = await vscode.window.showInputBox({
                prompt: '체크할 URL을 입력하세요'
            });
            if (url) {
                PageMonitor.createOrShow(context.extensionUri, url);
            }
        }
    });

    context.subscriptions.push(
        openMonitorCommand,
        startCheckCommand,
        stopCheckCommand,
        checkPageCommand
    );

    // 웹뷰가 이미 열려있으면 복원
    if (vscode.window.registerWebviewPanelSerializer) {
        const serializer = vscode.window.registerWebviewPanelSerializer(PageMonitor.viewType, {
            async deserializeWebviewPanel(webviewPanel: vscode.WebviewPanel, state: any) {
                PageMonitor.revive(webviewPanel, context.extensionUri);
            }
        });
        context.subscriptions.push(serializer);
    }
}

export function deactivate() {
    // 모든 ChatPanel 인스턴스 정리
    ChatPanel.disposeAll();
    PageMonitor.disposeAll();
    PageChecker.getInstance().dispose();
}
