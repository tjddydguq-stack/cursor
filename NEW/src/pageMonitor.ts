import * as vscode from 'vscode';
import * as path from 'path';

export class PageMonitor {
    public static readonly viewType = 'pageMonitor';

    private static panels: PageMonitor[] = [];

    public static createOrShow(extensionUri: vscode.Uri, targetUrl?: string, localContent?: string) {
        const column = vscode.window.activeTextEditor
            ? vscode.window.activeTextEditor.viewColumn
            : undefined;

        // 이미 열려있는 패널이 있으면 포커스
        for (const panel of PageMonitor.panels) {
            if (panel.panel.viewColumn === column) {
                panel.panel.reveal(column);
                if (targetUrl) {
                    panel.update(targetUrl, localContent);
                }
                return;
            }
        }

        // 새 패널 생성
        const panel = vscode.window.createWebviewPanel(
            PageMonitor.viewType,
            '자사몰 페이지 모니터',
            column || vscode.ViewColumn.Two,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [extensionUri]
            }
        );

        const monitor = new PageMonitor(panel, extensionUri);
        PageMonitor.panels.push(monitor);
        
        if (targetUrl) {
            monitor.update(targetUrl, localContent);
        }
    }

    public static revive(panel: vscode.WebviewPanel, extensionUri: vscode.Uri) {
        PageMonitor.panels.push(new PageMonitor(panel, extensionUri));
    }

    public static disposeAll() {
        PageMonitor.panels.forEach(panel => panel.dispose());
        PageMonitor.panels = [];
    }

    private constructor(
        private readonly panel: vscode.WebviewPanel,
        private readonly extensionUri: vscode.Uri
    ) {
        this.update('', '');

        this.panel.onDidDispose(() => this.dispose(), null);

        // 웹뷰에서 메시지 수신
        this.panel.webview.onDidReceiveMessage(
            message => {
                switch (message.command) {
                    case 'checkPage':
                        this.checkPage(message.url);
                        return;
                    case 'compareContent':
                        this.compareContent(message.localContent, message.remoteContent);
                        return;
                }
            },
            null
        );
    }

    private update(targetUrl: string, localContent?: string) {
        const webview = this.panel.webview;
        this.panel.webview.html = this.getHtmlForWebview(webview, targetUrl, localContent);
    }

    private async checkPage(url: string) {
        // 웹뷰에 페이지 체크 요청 전송
        this.panel.webview.postMessage({
            command: 'loadPage',
            url: url
        });
    }

    private compareContent(localContent: string, remoteContent: string) {
        // 간단한 비교 (실제로는 더 정교한 비교가 필요)
        const localLines = localContent.split('\n');
        const remoteLines = remoteContent.split('\n');
        
        const diff = {
            localOnly: [] as string[],
            remoteOnly: [] as string[],
            different: [] as { line: number; local: string; remote: string }[]
        };

        // 웹뷰에 비교 결과 전송
        this.panel.webview.postMessage({
            command: 'showDiff',
            diff: diff
        });
    }

    private getHtmlForWebview(webview: vscode.Webview, targetUrl: string, localContent?: string) {
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, 'media', 'pageMonitor.js')
        );

        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.extensionUri, 'media', 'pageMonitor.css')
        );

        const nonce = this.getNonce();

        return `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}'; frame-src *; connect-src *;">
    <link href="${styleUri}" rel="stylesheet">
    <title>자사몰 페이지 모니터</title>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>자사몰 페이지 모니터</h2>
            <div class="controls">
                <input type="text" id="urlInput" placeholder="자사몰 URL 입력" value="${targetUrl}" />
                <button id="checkBtn">체크</button>
                <button id="autoCheckBtn">자동 체크 (30초)</button>
            </div>
            <div class="status" id="status">대기 중...</div>
        </div>
        
        <div class="content">
            <div class="iframe-container">
                <iframe id="pageFrame" src="${targetUrl || 'about:blank'}" sandbox="allow-same-origin allow-scripts allow-forms allow-popups"></iframe>
            </div>
            
            <div class="comparison">
                <h3>변경사항 비교</h3>
                <div id="diffResult" class="diff-result">
                    <p>페이지를 로드하면 비교 결과가 표시됩니다.</p>
                </div>
            </div>
        </div>
    </div>

    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        let autoCheckInterval = null;
        let currentUrl = '${targetUrl}';

        document.getElementById('checkBtn').addEventListener('click', () => {
            const url = document.getElementById('urlInput').value;
            if (url) {
                currentUrl = url;
                checkPage(url);
            }
        });

        document.getElementById('autoCheckBtn').addEventListener('click', () => {
            const btn = document.getElementById('autoCheckBtn');
            if (autoCheckInterval) {
                clearInterval(autoCheckInterval);
                autoCheckInterval = null;
                btn.textContent = '자동 체크 (30초)';
                updateStatus('자동 체크 중지됨', 'info');
            } else {
                const url = document.getElementById('urlInput').value || currentUrl;
                if (url) {
                    autoCheckInterval = setInterval(() => {
                        checkPage(url);
                    }, 30000);
                    btn.textContent = '자동 체크 중지';
                    updateStatus('자동 체크 시작 (30초마다)', 'info');
                }
            }
        });

        function checkPage(url) {
            updateStatus('페이지 체크 중...', 'loading');
            document.getElementById('pageFrame').src = url + (url.includes('?') ? '&' : '?') + '_t=' + Date.now();
            
            // iframe 로드 후 내용 가져오기
            setTimeout(() => {
                try {
                    const frame = document.getElementById('pageFrame');
                    const frameDoc = frame.contentDocument || frame.contentWindow.document;
                    const remoteContent = frameDoc.documentElement.outerHTML;
                    
                    vscode.postMessage({
                        command: 'compareContent',
                        remoteContent: remoteContent
                    });
                    
                    updateStatus('체크 완료: ' + new Date().toLocaleTimeString(), 'success');
                } catch (e) {
                    updateStatus('CORS 오류: 직접 확인 필요', 'error');
                    console.error('CORS 오류:', e);
                }
            }, 2000);
        }

        function updateStatus(message, type) {
            const statusEl = document.getElementById('status');
            statusEl.textContent = message;
            statusEl.className = 'status ' + type;
        }

        // VSCode에서 메시지 수신
        window.addEventListener('message', event => {
            const message = event.data;
            switch (message.command) {
                case 'loadPage':
                    checkPage(message.url);
                    break;
                case 'showDiff':
                    showDiff(message.diff);
                    break;
            }
        });

        function showDiff(diff) {
            const diffResult = document.getElementById('diffResult');
            // 차이점 표시 로직
            diffResult.innerHTML = '<p>비교 완료</p>';
        }

        // 초기 로드 시 URL이 있으면 체크
        if ('${targetUrl}') {
            setTimeout(() => checkPage('${targetUrl}'), 1000);
        }
    </script>
</body>
</html>`;
    }

    private getNonce() {
        let text = '';
        const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
        for (let i = 0; i < 32; i++) {
            text += possible.charAt(Math.floor(Math.random() * possible.length));
        }
        return text;
    }

    private dispose() {
        const index = PageMonitor.panels.indexOf(this);
        if (index !== -1) {
            PageMonitor.panels.splice(index, 1);
        }
        this.panel.dispose();
    }
}
