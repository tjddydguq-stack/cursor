# Claude Chat Extension

Cursor에서 Claude와 채팅할 수 있는 VS Code 확장 프로그램입니다.

## 개발 환경 설정

### 1. 의존성 설치

터미널에서 다음 명령어를 실행하여 필요한 패키지를 설치합니다:

```bash
npm install
```

### 2. Cursor에서 개발하는 방법

#### 방법 1: 디버그 모드로 실행

1. **F5 키를 누르거나** 좌측 사이드바의 "실행 및 디버그" 아이콘을 클릭합니다
2. "Extension" 구성이 선택되어 있는지 확인합니다
3. 새로운 Cursor 창(확장 프로그램 호스트)이 열립니다
4. 새 창에서 `Ctrl+Shift+P` (또는 `Cmd+Shift+P`)를 눌러 명령 팔레트를 엽니다
5. "Claude Chat 열기" 명령을 실행합니다

#### 방법 2: 수동 컴파일 후 테스트

터미널에서 다음 명령어로 TypeScript를 컴파일합니다:

```bash
npm run compile
```

또는 자동 감시 모드로 개발:

```bash
npm run watch
```

### 3. 프로젝트 구조

```
.
├── src/
│   ├── extension.ts      # 확장 프로그램 진입점
│   └── chatPanel.ts      # 채팅 패널 웹뷰 관리
├── out/                  # 컴파일된 JavaScript 파일 (자동 생성)
├── package.json          # 확장 프로그램 매니페스트
├── tsconfig.json         # TypeScript 설정
└── .vscode/
    ├── launch.json       # 디버그 구성
    └── tasks.json        # 빌드 작업
```

### 4. 주요 기능

- **웹뷰 기반 채팅 인터페이스**: VS Code 내에서 채팅 UI 제공
- **명령어 등록**: "Claude Chat 열기" 명령어로 채팅 패널 열기
- **상태 복원**: Cursor가 재시작되어도 열려있던 웹뷰 복원
- **자사몰 페이지 모니터**: HTML 파일 변경 시 자사몰 페이지 자동 체크
  - 파일 변경 감지 및 자동 체크
  - 웹뷰에서 자사몰 페이지 실시간 확인
  - 30초마다 자동 체크 기능

### 5. 자사몰 페이지 체크 사용법

1. **HTML 파일 열기**: 수정 중인 HTML 파일을 Cursor에서 엽니다
2. **명령어 실행**: `Ctrl+Shift+P` (또는 `Cmd+Shift+P`)를 눌러 명령 팔레트를 엽니다
3. **페이지 체크 시작**: "자사몰 페이지 체크 시작" 명령어를 선택합니다
4. **URL 입력**: 자사몰 상세페이지 URL을 입력합니다 (예: `https://yourstore.com/product/detail.html?product_no=123`)
5. **자동 감시 시작**: 
   - HTML 파일이 변경되면 자동으로 2초 후 체크합니다
   - 30초마다 자동으로 페이지를 체크합니다
   - 상태바에서 현재 상태를 확인할 수 있습니다

**수동 체크**: "자사몰 페이지 모니터 열기" 명령어로 웹뷰 패널을 열어 수동으로 체크할 수도 있습니다.

### 6. 개발 팁

- `Ctrl+Shift+P` → "Developer: Reload Window"로 확장 프로그램 재로드
- 디버그 콘솔에서 `console.log` 출력 확인 가능
- `out/` 디렉토리는 자동 생성되므로 `.gitignore`에 포함됨
- 페이지 체크 중지: "자사몰 페이지 체크 중지" 명령어로 감시를 중지할 수 있습니다

### 7. 확장 프로그램 패키징

배포용 패키지를 만들려면:

```bash
npm install -g @vscode/vsce
vsce package
```

이렇게 하면 `.vsix` 파일이 생성됩니다.

## 다음 단계

- [ ] Claude API 연동 구현
- [ ] 채팅 기록 저장 기능
- [ ] 설정 페이지 추가
- [ ] 더 많은 커스터마이징 옵션
