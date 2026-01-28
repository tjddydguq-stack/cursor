# CLASP 사용 가이드

## CLASP란?

**CLASP (Command Line Apps Script Projects)**는 Google Apps Script 프로젝트를 로컬에서 개발하고 관리할 수 있게 해주는 명령줄 도구입니다.

### 왜 CLASP를 사용하나요?

1. **로컬 개발**: 코드를 로컬에서 작성하고 Git으로 버전 관리 가능
2. **자동 배포**: 명령어 하나로 Google Apps Script에 코드 업로드
3. **편리한 관리**: 여러 프로젝트를 쉽게 관리
4. **IDE 사용**: VS Code 등에서 코드 작성 가능

---

## 설치 방법

### 1단계: Node.js 설치 확인

```bash
node --version
npm --version
```

Node.js가 없다면 [nodejs.org](https://nodejs.org)에서 설치하세요.

### 2단계: CLASP 설치

```bash
npm install -g @google/clasp
```

### 3단계: CLASP 로그인

```bash
clasp login
```

브라우저가 열리면 Google 계정으로 로그인하세요.

---

## 프로젝트 설정

### 1단계: Apps Script 프로젝트 생성

1. [Google Apps Script](https://script.google.com) 접속
2. 새 프로젝트 생성
3. 프로젝트 설정에서 **프로젝트 ID** 복사

### 2단계: 로컬 프로젝트 연결

프로젝트 폴더에서 실행:

```bash
cd google-sheets-keyword
clasp clone <프로젝트_ID>
```

또는 기존 코드가 있다면:

```bash
clasp create --title "키워드 자동 분류 시스템" --type sheets
```

### 3단계: .clasp.json 파일 생성

프로젝트 루트에 `.clasp.json` 파일이 생성됩니다:

```json
{
  "scriptId": "여기에_프로젝트_ID",
  "rootDir": "."
}
```

---

## 주요 명령어

### 코드 업로드 (Push)

로컬 코드를 Google Apps Script에 업로드:

```bash
clasp push
```

### 코드 다운로드 (Pull)

Google Apps Script에서 최신 코드 다운로드:

```bash
clasp pull
```

### 프로젝트 열기

브라우저에서 Apps Script 프로젝트 열기:

```bash
clasp open
```

### 배포

스크립트를 웹 앱으로 배포:

```bash
clasp deploy
```

### 로그 확인

실행 로그 확인:

```bash
clasp logs
```

---

## 현재 프로젝트 설정 방법

### 방법 1: 새 프로젝트 생성

```bash
cd google-sheets-keyword

# 1. 새 Apps Script 프로젝트 생성
clasp create --title "키워드 자동 분류 시스템" --type sheets

# 2. 코드 업로드
clasp push

# 3. Google Sheets에 연결
# clasp open 명령어로 열린 Apps Script에서
# "리소스 > 스프레드시트에 연결" 선택
```

### 방법 2: 기존 프로젝트 연결

```bash
cd google-sheets-keyword

# 1. Apps Script 프로젝트 ID 확인
# (Apps Script 편집기 URL에서 확인: https://script.google.com/home/projects/프로젝트_ID/edit)

# 2. .clasp.json 파일 생성
echo '{"scriptId": "여기에_프로젝트_ID", "rootDir": "."}' > .clasp.json

# 3. 코드 다운로드 (기존 코드 백업)
clasp pull

# 4. 로컬 코드 업로드
clasp push
```

---

## 파일 구조

```
google-sheets-keyword/
├── .clasp.json          # CLASP 설정 파일 (자동 생성)
├── appsscript.json      # Apps Script 설정 (이미 있음)
├── Code.js              # 메인 코드
├── Code_Public.js       # 공개용 코드
└── ...
```

---

## 주의사항

### 1. .clasp.json은 Git에 포함해도 됨

프로젝트 ID는 공개되어도 문제없습니다.

### 2. 민감한 정보는 환경변수 사용

API 키 등은 Apps Script의 속성 서비스 사용:

```javascript
// Code.js에서
const API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');
```

설정 방법:
```bash
clasp open
# Apps Script 편집기에서:
# 파일 > 프로젝트 설정 > 스크립트 속성
```

### 3. 여러 파일 관리

여러 `.gs` 파일이 있다면:
- `Code.js` → `Code.gs`
- `Utils.js` → `Utils.gs`

CLASP는 `.js` 파일을 자동으로 `.gs`로 변환합니다.

---

## 자주 사용하는 워크플로우

### 개발 → 배포

```bash
# 1. 로컬에서 코드 수정
# (Code.js 파일 편집)

# 2. Google Apps Script에 업로드
clasp push

# 3. Google Sheets에서 테스트
# (시트 새로고침 후 메뉴 확인)
```

### Google에서 수정 → 로컬 동기화

```bash
# 1. Google Apps Script에서 코드 수정

# 2. 로컬로 다운로드
clasp pull

# 3. Git에 커밋
git add .
git commit -m "Update from Apps Script"
```

---

## 문제 해결

### "clasp: command not found"

```bash
# 전역 설치 확인
npm list -g @google/clasp

# 재설치
npm install -g @google/clasp
```

### "Please run clasp login"

```bash
clasp login
```

### "Script ID required"

`.clasp.json` 파일에 `scriptId`가 있는지 확인하세요.

---

## 참고 자료

- [CLASP 공식 문서](https://github.com/google/clasp)
- [Apps Script 가이드](https://developers.google.com/apps-script)
- [CLASP 명령어 목록](https://github.com/google/clasp#commands)
