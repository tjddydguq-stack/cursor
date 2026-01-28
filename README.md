# OAuth 토큰 관리 시스템

OAuth 토큰 만료 오류를 자동으로 처리하는 Python 유틸리티입니다.

## 기능

- ✅ OAuth 토큰 자동 관리
- ✅ 토큰 만료 감지 및 자동 갱신
- ✅ 401 오류 발생 시 자동 재로그인
- ✅ 토큰 파일 기반 저장/로드

## 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 1. 기본 사용 (AuthManager)

```python
from auth_manager import AuthManager

# AuthManager 초기화
auth = AuthManager(login_endpoint="/login")

# 유효한 토큰 가져오기 (자동 갱신 포함)
token = auth.get_valid_token()

# 인증이 필요한 API 호출
response = auth.make_authenticated_request(
    'GET',
    'https://api.example.com/data'
)
```

### 2. API 클라이언트 사용

```python
from api_client import APIClient

# API 클라이언트 초기화
client = APIClient(
    base_url="https://api.example.com",
    login_endpoint="/login"
)

# API 호출 (자동 토큰 관리)
response = client.get("/data")
if response.status_code == 200:
    print(response.json())
```

### 3. 로그인 실행

```bash
python login.py
```

또는

```bash
python -m login
```

## 설정

### 실제 API 연동하기

1. **auth_manager.py** 수정:
   - `login()` 메서드에 실제 로그인 API 호출 로직 추가
   - `refresh_access_token()` 메서드에 리프레시 토큰 API 호출 로직 추가

2. **api_client.py** 수정:
   - `base_url`을 실제 API URL로 변경

3. **login.py** 수정:
   - 실제 로그인 프로세스 구현

## 파일 구조

- `auth_manager.py`: OAuth 토큰 관리 클래스
- `api_client.py`: API 클라이언트 (자동 토큰 관리 포함)
- `login.py`: 로그인 스크립트
- `.token.json`: 저장된 토큰 (자동 생성)

## 오류 처리

시스템은 다음 상황을 자동으로 처리합니다:

1. **토큰 만료**: 자동으로 리프레시 토큰으로 갱신 시도
2. **401 오류**: 토큰 갱신 또는 재로그인 후 자동 재시도
3. **리프레시 실패**: 자동으로 재로그인 시도

## 주의사항

- `.token.json` 파일은 민감한 정보를 포함하므로 `.gitignore`에 추가하세요
- 실제 API 엔드포인트와 인증 방식을 코드에 맞게 수정해야 합니다


