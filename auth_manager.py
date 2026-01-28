"""
OAuth 토큰 관리 및 자동 갱신 유틸리티
"""
import os
import json
import time
from typing import Optional, Dict, Any
from datetime import datetime, timedelta
import requests


class AuthManager:
    """OAuth 토큰을 관리하고 자동으로 갱신하는 클래스"""
    
    def __init__(self, token_file: str = ".token.json", login_endpoint: str = "/login"):
        """
        Args:
            token_file: 토큰을 저장할 파일 경로
            login_endpoint: 로그인 엔드포인트 경로
        """
        self.token_file = token_file
        self.login_endpoint = login_endpoint
        self.token: Optional[str] = None
        self.refresh_token: Optional[str] = None
        self.expires_at: Optional[datetime] = None
        self.load_token()
    
    def load_token(self) -> bool:
        """저장된 토큰을 파일에서 로드"""
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.token = data.get('access_token')
                    self.refresh_token = data.get('refresh_token')
                    if data.get('expires_at'):
                        self.expires_at = datetime.fromisoformat(data['expires_at'])
                    return True
            except Exception as e:
                print(f"토큰 로드 실패: {e}")
        return False
    
    def save_token(self, access_token: str, refresh_token: Optional[str] = None, expires_in: int = 3600):
        """토큰을 파일에 저장"""
        self.token = access_token
        if refresh_token:
            self.refresh_token = refresh_token
        self.expires_at = datetime.now() + timedelta(seconds=expires_in)
        
        data = {
            'access_token': access_token,
            'expires_at': self.expires_at.isoformat()
        }
        if refresh_token:
            data['refresh_token'] = refresh_token
        
        try:
            with open(self.token_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"토큰 저장 실패: {e}")
    
    def is_token_expired(self, buffer_seconds: int = 60) -> bool:
        """토큰이 만료되었는지 확인 (버퍼 시간 포함)"""
        if not self.token or not self.expires_at:
            return True
        return datetime.now() >= (self.expires_at - timedelta(seconds=buffer_seconds))
    
    def refresh_access_token(self) -> bool:
        """리프레시 토큰을 사용하여 액세스 토큰 갱신"""
        if not self.refresh_token:
            return False
        
        try:
            # 여기에 실제 리프레시 토큰 API 호출 로직을 추가하세요
            # 예시:
            # response = requests.post(
            #     "https://api.example.com/oauth/refresh",
            #     data={"refresh_token": self.refresh_token}
            # )
            # if response.status_code == 200:
            #     data = response.json()
            #     self.save_token(data['access_token'], self.refresh_token, data.get('expires_in', 3600))
            #     return True
            print("리프레시 토큰으로 갱신 시도...")
            return False
        except Exception as e:
            print(f"토큰 갱신 실패: {e}")
            return False
    
    def login(self) -> bool:
        """로그인을 수행하고 새 토큰을 획득"""
        try:
            # 여기에 실제 로그인 API 호출 로직을 추가하세요
            # 예시:
            # response = requests.post(
            #     "https://api.example.com/login",
            #     json={"username": username, "password": password}
            # )
            # if response.status_code == 200:
            #     data = response.json()
            #     self.save_token(
            #         data['access_token'],
            #         data.get('refresh_token'),
            #         data.get('expires_in', 3600)
            #     )
            #     return True
            
            print(f"{self.login_endpoint} 엔드포인트로 로그인 시도...")
            print("실제 로그인 로직을 구현해주세요.")
            return False
        except Exception as e:
            print(f"로그인 실패: {e}")
            return False
    
    def get_valid_token(self) -> Optional[str]:
        """유효한 토큰을 반환 (필요시 자동 갱신 또는 재로그인)"""
        # 토큰이 없거나 만료된 경우
        if self.is_token_expired():
            print("토큰이 만료되었습니다. 갱신 시도...")
            
            # 리프레시 토큰으로 갱신 시도
            if self.refresh_token and self.refresh_access_token():
                return self.token
            
            # 리프레시 실패 시 재로그인
            print("리프레시 실패. 재로그인 시도...")
            if self.login():
                return self.token
            
            print("토큰 획득 실패. 수동으로 로그인해주세요.")
            return None
        
        return self.token
    
    def make_authenticated_request(self, method: str, url: str, **kwargs) -> requests.Response:
        """인증이 필요한 API 요청을 수행 (자동 토큰 갱신 포함)"""
        token = self.get_valid_token()
        if not token:
            raise Exception("유효한 토큰을 획득할 수 없습니다. 로그인이 필요합니다.")
        
        headers = kwargs.get('headers', {})
        headers['Authorization'] = f'Bearer {token}'
        kwargs['headers'] = headers
        
        response = requests.request(method, url, **kwargs)
        
        # 401 오류 발생 시 토큰 갱신 후 재시도
        if response.status_code == 401:
            print("401 오류 발생. 토큰 갱신 후 재시도...")
            if self.refresh_access_token() or self.login():
                token = self.get_valid_token()
                if token:
                    headers['Authorization'] = f'Bearer {token}'
                    kwargs['headers'] = headers
                    response = requests.request(method, url, **kwargs)
        
        return response


# 사용 예시
if __name__ == "__main__":
    # AuthManager 초기화
    auth = AuthManager(login_endpoint="/login")
    
    # 유효한 토큰 가져오기
    token = auth.get_valid_token()
    if token:
        print(f"토큰 획득 성공: {token[:20]}...")
    else:
        print("토큰 획득 실패. 로그인을 수행해주세요.")
    
    # 인증이 필요한 API 호출 예시
    # response = auth.make_authenticated_request(
    #     'GET',
    #     'https://api.example.com/data'
    # )
    # print(response.json())


