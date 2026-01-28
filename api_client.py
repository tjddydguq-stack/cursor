"""
API 클라이언트 - OAuth 토큰 자동 관리 포함
"""
import requests
from auth_manager import AuthManager
from typing import Optional, Dict, Any


class APIClient:
    """OAuth 인증을 자동으로 처리하는 API 클라이언트"""
    
    def __init__(self, base_url: str, login_endpoint: str = "/login"):
        """
        Args:
            base_url: API 기본 URL
            login_endpoint: 로그인 엔드포인트
        """
        self.base_url = base_url.rstrip('/')
        self.auth_manager = AuthManager(login_endpoint=login_endpoint)
    
    def _get_headers(self) -> Dict[str, str]:
        """인증 헤더를 포함한 요청 헤더 반환"""
        token = self.auth_manager.get_valid_token()
        if not token:
            raise Exception("유효한 토큰을 획득할 수 없습니다. /login을 실행해주세요.")
        
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
    
    def _handle_401_error(self, response: requests.Response) -> bool:
        """401 오류 처리 및 토큰 갱신"""
        if response.status_code == 401:
            error_data = response.json() if response.content else {}
            error_type = error_data.get('error', {}).get('type', '')
            
            if 'authentication_error' in error_type or 'expired' in str(error_data).lower():
                print("OAuth 토큰이 만료되었습니다. 토큰 갱신 시도...")
                
                # 리프레시 토큰으로 갱신 시도
                if self.auth_manager.refresh_access_token():
                    return True
                
                # 갱신 실패 시 재로그인
                print("토큰 갱신 실패. 재로그인 시도...")
                if self.auth_manager.login():
                    return True
                
                print("토큰 갱신 및 로그인 실패. 수동으로 /login을 실행해주세요.")
                return False
        
        return False
    
    def request(self, method: str, endpoint: str, **kwargs) -> requests.Response:
        """API 요청 수행 (자동 토큰 갱신 포함)"""
        url = f"{self.base_url}{endpoint}"
        headers = self._get_headers()
        
        # 기존 헤더와 병합
        if 'headers' in kwargs:
            headers.update(kwargs['headers'])
        kwargs['headers'] = headers
        
        # 첫 번째 요청 시도
        response = requests.request(method, url, **kwargs)
        
        # 401 오류 발생 시 토큰 갱신 후 재시도
        if response.status_code == 401:
            if self._handle_401_error(response):
                # 토큰 갱신 성공 시 재요청
                headers = self._get_headers()
                if 'headers' in kwargs:
                    headers.update(kwargs.get('headers', {}))
                kwargs['headers'] = headers
                response = requests.request(method, url, **kwargs)
        
        return response
    
    def get(self, endpoint: str, **kwargs) -> requests.Response:
        """GET 요청"""
        return self.request('GET', endpoint, **kwargs)
    
    def post(self, endpoint: str, data: Optional[Dict] = None, json: Optional[Dict] = None, **kwargs) -> requests.Response:
        """POST 요청"""
        if json:
            kwargs['json'] = json
        elif data:
            kwargs['data'] = data
        return self.request('POST', endpoint, **kwargs)
    
    def put(self, endpoint: str, data: Optional[Dict] = None, json: Optional[Dict] = None, **kwargs) -> requests.Response:
        """PUT 요청"""
        if json:
            kwargs['json'] = json
        elif data:
            kwargs['data'] = data
        return self.request('PUT', endpoint, **kwargs)
    
    def delete(self, endpoint: str, **kwargs) -> requests.Response:
        """DELETE 요청"""
        return self.request('DELETE', endpoint, **kwargs)


# 사용 예시
if __name__ == "__main__":
    # API 클라이언트 초기화
    # client = APIClient(base_url="https://api.example.com", login_endpoint="/login")
    
    # API 호출 예시 (자동으로 토큰 갱신 처리)
    # response = client.get("/data")
    # if response.status_code == 200:
    #     print(response.json())
    # else:
    #     print(f"오류: {response.status_code} - {response.text}")
    
    print("API 클라이언트가 준비되었습니다.")
    print("실제 API URL과 로그인 로직을 설정해주세요.")


