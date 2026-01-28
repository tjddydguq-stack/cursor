"""
로그인 스크립트 - OAuth 토큰 획득
"""
import os
import sys
from auth_manager import AuthManager


def login():
    """로그인을 수행하고 토큰을 저장"""
    auth = AuthManager(login_endpoint="/login")
    
    print("=" * 50)
    print("OAuth 로그인")
    print("=" * 50)
    
    # 여기에 실제 로그인 로직을 구현하세요
    # 예시:
    # username = input("사용자명: ")
    # password = input("비밀번호: ")
    # 
    # response = requests.post(
    #     "https://api.example.com/login",
    #     json={"username": username, "password": password}
    # )
    # 
    # if response.status_code == 200:
    #     data = response.json()
    #     auth.save_token(
    #         data['access_token'],
    #         data.get('refresh_token'),
    #         data.get('expires_in', 3600)
    #     )
    #     print("로그인 성공!")
    #     return True
    # else:
    #     print(f"로그인 실패: {response.status_code}")
    #     return False
    
    print("\n실제 로그인 API 엔드포인트와 인증 정보를 설정해주세요.")
    print("auth_manager.py와 api_client.py의 login() 메서드를 수정하세요.")
    
    return False


if __name__ == "__main__":
    success = login()
    sys.exit(0 if success else 1)


