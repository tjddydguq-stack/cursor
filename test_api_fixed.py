# -*- coding: utf-8 -*-
import requests
import pandas as pd
import time
import hmac
import base64
import hashlib
import json

# 네이버 광고 API 인증정보
CUSTOMER_ID = '3583733'
API_KEY = '010000000016e34f1a839705f3e41b69e7d00c6c1627a57ed731dbd0bc9ac2b18bd3cb961e'
SECRET_KEY = 'AQAAAAAW408ag5cF8+QbaefQDGwWiWwOdOvc1ywtgIIvM//4Fw=='
API_URL = 'https://api.searchad.naver.com'

def get_header_fixed(customer_id, api_key, secret_key):
    """수정된 헤더 생성 함수"""
    timestamp = str(int(time.time() * 1000))
    
    # 서명 생성 방식 수정
    message = timestamp + '.' + api_key
    signature = hmac.new(
        secret_key.encode('utf-8'), 
        message.encode('utf-8'), 
        digestmod=hashlib.sha256
    ).digest()
    signature_b64 = base64.b64encode(signature).decode('utf-8')
    
    print(f"타임스탬프: {timestamp}")
    print(f"메시지: {message}")
    print(f"서명: {signature_b64}")
    
    return {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Timestamp': timestamp,
        'X-API-KEY': api_key,
        'X-Customer': customer_id,
        'X-Signature': signature_b64
    }

def test_keyword_search(keyword):
    """키워드 검색량 조회 테스트"""
    url = f"{API_URL}/keywordstool"
    headers = get_header_fixed(CUSTOMER_ID, API_KEY, SECRET_KEY)
    
    print(f"\n=== '{keyword}' 키워드 테스트 ===")
    print(f"API URL: {url}")
    
    # 간단한 파라미터로 테스트
    params = {
        "hintKeywords": [keyword],
        "showDetail": "1"
    }
    
    print(f"파라미터: {params}")
    
    try:
        response = requests.get(url, headers=headers, params=params)
        print(f"응답 상태코드: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"응답 데이터: {json.dumps(data, ensure_ascii=False, indent=2)}")
            
            keyword_list = data.get('keywordList', [])
            if keyword_list:
                first_item = keyword_list[0]
                print(f"첫 번째 키워드 데이터: {first_item}")
                
                # 모든 필드 확인
                for key, value in first_item.items():
                    print(f"  {key}: {value}")
                
                return first_item.get('monthPcQcCnt', 0), first_item.get('monthMobileQcCnt', 0)
            else:
                print("키워드 리스트가 비어있습니다.")
                return 0, 0
        else:
            print(f"API 오류: {response.status_code}")
            print(f"응답 내용: {response.text}")
            return 0, 0
            
    except Exception as e:
        print(f"오류 발생: {e}")
        return 0, 0

if __name__ == "__main__":
    keyword = "음식물처리기"
    pc_count, mobile_count = test_keyword_search(keyword)
    
    print(f"\n=== 최종 결과 ===")
    print(f"키워드: {keyword}")
    print(f"PC 검색량: {pc_count:,}회")
    print(f"모바일 검색량: {mobile_count:,}회")
    print(f"총 검색량: {pc_count + mobile_count:,}회")
    
    print("\n프로그램 종료")






