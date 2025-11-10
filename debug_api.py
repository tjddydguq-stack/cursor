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

def get_header(customer_id, api_key, secret_key):
    timestamp = str(int(time.time() * 1000))
    signature = make_signature(timestamp, api_key, secret_key)
    return {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Timestamp': timestamp,
        'X-API-KEY': api_key,
        'X-Customer': customer_id,
        'X-Signature': signature
    }

def make_signature(timestamp, api_key, secret_key):
    message = timestamp + '.' + api_key
    sign = hmac.new(secret_key.encode('utf-8'), message.encode('utf-8'), digestmod=hashlib.sha256).digest()
    return base64.b64encode(sign).decode('utf-8')

def debug_keyword_search(keyword):
    """키워드 검색량 조회 (디버깅 버전)"""
    url = f"{API_URL}/keywordstool"
    headers = get_header(CUSTOMER_ID, API_KEY, SECRET_KEY)
    
    print(f"\n=== '{keyword}' 키워드 디버깅 ===")
    print(f"API URL: {url}")
    print(f"Headers: {headers}")
    
    # 여러 가지 파라미터 조합 시도
    param_combinations = [
        {
            "hintKeywords": [keyword],
            "showDetail": "1"
        },
        {
            "hintKeywords": keyword,  # 리스트가 아닌 문자열
            "showDetail": "1"
        },
        {
            "hintKeywords": [keyword],
            "showDetail": 1  # 숫자
        },
        {
            "hintKeywords": keyword,
            "showDetail": 1
        }
    ]
    
    for i, params in enumerate(param_combinations):
        print(f"\n--- 파라미터 조합 {i+1}: {params} ---")
        
        try:
            response = requests.get(url, headers=headers, params=params)
            print(f"응답 상태코드: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                print(f"응답 데이터: {json.dumps(data, ensure_ascii=False, indent=2)}")
                
                keyword_list = data.get('keywordList', [])
                print(f"키워드 리스트 개수: {len(keyword_list)}")
                
                if keyword_list:
                    print("첫 번째 키워드 데이터:")
                    first_item = keyword_list[0]
                    for key, value in first_item.items():
                        print(f"  {key}: {value}")
                    
                    # 검색량 필드들 확인
                    search_fields = ['monthPcQcCnt', 'monthMobileQcCnt', 'monthlyPcQcCnt', 'monthlyMobileQcCnt', 'pcQcCnt', 'mobileQcCnt']
                    for field in search_fields:
                        if field in first_item:
                            print(f"  검색량 필드 '{field}': {first_item[field]}")
                
                # 정확한 키워드 찾기
                for item in keyword_list:
                    if item.get('relKeyword') == keyword:
                        print(f"정확한 키워드 '{keyword}' 발견!")
                        return item.get('monthPcQcCnt', 0), item.get('monthMobileQcCnt', 0)
                
                # 정확한 키워드가 없으면 첫 번째 결과 사용
                if keyword_list:
                    first_item = keyword_list[0]
                    pc_count = first_item.get('monthPcQcCnt', first_item.get('monthlyPcQcCnt', 0))
                    mobile_count = first_item.get('monthMobileQcCnt', first_item.get('monthlyMobileQcCnt', 0))
                    print(f"첫 번째 결과 사용 - PC: {pc_count}, 모바일: {mobile_count}")
                    return pc_count, mobile_count
                
            else:
                print(f"API 오류: {response.status_code}")
                print(f"응답 내용: {response.text}")
                
        except Exception as e:
            print(f"오류 발생: {e}")
    
    return 0, 0

if __name__ == "__main__":
    keyword = "음식물처리기"
    pc_count, mobile_count = debug_keyword_search(keyword)
    
    print(f"\n=== 최종 결과 ===")
    print(f"키워드: {keyword}")
    print(f"PC 검색량: {pc_count:,}회")
    print(f"모바일 검색량: {mobile_count:,}회")
    print(f"총 검색량: {pc_count + mobile_count:,}회")
    
    input("\n엔터를 누르면 종료...")
