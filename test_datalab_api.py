# -*- coding: utf-8 -*-
import requests
import pandas as pd
import json
import time

# 네이버 데이터랩 API (검색어 트렌드)
CLIENT_ID = 'doNogXAl6_CjwOMFYgg3'
CLIENT_SECRET = 'mmGfXSl4zB'

def get_keyword_trend_data(keywords, start_date, end_date):
    """네이버 데이터랩 API로 키워드 트렌드 조회"""
    url = 'https://openapi.naver.com/v1/datalab/search'
    headers = {
        'X-Naver-Client-Id': CLIENT_ID,
        'X-Naver-Client-Secret': CLIENT_SECRET,
        'Content-Type': 'application/json'
    }
    
    request_body = {
        "startDate": start_date,
        "endDate": end_date,
        "timeUnit": "date",
        "keywordGroups": [
            {
                "groupName": "키워드",
                "keywords": keywords
            }
        ],
        "device": "",
        "ages": [],
        "gender": ""
    }
    
    try:
        response = requests.post(url, headers=headers, data=json.dumps(request_body))
        print(f"데이터랩 API 응답 상태: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"데이터랩 API 응답: {json.dumps(data, ensure_ascii=False, indent=2)}")
            return data
        else:
            print(f"데이터랩 API 오류: {response.status_code}")
            print(f"응답 내용: {response.text}")
            return None
            
    except Exception as e:
        print(f"데이터랩 API 오류: {e}")
        return None

def test_data_lab_api():
    """데이터랩 API 테스트"""
    print("=== 네이버 데이터랩 API 테스트 ===")
    
    # 최근 30일 데이터 조회
    from datetime import datetime, timedelta
    end_date = datetime.now()
    start_date = end_date - timedelta(days=30)
    
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')
    
    print(f"조회 기간: {start_str} ~ {end_str}")
    
    keywords = ["음식물처리기", "카페", "커피"]
    
    result = get_keyword_trend_data(keywords, start_str, end_str)
    
    if result:
        print("데이터랩 API 성공!")
        return True
    else:
        print("데이터랩 API 실패!")
        return False

if __name__ == "__main__":
    success = test_data_lab_api()
    
    if success:
        print("\n✅ 데이터랩 API가 정상 동작합니다!")
        print("검색광고 API 대신 데이터랩 API를 사용하는 것을 권장합니다.")
    else:
        print("\n❌ 데이터랩 API도 실패했습니다.")
        print("API 키를 다시 확인해주세요.")
    
    print("\n프로그램 종료")







