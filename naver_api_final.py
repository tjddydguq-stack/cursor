# -*- coding: utf-8 -*-
import requests
import pandas as pd
import time
import hmac
import base64
import hashlib
import json
from datetime import datetime, timedelta

# 네이버 검색광고 API 인증정보
CUSTOMER_ID = '3583733'
API_KEY = '010000000016e34f1a839705f3e41b69e7d00c6c1627a57ed731dbd0bc9ac2b18bd3cb961e'
SECRET_KEY = 'AQAAAAAW408ag5cF8+QbaefQDGwWiWwOdOvc1ywtgIIvM//4Fw=='
API_URL = 'https://api.searchad.naver.com'

def create_signature(timestamp, api_key, secret_key):
    """네이버 검색광고 API 서명 생성"""
    message = timestamp + '.' + api_key
    signature = hmac.new(
        secret_key.encode('utf-8'),
        message.encode('utf-8'),
        digestmod=hashlib.sha256
    ).digest()
    return base64.b64encode(signature).decode('utf-8')

def get_headers():
    """API 헤더 생성"""
    timestamp = str(int(time.time() * 1000))
    signature = create_signature(timestamp, API_KEY, SECRET_KEY)
    
    return {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Timestamp': timestamp,
        'X-API-KEY': API_KEY,
        'X-Customer': CUSTOMER_ID,
        'X-Signature': signature
    }

def get_keyword_search_volume(keyword):
    """키워드 검색량 조회"""
    url = f"{API_URL}/keywordstool"
    headers = get_headers()
    
    # 네이버 검색광고 API 파라미터
    params = {
        "hintKeywords": [keyword],
        "showDetail": "1"
    }
    
    try:
        print(f"키워드 '{keyword}' 조회 중...")
        print(f"URL: {url}")
        print(f"Headers: {headers}")
        print(f"Params: {params}")
        
        response = requests.get(url, headers=headers, params=params)
        print(f"응답 상태: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"응답 데이터: {json.dumps(data, ensure_ascii=False, indent=2)}")
            
            keyword_list = data.get('keywordList', [])
            if keyword_list:
                # 정확한 키워드 찾기
                for item in keyword_list:
                    if item.get('relKeyword') == keyword:
                        pc_count = item.get('monthPcQcCnt', 0)
                        mobile_count = item.get('monthMobileQcCnt', 0)
                        return pc_count, mobile_count
                
                # 정확한 키워드가 없으면 첫 번째 결과 사용
                first_item = keyword_list[0]
                pc_count = first_item.get('monthPcQcCnt', 0)
                mobile_count = first_item.get('monthMobileQcCnt', 0)
                return pc_count, mobile_count
            
            return 0, 0
        else:
            print(f"API 오류: {response.status_code}")
            print(f"응답 내용: {response.text}")
            return 0, 0
            
    except Exception as e:
        print(f"오류 발생: {e}")
        return 0, 0

def test_single_keyword():
    """단일 키워드 테스트"""
    print("=== 네이버 검색광고 API 테스트 ===")
    
    keyword = "음식물처리기"
    pc_count, mobile_count = get_keyword_search_volume(keyword)
    
    print(f"\n=== 결과 ===")
    print(f"키워드: {keyword}")
    print(f"PC 검색량: {pc_count:,}회")
    print(f"모바일 검색량: {mobile_count:,}회")
    print(f"총 검색량: {pc_count + mobile_count:,}회")
    
    return pc_count, mobile_count

def process_excel_file(file_path):
    """엑셀 파일 처리"""
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path)
        print(f"엑셀 파일 읽기 성공: {len(df)}행")
        
        # A열에서 키워드 추출
        keywords = []
        for idx, row in df.iterrows():
            keyword = str(row.iloc[0]).strip()
            if keyword and keyword != 'nan' and keyword != 'None':
                keywords.append(keyword)
        
        if not keywords:
            print("키워드를 찾을 수 없습니다.")
            return
        
        print(f"추출된 키워드: {keywords}")
        
        # 키워드별 검색량 조회
        results = []
        for i, keyword in enumerate(keywords, 1):
            print(f"\n[{i}/{len(keywords)}] '{keyword}' 처리 중...")
            pc_count, mobile_count = get_keyword_search_volume(keyword)
            
            result = {
                '키워드': keyword,
                'PC검색량': pc_count,
                '모바일검색량': mobile_count,
                '총검색량': pc_count + mobile_count
            }
            results.append(result)
            
            print(f"  PC: {pc_count:,}회, 모바일: {mobile_count:,}회")
            time.sleep(1)  # API 호출 간격
        
        # 결과 저장
        result_df = pd.DataFrame(results)
        output_file = 'keyword_search_results.xlsx'
        result_df.to_excel(output_file, index=False)
        
        print(f"\n✅ 완료! '{output_file}' 파일로 저장되었습니다.")
        
        return results
        
    except Exception as e:
        print(f"엑셀 파일 처리 오류: {e}")
        return None

if __name__ == "__main__":
    print("네이버 키워드 검색량 조회 프로그램")
    print("=" * 50)
    
    # 먼저 단일 키워드로 API 테스트
    pc, mobile = test_single_keyword()
    
    if pc > 0 or mobile > 0:
        print("\n✅ API가 정상 동작합니다!")
        
        # 엑셀 파일 처리
        excel_file = input("\n엑셀 파일명을 입력하세요 (예: keywords.xlsx): ").strip()
        if excel_file:
            process_excel_file(excel_file)
    else:
        print("\n❌ API 호출에 문제가 있습니다.")
        print("API 키나 권한을 다시 확인해주세요.")
    
    print("\n프로그램 종료")







