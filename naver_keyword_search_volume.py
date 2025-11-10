# -*- coding: utf-8 -*-
import requests
import pandas as pd
import time
import hmac
import base64
import hashlib
import os
import sys

# 한글 출력을 위한 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

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

def get_keyword_search_volume(keyword):
    """키워드의 PC/모바일 검색량 조회"""
    url = f"{API_URL}/keywordstool"
    headers = get_header(CUSTOMER_ID, API_KEY, SECRET_KEY)
    
    # 올바른 파라미터 설정
    params = {
        "hintKeywords": [keyword],  # 리스트 형태로 전달
        "showDetail": "1",
        "device": "",
        "gender": "",
        "ages": []
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        print(f"API 응답 상태: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"API 응답 데이터: {data}")  # 디버깅용
            
            keyword_list = data.get('keywordList', [])
            
            for item in keyword_list:
                if item.get('relKeyword') == keyword:
                    # 올바른 필드명 사용
                    pc_count = item.get('monthlyPcQcCnt', 0)  # monthPcQcCnt -> monthlyPcQcCnt
                    mobile_count = item.get('monthlyMobileQcCnt', 0)  # monthMobileQcCnt -> monthlyMobileQcCnt
                    
                    # 다른 가능한 필드명들도 시도
                    if pc_count == 0:
                        pc_count = item.get('monthPcQcCnt', 0)
                    if mobile_count == 0:
                        mobile_count = item.get('monthMobileQcCnt', 0)
                    
                    return pc_count, mobile_count
            
            # 정확한 키워드가 없으면 첫 번째 결과 사용
            if keyword_list:
                first_item = keyword_list[0]
                pc_count = first_item.get('monthlyPcQcCnt', first_item.get('monthPcQcCnt', 0))
                mobile_count = first_item.get('monthlyMobileQcCnt', first_item.get('monthMobileQcCnt', 0))
                return pc_count, mobile_count
            
            return 0, 0
        else:
            print(f"API 오류 (상태코드: {response.status_code})")
            print(f"응답 내용: {response.text}")
            return 0, 0
            
    except Exception as e:
        print(f"키워드 '{keyword}' 조회 중 오류: {e}")
        return 0, 0

def main():
    print("=" * 50)
    print("네이버 키워드 검색량 조회 프로그램")
    print("=" * 50)
    
    # 엑셀 파일 입력받기
    while True:
        excel_file = input("\n엑셀 파일명을 입력하세요 (예: keywords.xlsx): ").strip()
        
        if not excel_file:
            print("파일명을 입력해주세요.")
            continue
            
        if not os.path.exists(excel_file):
            print(f"파일이 존재하지 않습니다: {excel_file}")
            continue
            
        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(excel_file)
            print(f"\n파일 읽기 성공! 총 {len(df)}개 행 발견")
            
            # A열(첫 번째 열)에서 키워드 추출
            keywords = []
            for idx, row in df.iterrows():
                keyword = str(row.iloc[0]).strip()
                if keyword and keyword != 'nan' and keyword != 'None':
                    keywords.append(keyword)
            
            if not keywords:
                print("A열에서 유효한 키워드를 찾을 수 없습니다.")
                continue
                
            print(f"추출된 키워드: {keywords}")
            break
            
        except Exception as e:
            print(f"엑셀 파일 읽기 오류: {e}")
            continue
    
    # 키워드별 검색량 조회
    print(f"\n{'='*50}")
    print("키워드별 검색량 조회 중...")
    print(f"{'='*50}")
    
    results = []
    
    for i, keyword in enumerate(keywords, 1):
        print(f"[{i}/{len(keywords)}] '{keyword}' 조회 중...")
        
        pc_count, mobile_count = get_keyword_search_volume(keyword)
        
        print(f"  → PC: {pc_count:,}회, 모바일: {mobile_count:,}회")
        
        results.append({
            '키워드': keyword,
            'PC검색량': pc_count,
            '모바일검색량': mobile_count,
            '총검색량': pc_count + mobile_count
        })
        
        # API 호출 간격 (너무 빠르면 차단될 수 있음)
        time.sleep(0.5)
    
    # 결과 출력
    print(f"\n{'='*50}")
    print("최종 결과")
    print(f"{'='*50}")
    
    for result in results:
        print(f"키워드: {result['키워드']}")
        print(f"  PC 검색량: {result['PC검색량']:,}회")
        print(f"  모바일 검색량: {result['모바일검색량']:,}회")
        print(f"  총 검색량: {result['총검색량']:,}회")
        print("-" * 30)
    
    # 결과를 엑셀 파일로 저장
    result_df = pd.DataFrame(results)
    output_file = 'keyword_search_results.xlsx'
    result_df.to_excel(output_file, index=False)
    
    print(f"\n결과가 '{output_file}' 파일로 저장되었습니다!")
    
    input("\n엔터를 누르면 프로그램이 종료됩니다...")

if __name__ == "__main__":
    main()