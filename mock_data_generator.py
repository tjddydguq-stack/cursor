# -*- coding: utf-8 -*-
import pandas as pd
import random
import time

def get_mock_search_volume(keyword):
    """모의 검색량 데이터 생성 (API 문제 해결 전까지 임시 사용)"""
    # 키워드별로 다른 범위의 검색량 생성
    base_ranges = {
        '음식물처리기': (500, 2000),
        '카페': (10000, 50000),
        '커피': (20000, 80000),
        '블로그': (5000, 25000),
        '원룸청소기': (300, 1500),
        '맛집': (8000, 30000)
    }
    
    # 기본 범위 설정
    pc_min, pc_max = base_ranges.get(keyword, (100, 1000))
    mobile_min, mobile_max = base_ranges.get(keyword, (200, 2000))
    
    # 모바일이 보통 PC보다 2-3배 높음
    pc_count = random.randint(pc_min, pc_max)
    mobile_count = random.randint(mobile_min, mobile_max)
    
    return pc_count, mobile_count

def process_keywords_with_mock_data(keywords):
    """모의 데이터로 키워드 처리"""
    print("=== 모의 데이터로 키워드 검색량 조회 ===")
    print("⚠️  주의: 실제 API가 아닌 모의 데이터입니다.")
    print()
    
    results = []
    
    for i, keyword in enumerate(keywords, 1):
        print(f"[{i}/{len(keywords)}] '{keyword}' 처리 중...")
        
        pc_count, mobile_count = get_mock_search_volume(keyword)
        
        result_text = f"키워드: {keyword}\n"
        result_text += f"  PC 검색량: {pc_count:,}회\n"
        result_text += f"  모바일 검색량: {mobile_count:,}회\n"
        result_text += f"  총 검색량: {pc_count + mobile_count:,}회\n"
        result_text += "-" * 30 + "\n"
        
        print(result_text)
        
        results.append({
            '키워드': keyword,
            'PC검색량': pc_count,
            '모바일검색량': mobile_count,
            '총검색량': pc_count + mobile_count
        })
        
        time.sleep(0.5)  # 처리 시간 시뮬레이션
    
    return results

if __name__ == "__main__":
    # 테스트 키워드
    test_keywords = ['음식물처리기', '카페', '커피', '블로그']
    
    results = process_keywords_with_mock_data(test_keywords)
    
    # 결과를 엑셀 파일로 저장
    result_df = pd.DataFrame(results)
    output_file = 'mock_search_results.xlsx'
    result_df.to_excel(output_file, index=False)
    
    print(f"✅ 결과가 '{output_file}' 파일로 저장되었습니다!")
    print("\n📋 API 문제 해결 방법:")
    print("1. 네이버 개발자센터에서 API 키 재확인")
    print("2. API 사용 권한 활성화 확인")
    print("3. 애플리케이션 승인 상태 확인")
    
    print("\n프로그램 종료")







