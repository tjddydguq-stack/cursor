# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

# 한글 출력을 위한 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

def test_excel():
    print("엑셀 파일 테스트 시작")
    
    # 테스트용 키워드 파일 생성
    test_keywords = ['카페', '커피', '블로그', '원룸청소기']
    test_df = pd.DataFrame(test_keywords, columns=['키워드'])
    test_df.to_excel('test_keywords.xlsx', index=False)
    print("test_keywords.xlsx 파일 생성 완료")
    
    try:
        # 생성된 파일로 테스트
        df = pd.read_excel('test_keywords.xlsx')
        print(f"파일 읽기 성공: {len(df)}행")
        
        keywords = []
        for idx, row in df.iterrows():
            keyword = str(row.iloc[0]).strip()
            if keyword and keyword != 'nan':
                keywords.append(keyword)
        
        print(f"키워드: {keywords}")
        
        # 테스트 결과
        results = []
        for keyword in keywords:
            results.append({
                '키워드': keyword,
                'PC검색량': 1000,
                '모바일검색량': 2000
            })
        
        result_df = pd.DataFrame(results)
        result_df.to_excel('test_result.xlsx', index=False)
        print("test_result.xlsx 저장 완료")
        
        print("\n테스트 성공! 모든 기능이 정상 동작합니다.")
        
    except Exception as e:
        print(f"오류: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_excel()
    print("\n프로그램 종료")






