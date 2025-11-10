# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

# 한글 출력을 위한 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

def test_excel():
    print("엑셀 파일 테스트 시작")
    
    try:
        excel_file = input("엑셀 파일명: ").strip()
        
        if not os.path.exists(excel_file):
            print(f"파일 없음: {excel_file}")
            return
        
        df = pd.read_excel(excel_file)
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
        
    except Exception as e:
        print(f"오류: {e}")
        import traceback
        traceback.print_exc()
    
    try:
        input("엔터를 누르면 종료...")
    except EOFError:
        print("입력 대기 중...")

if __name__ == "__main__":
    test_excel()
