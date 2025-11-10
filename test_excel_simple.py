import pandas as pd
import os

print("=" * 50)
print("엑셀 파일 테스트 (API 없음)")
print("=" * 50)

try:
    # 엑셀 파일 입력받기
    excel_file = input("엑셀 파일명을 입력하세요: ").strip()
    
    if not excel_file:
        print("파일명을 입력해주세요.")
        input("엔터를 누르면 창이 닫힙니다...")
        exit()
    
    if not os.path.exists(excel_file):
        print(f"파일이 존재하지 않습니다: {excel_file}")
        input("엔터를 누르면 창이 닫힙니다...")
        exit()
    
    # 엑셀 파일 읽기
    print(f"파일 읽는 중: {excel_file}")
    df = pd.read_excel(excel_file)
    print(f"파일 읽기 성공! 총 {len(df)}개 행")
    
    # A열 키워드 추출
    keywords = []
    for idx, row in df.iterrows():
        keyword = str(row.iloc[0]).strip()
        if keyword and keyword != 'nan' and keyword != 'None':
            keywords.append(keyword)
    
    print(f"추출된 키워드: {keywords}")
    
    # 테스트 결과 생성
    results = []
    for keyword in keywords:
        results.append({
            '키워드': keyword,
            'PC검색량': 1000,  # 테스트용 더미 데이터
            '모바일검색량': 2000,  # 테스트용 더미 데이터
            '총검색량': 3000
        })
    
    # 결과 저장
    result_df = pd.DataFrame(results)
    output_file = 'test_results.xlsx'
    result_df.to_excel(output_file, index=False)
    
    print(f"\n테스트 완료! {output_file} 파일로 저장됨")
    
except Exception as e:
    print(f"오류 발생: {e}")
    print(f"오류 타입: {type(e).__name__}")
    
    # 오류를 파일로 저장
    with open('error_log.txt', 'w', encoding='utf-8') as f:
        f.write(f"오류: {e}\n타입: {type(e).__name__}")

input("\n엔터를 누르면 창이 닫힙니다...")






