@echo off
echo 라이브러리 확인 중...
python check_libraries.py
echo.
echo 엑셀 테스트 실행 중...
python test_excel_simple.py
echo.
echo 모든 테스트 완료
pause






