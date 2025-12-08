import sys

print("Python 버전:", sys.version)
print("=" * 50)

# 필요한 라이브러리들 확인
libraries = ['pandas', 'requests', 'openpyxl']

for lib in libraries:
    try:
        __import__(lib)
        print(f"✓ {lib} 설치됨")
    except ImportError:
        print(f"✗ {lib} 미설치")

print("=" * 50)
print("라이브러리 설치 명령어:")
print("pip install pandas requests openpyxl")

input("\n엔터를 누르면 창이 닫힙니다...")







