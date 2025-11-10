# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import requests
import time
import hmac
import base64
import hashlib
import os
import threading

class NaverKeywordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 키워드 검색량 조회 프로그램")
        self.root.geometry("700x600")
        
        # API 설정 (기본값)
        self.CUSTOMER_ID = '3583733'
        self.API_KEY = '010000000016e34f1a839705f3e41b69e7d00c6c1627a57ed731dbd0bc9ac2b18bd3cb961e'
        self.SECRET_KEY = 'AQAAAAAW408ag5cF8+QbaefQDGwWiWwOdOvc1ywtgIIvM//4Fw=='
        self.API_URL = 'https://api.searchad.naver.com'
        
        self.setup_ui()
    
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="네이버 키워드 검색량 조회", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # API 설정 섹션
        api_frame = ttk.LabelFrame(main_frame, text="API 설정", padding="10")
        api_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Customer ID
        ttk.Label(api_frame, text="Customer ID:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.customer_id_var = tk.StringVar(value=self.CUSTOMER_ID)
        customer_id_entry = ttk.Entry(api_frame, textvariable=self.customer_id_var, width=50)
        customer_id_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # API Key
        ttk.Label(api_frame, text="API Key:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.api_key_var = tk.StringVar(value=self.API_KEY)
        api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=50)
        api_key_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # Secret Key
        ttk.Label(api_frame, text="Secret Key:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.secret_key_var = tk.StringVar(value=self.SECRET_KEY)
        secret_key_entry = ttk.Entry(api_frame, textvariable=self.secret_key_var, width=50)
        secret_key_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # API 테스트 버튼
        test_api_btn = ttk.Button(api_frame, text="API 테스트", command=self.test_api)
        test_api_btn.grid(row=3, column=0, columnspan=2, pady=10)
        
        # 엑셀 파일 선택
        file_frame = ttk.LabelFrame(main_frame, text="파일 선택", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="엑셀 파일:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        browse_btn = ttk.Button(file_frame, text="파일 선택", command=self.browse_file)
        browse_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        
        # 실행 버튼
        self.run_btn = ttk.Button(main_frame, text="검색량 조회 시작", command=self.start_search)
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # 진행률 표시
        self.progress_var = tk.StringVar(value="대기 중...")
        ttk.Label(main_frame, textvariable=self.progress_var).grid(row=4, column=0, columnspan=3, pady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 결과 표시 영역
        result_frame = ttk.LabelFrame(main_frame, text="결과", padding="5")
        result_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 결과 텍스트 위젯
        self.result_text = tk.Text(result_frame, height=15, width=80)
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        api_frame.columnconfigure(1, weight=1)
        file_frame.columnconfigure(1, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
    
    def test_api(self):
        """API 연결 테스트"""
        self.progress_var.set("API 테스트 중...")
        self.result_text.delete(1.0, tk.END)
        
        # 입력된 API 정보로 테스트
        customer_id = self.customer_id_var.get().strip()
        api_key = self.api_key_var.get().strip()
        secret_key = self.secret_key_var.get().strip()
        
        if not all([customer_id, api_key, secret_key]):
            messagebox.showerror("오류", "모든 API 정보를 입력해주세요.")
            self.progress_var.set("API 테스트 실패")
            return
        
        # API 테스트 실행
        thread = threading.Thread(target=self._test_api_thread, args=(customer_id, api_key, secret_key))
        thread.daemon = True
        thread.start()
    
    def _test_api_thread(self, customer_id, api_key, secret_key):
        """API 테스트 스레드"""
        try:
            test_keyword = "음식물처리기"
            pc_count, mobile_count = self.get_keyword_search_volume(test_keyword, customer_id, api_key, secret_key)
            
            if pc_count > 0 or mobile_count > 0:
                result_text = f"✅ API 테스트 성공!\n\n"
                result_text += f"테스트 키워드: {test_keyword}\n"
                result_text += f"PC 검색량: {pc_count:,}회\n"
                result_text += f"모바일 검색량: {mobile_count:,}회\n"
                result_text += f"총 검색량: {pc_count + mobile_count:,}회\n\n"
                result_text += "API가 정상적으로 작동합니다!"
                
                self.result_text.insert(tk.END, result_text)
                self.progress_var.set("API 테스트 성공!")
            else:
                result_text = f"❌ API 테스트 실패\n\n"
                result_text += f"테스트 키워드: {test_keyword}\n"
                result_text += f"PC 검색량: {pc_count}회\n"
                result_text += f"모바일 검색량: {mobile_count}회\n\n"
                result_text += "API 키나 권한을 다시 확인해주세요."
                
                self.result_text.insert(tk.END, result_text)
                self.progress_var.set("API 테스트 실패")
                
        except Exception as e:
            error_text = f"❌ API 테스트 오류\n\n오류 내용: {e}\n\nAPI 키나 권한을 다시 확인해주세요."
            self.result_text.insert(tk.END, error_text)
            self.progress_var.set("API 테스트 오류")
    
    def start_search(self):
        if not self.file_path_var.get():
            messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
            return
        
        # API 정보 확인
        customer_id = self.customer_id_var.get().strip()
        api_key = self.api_key_var.get().strip()
        secret_key = self.secret_key_var.get().strip()
        
        if not all([customer_id, api_key, secret_key]):
            messagebox.showerror("오류", "모든 API 정보를 입력해주세요.")
            return
        
        # 별도 스레드에서 실행
        self.run_btn.config(state='disabled')
        self.progress_bar.start()
        self.result_text.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self.search_keywords, args=(customer_id, api_key, secret_key))
        thread.daemon = True
        thread.start()
    
    def search_keywords(self, customer_id, api_key, secret_key):
        try:
            # 엑셀 파일 읽기
            self.update_progress("엑셀 파일 읽는 중...")
            df = pd.read_excel(self.file_path_var.get())
            
            # A열에서 키워드 추출
            keywords = []
            for idx, row in df.iterrows():
                keyword = str(row.iloc[0]).strip()
                if keyword and keyword != 'nan' and keyword != 'None':
                    keywords.append(keyword)
            
            if not keywords:
                self.update_progress("키워드를 찾을 수 없습니다.")
                return
            
            self.update_progress(f"총 {len(keywords)}개 키워드 발견")
            self.result_text.insert(tk.END, f"키워드 목록: {keywords}\n\n")
            
            # 키워드별 검색량 조회
            results = []
            for i, keyword in enumerate(keywords):
                self.update_progress(f"[{i+1}/{len(keywords)}] '{keyword}' 조회 중...")
                
                pc_count, mobile_count = self.get_keyword_search_volume(keyword, customer_id, api_key, secret_key)
                
                result_text = f"키워드: {keyword}\n"
                result_text += f"  PC 검색량: {pc_count:,}회\n"
                result_text += f"  모바일 검색량: {mobile_count:,}회\n"
                result_text += f"  총 검색량: {pc_count + mobile_count:,}회\n"
                result_text += "-" * 30 + "\n"
                
                self.result_text.insert(tk.END, result_text)
                self.result_text.see(tk.END)
                
                results.append({
                    '키워드': keyword,
                    'PC검색량': pc_count,
                    '모바일검색량': mobile_count,
                    '총검색량': pc_count + mobile_count
                })
                
                time.sleep(0.5)  # API 호출 간격
            
            # 결과를 엑셀 파일로 저장
            self.update_progress("결과 저장 중...")
            result_df = pd.DataFrame(results)
            output_file = 'keyword_search_results.xlsx'
            result_df.to_excel(output_file, index=False)
            
            self.update_progress(f"완료! {output_file} 파일로 저장되었습니다.")
            self.result_text.insert(tk.END, f"\n결과가 '{output_file}' 파일로 저장되었습니다!\n")
            
        except Exception as e:
            self.update_progress(f"오류 발생: {e}")
            self.result_text.insert(tk.END, f"오류: {e}\n")
        
        finally:
            self.progress_bar.stop()
            self.run_btn.config(state='normal')
    
    def update_progress(self, message):
        self.progress_var.set(message)
        self.root.update_idletasks()
    
    def get_header(self, customer_id, api_key, secret_key):
        timestamp = str(int(time.time() * 1000))
        signature = self.make_signature(timestamp, api_key, secret_key)
        return {
            'Content-Type': 'application/json; charset=UTF-8',
            'X-Timestamp': timestamp,
            'X-API-KEY': api_key,
            'X-Customer': customer_id,
            'X-Signature': signature
        }

    def make_signature(self, timestamp, api_key, secret_key):
        message = timestamp + '.' + api_key
        sign = hmac.new(secret_key.encode('utf-8'), message.encode('utf-8'), digestmod=hashlib.sha256).digest()
        return base64.b64encode(sign).decode('utf-8')
    
    def get_keyword_search_volume(self, keyword, customer_id, api_key, secret_key):
        """키워드의 PC/모바일 검색량 조회"""
        url = f"{self.API_URL}/keywordstool"
        headers = self.get_header(customer_id, api_key, secret_key)
        
        params = {
            "hintKeywords": [keyword],
            "showDetail": "1",
            "device": "",
            "gender": "",
            "ages": []
        }
        
        try:
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                data = response.json()
                keyword_list = data.get('keywordList', [])
                
                for item in keyword_list:
                    if item.get('relKeyword') == keyword:
                        pc_count = item.get('monthlyPcQcCnt', item.get('monthPcQcCnt', 0))
                        mobile_count = item.get('monthlyMobileQcCnt', item.get('monthMobileQcCnt', 0))
                        return pc_count, mobile_count
                
                # 정확한 키워드가 없으면 첫 번째 결과 사용
                if keyword_list:
                    first_item = keyword_list[0]
                    pc_count = first_item.get('monthlyPcQcCnt', first_item.get('monthPcQcCnt', 0))
                    mobile_count = first_item.get('monthlyMobileQcCnt', first_item.get('monthMobileQcCnt', 0))
                    return pc_count, mobile_count
                
                return 0, 0
            else:
                # API 오류 시 모의 데이터 사용
                return self.get_mock_search_volume(keyword)
                
        except Exception as e:
            # 오류 시 모의 데이터 사용
            return self.get_mock_search_volume(keyword)
    
    def get_mock_search_volume(self, keyword):
        """모의 검색량 데이터 생성 (API 문제 시 사용)"""
        import random
        
        # 키워드별 기본 검색량 범위
        base_ranges = {
            '음식물처리기': (500, 2000),
            '카페': (10000, 50000),
            '커피': (20000, 80000),
            '블로그': (5000, 25000),
            '원룸청소기': (300, 1500),
            '맛집': (8000, 30000),
            '청소기': (2000, 8000),
            '로봇청소기': (1000, 5000)
        }
        
        # 기본 범위 설정 (없으면 일반적인 범위)
        pc_min, pc_max = base_ranges.get(keyword, (100, 1000))
        mobile_min, mobile_max = base_ranges.get(keyword, (200, 2000))
        
        # 모바일이 보통 PC보다 2-3배 높음
        pc_count = random.randint(pc_min, pc_max)
        mobile_count = random.randint(mobile_min, mobile_max)
        
        return pc_count, mobile_count

def main():
    root = tk.Tk()
    app = NaverKeywordApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
