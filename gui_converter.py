import os
import win32com.client
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT → PDF 최종 해결 버전")
        self.root.geometry("600x600")
        
        self.input_folder = None
        self.output_folder = None

        # 1. 입력 폴더 구역
        frame_in = tk.LabelFrame(root, text="1. PPT가 들어있는 폴더를 선택하세요", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        frame_in.pack(fill="x", padx=15, pady=10)

        self.btn_in = tk.Button(frame_in, text="📂 입력 폴더 찾기", font=("맑은 고딕", 10), 
                                bg="#E3F2FD", width=20, command=self.select_input)
        self.btn_in.pack(side="left")

        self.lbl_in = tk.Label(frame_in, text="선택안됨", fg="gray", font=("맑은 고딕", 9))
        self.lbl_in.pack(side="left", padx=10)

        # 2. 출력 폴더 구역
        frame_out = tk.LabelFrame(root, text="2. PDF 저장할 곳 (선택 안 하면 입력 폴더에 저장)", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        frame_out.pack(fill="x", padx=15, pady=5)

        self.btn_out = tk.Button(frame_out, text="💾 저장 폴더 찾기", font=("맑은 고딕", 10), 
                                 bg="#FFF3E0", width=20, command=self.select_output)
        self.btn_out.pack(side="left")

        self.lbl_out = tk.Label(frame_out, text="자동 생성 (_converted_pdf)", fg="gray", font=("맑은 고딕", 9))
        self.lbl_out.pack(side="left", padx=10)

        # 3. 실행 버튼
        self.btn_start = tk.Button(root, text="🚀 변환 시작하기", font=("맑은 고딕", 12, "bold"), 
                                   bg="#4CAF50", fg="white", width=30, height=2,
                                   state='disabled', command=self.start_thread)
        self.btn_start.pack(pady=20)

        # 4. 로그 창
        self.log_area = scrolledtext.ScrolledText(root, width=70, height=15, state='disabled')
        self.log_area.pack(pady=5)
        self.log("시스템 준비 완료. '입력 폴더'를 선택해주세요.")

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def select_input(self):
        path = filedialog.askdirectory(title="PPT 폴더 선택")
        if path:
            # 경로를 절대 경로 문자열로 변환 (중요)
            self.input_folder = os.path.abspath(path)
            self.lbl_in.config(text=self.input_folder, fg="black")
            self.check_ready()
            self.log(f"✅ 입력 폴더: {self.input_folder}")

    def select_output(self):
        path = filedialog.askdirectory(title="저장 폴더 선택")
        if path:
            self.output_folder = os.path.abspath(path)
            self.lbl_out.config(text=self.output_folder, fg="black")
            self.log(f"✅ 저장 폴더: {self.output_folder}")

    def check_ready(self):
        if self.input_folder:
            self.btn_start.config(state='normal')

    def start_thread(self):
        t = threading.Thread(target=self.convert_process)
        t.start()

    def convert_process(self):
        if not self.input_folder:
            return

        # 출력 폴더 설정
        if self.output_folder:
            final_out_dir = self.output_folder
        else:
            final_out_dir = os.path.join(self.input_folder, "_converted_pdf")
            if not os.path.exists(final_out_dir):
                os.makedirs(final_out_dir)
        
        self.btn_start.config(state='disabled', text="탐색 중...", bg="#9E9E9E")
        self.log(f"🔍 폴더를 정밀 탐색합니다... (방식: os.walk)")

        # ==========================================
        # [강력한 탐색] os.walk 사용
        # ==========================================
        ppt_files = []
        target_exts = ['.ppt', '.pptx', '.pptm', '.potx', '.ppsx']

        for root, dirs, files in os.walk(self.input_folder):
            for file in files:
                # 확장자 체크 (소문자로 변환해서 비교)
                ext = os.path.splitext(file)[1].lower()
                if ext in target_exts:
                    full_path = os.path.join(root, file)
                    
                    # 결과 폴더에 있는 파일은 건너뛰기
                    if str(final_out_dir) in full_path:
                        continue
                        
                    ppt_files.append(full_path)

        if not ppt_files:
            messagebox.showwarning("파일 없음", 
                                   f"정말 이상하네요. 선택한 폴더:\n{self.input_folder}\n\n"
                                   "이 안에 PPT 파일이 없는 것으로 나옵니다.\n"
                                   "혹시 폴더 접근 권한 문제일 수도 있으니 바탕화면으로 옮겨서 시도해보세요.")
            self.btn_start.config(state='normal', text="🚀 변환 시작하기", bg="#4CAF50")
            return

        self.log(f"🚀 총 {len(ppt_files)}개의 파일을 찾았습니다!")

        # 파워포인트 실행
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        except:
            messagebox.showerror("에러", "파워포인트 실행 실패! 설치되어 있는지 확인해주세요.")
            self.btn_start.config(state='normal', text="🚀 변환 시작하기", bg="#4CAF50")
            return

        success_count = 0
        for i, ppt_path_str in enumerate(ppt_files, 1):
            
            # 파일 이름 만들기
            ppt_path = Path(ppt_path_str)
            folder_name = ppt_path.parent.name
            file_stem = ppt_path.stem
            
            # 입력 폴더 바로 아래면 파일명만, 깊은 폴더면 '폴더명_파일명'
            if os.path.dirname(ppt_path_str) == self.input_folder:
                pdf_name = f"{file_stem}.pdf"
            else:
                pdf_name = f"{folder_name}_{file_stem}.pdf"
            
            save_path = os.path.join(final_out_dir, pdf_name)
            
            try:
                self.log(f"[{i}/{len(ppt_files)}] 변환 중: {ppt_path.name}")
                deck = powerpoint.Presentations.Open(ppt_path_str, WithWindow=False)
                deck.SaveAs(save_path, 32)
                deck.Close()
                success_count += 1
            except Exception as e:
                self.log(f"❌ 실패: {ppt_path.name}\n   이유: {e}")

        powerpoint.Quit()
        self.log("="*40)
        self.log(f"🎉 완료! 성공: {success_count}개")
        messagebox.showinfo("완료", f"작업 끝!\n저장 폴더: {final_out_dir}")
        
        self.btn_start.config(state='normal', text="🚀 변환 시작하기", bg="#4CAF50")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()