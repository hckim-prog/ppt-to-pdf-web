import os
import win32com.client
from pathlib import Path

# ==========================================
# 설정: 여기에 PPT를 넣을 폴더 이름을 정해요
# ==========================================
INPUT_FOLDER = "ppt_files"      # PPT 넣어둘 폴더
OUTPUT_FOLDER = "pdf_results"   # PDF가 나올 폴더
# ==========================================

def convert_all_ppts():
    # 1. 내 위치 확인
    base_dir = Path.cwd()
    in_dir = base_dir / INPUT_FOLDER
    out_dir = base_dir / OUTPUT_FOLDER
    
    # 폴더가 없으면 알려주기
    if not in_dir.exists():
        in_dir.mkdir()
        print(f"⚠️ '{INPUT_FOLDER}' 폴더가 없어서 제가 만들었습니다!")
        print(f"👉 이 폴더 안에 변환할 PPT 파일들을 넣어주세요.")
        return
    
    if not out_dir.exists():
        out_dir.mkdir()

    # 2. 파워포인트 프로그램 몰래 실행
    print("⏳ 파워포인트를 실행하고 있습니다...")
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print("❌ 오류: 파워포인트가 설치되어 있지 않거나 실행할 수 없습니다.")
        return

    # 3. PPT 파일 찾기 (ppt, pptx, pptm 모두)
    ppt_files = list(in_dir.glob("*.ppt")) + list(in_dir.glob("*.pptx")) + list(in_dir.glob("*.pptm"))
    
    if not ppt_files:
        print(f"⚠️ '{INPUT_FOLDER}' 폴더가 비어있어요. PPT를 넣어주세요!")
        return

    print(f"🚀 총 {len(ppt_files)}개의 파일을 변환합니다! 잠시만 기다려주세요.")

    # 4. 하나씩 변환 시작
    success_count = 0
    for i, ppt_path in enumerate(ppt_files, 1):
        pdf_name = f"{ppt_path.stem}.pdf"
        pdf_path = out_dir / pdf_name
        
        try:
            print(f"[{i}/{len(ppt_files)}] 변환 중: {ppt_path.name}")
            
            # 파워포인트로 파일 열기 (창 안 띄우고)
            deck = powerpoint.Presentations.Open(str(ppt_path), WithWindow=False)
            
            # PDF로 다른 이름으로 저장 (32번 형식이 PDF)
            deck.SaveAs(str(pdf_path), 32)
            
            # 파일 닫기
            deck.Close()
            success_count += 1
            
        except Exception as e:
            print(f"❌ 실패: {ppt_path.name} -> {e}")

    powerpoint.Quit()
    print("="*30)
    print(f"🎉 모든 작업 끝! 성공: {success_count}개")
    print(f"📂 결과는 '{OUTPUT_FOLDER}' 폴더에 있습니다.")

if __name__ == "__main__":
    convert_all_ppts()