import re
import zipfile
import tempfile
import subprocess
import shutil
import uuid
from pathlib import Path

import streamlit as st


# -----------------------------
# 기본 화면
# -----------------------------
st.set_page_config(page_title="PPT → PDF 일괄 변환", layout="centered")
st.title("ZIP 폴더 안의 PPT만 골라서 PDF로 일괄 변환")

st.write("✅ PPT/PPTX/PPTM만 PDF로 변환합니다. (워드/엑셀/이미지 등은 자동으로 제외)")
st.write("✅ 변환 실패(0KB/손상) PDF는 결과 ZIP에 넣지 않고, 실패 목록만 보여줍니다.")
st.write("사용법: PPT가 들어있는 폴더를 ZIP으로 압축 → 업로드 → 변환 → PDFs.zip 다운로드")


# -----------------------------
# 유틸: soffice 경로 찾기(로컬/서버 모두)
# -----------------------------
def find_soffice():
    # 1) PATH에서 찾기
    p = shutil.which("soffice") or shutil.which("soffice.exe")
    if p:
        return p

    # 2) 리눅스에 흔한 경로
    for c in ["/usr/bin/soffice", "/usr/local/bin/soffice"]:
        if Path(c).exists():
            return str(Path(c))

    # 3) 윈도우 기본 경로(로컬 테스트용)
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return str(Path(c))

    return None


SOFFICE = find_soffice()


# -----------------------------
# 파일명 정렬: CH01 / CH02 ... 우선
# -----------------------------
def chapter_key(path: Path):
    name = path.name.upper()
    m = re.search(r"CH\s*0*(\d+)", name)
    if m:
        return (0, int(m.group(1)), path.name)

    m = re.search(r"0*(\d+)", name)
    if m:
        return (1, int(m.group(1)), path.name)

    return (2, 9999, path.name)


# -----------------------------
# 안전한 파일명 만들기
# -----------------------------
def safe_filename(name: str, max_len: int = 80) -> str:
    # OS에서 문제되는 문자 제거
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = name.strip()
    if len(name) > max_len:
        name = name[:max_len].rstrip()
    if not name:
        name = "file"
    return name


# -----------------------------
# PPT → PDF 변환 (0KB/손상 방지 포함)
# -----------------------------
def convert_ppt_to_pdf(input_path: Path, work_dir: Path, out_dir: Path, out_name: str) -> Path:
    if not SOFFICE:
        raise RuntimeError("LibreOffice(soffice)를 찾지 못했습니다. 서버에는 packages.txt로 설치되어야 합니다.")

    # 1) 입력을 '영어/짧은 임시 이름'으로 복사 (한글/긴 이름 이슈 방지)
    tmp_stem = f"input_{uuid.uuid4().hex}"
    safe_in = work_dir / f"{tmp_stem}{input_path.suffix.lower()}"
    shutil.copyfile(input_path, safe_in)

    # 2) LibreOffice로 변환 (필터 지정 + 타임아웃)
    cmd = [
        SOFFICE,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to", "pdf:impress_pdf_Export",
        "--outdir", str(out_dir),
        str(safe_in),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=240)

    if result.returncode != 0:
        raise RuntimeError(
            f"[변환 실패] {input_path.name}\n\nstderr:\n{result.stderr}\n\nstdout:\n{result.stdout}"
        )

    # 3) 변환된 PDF 찾기
    made_pdf = out_dir / f"{tmp_stem}.pdf"
    if not made_pdf.exists():
        raise RuntimeError(f"PDF 생성 실패(파일 없음): {input_path.name}")

    # 4) 0KB/손상 검사 (핵심)
    if made_pdf.stat().st_size < 5_000:
        raise RuntimeError(f"PDF가 0KB/비정상 크기(변환 실패): {input_path.name}")

    with open(made_pdf, "rb") as f:
        if f.read(5) != b"%PDF-":
            raise RuntimeError(f"PDF 헤더 없음(손상/비PDF): {input_path.name}")

    # 5) 최종 이름으로 변경
    final_pdf = out_dir / out_name
    if final_pdf.exists():
        final_pdf.unlink()
    made_pdf.rename(final_pdf)

    # 6) 임시 입력 정리
    if safe_in.exists():
        safe_in.unlink()

    return final_pdf


# -----------------------------
# 업로드
# -----------------------------
uploaded_zip = st.file_uploader("PPT가 들어있는 폴더를 ZIP으로 압축한 파일을 업로드", type=["zip"])

if uploaded_zip is None:
    st.info("ZIP 파일을 업로드하면 변환 버튼이 나타납니다.")
    st.stop()

if st.button("PDF로 일괄 변환 시작"):
    with st.spinner("변환 중… (파일 수/용량에 따라 시간이 걸릴 수 있어요)"):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            in_dir = tmp_dir / "inputs"
            extract_dir = tmp_dir / "unzipped"
            work_dir = tmp_dir / "work"
            pdf_dir = tmp_dir / "pdfs"

            in_dir.mkdir()
            extract_dir.mkdir()
            work_dir.mkdir()
            pdf_dir.mkdir()

            # 1) 업로드된 ZIP 저장 & 해제
            zip_path = in_dir / uploaded_zip.name
            zip_path.write_bytes(uploaded_zip.getvalue())

            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(extract_dir)

            # 2) PPT만 찾기 (다른 파일은 자동 제외)
            allowed_ext = {".ppt", ".pptx", ".pptm"}
            all_files = list(extract_dir.rglob("*"))
            ppt_files = [p for p in all_files if p.is_file() and p.suffix.lower() in allowed_ext]

            # (참고용) 제외된 파일 목록
            skipped_files = [p for p in all_files if p.is_file() and p.suffix.lower() not in allowed_ext]

            if not ppt_files:
                st.error("ZIP 안에서 PPT/PPTX/PPTM 파일을 찾지 못했습니다.")
                st.stop()

            # 3) 정렬
            ppt_files.sort(key=chapter_key)

            st.subheader("변환 대상(PPT만)")
            for p in ppt_files:
                st.write(f"- {p.relative_to(extract_dir)}")

            if skipped_files:
                st.subheader("자동 제외된 파일(변환 안 함)")
                # 너무 길면 일부만 보여주기
                for p in skipped_files[:30]:
                    st.write(f"- {p.relative_to(extract_dir)}")
                if len(skipped_files) > 30:
                    st.write(f"(…외 {len(skipped_files)-30}개 더 있음)")

            # 4) 변환
            progress = st.progress(0)
            errors = []
            success_count = 0

            for i, ppt in enumerate(ppt_files, start=1):
                try:
                    base = safe_filename(ppt.stem)
                    out_name = f"{i:02d}_{base}.pdf"
                    convert_ppt_to_pdf(ppt, work_dir=work_dir, out_dir=pdf_dir, out_name=out_name)
                    success_count += 1
                except Exception as e:
                    errors.append(str(e))

                progress.progress(int(i / len(ppt_files) * 100))

            # 5) 성공한 PDF만 ZIP으로 묶기
            out_zip = tmp_dir / "PDFs.zip"
            pdf_list = sorted(pdf_dir.glob("*.pdf"))

            with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for pdf in pdf_list:
                    z.write(pdf, arcname=pdf.name)

            st.success(f"완료! 성공: {success_count}개 / 전체: {len(ppt_files)}개")
            st.download_button(
                "PDFs.zip 다운로드(정상 PDF만 포함)",
                data=out_zip.read_bytes(),
                file_name="PDFs.zip",
                mime="application/zip"
            )

            # 6) 실패 목록 출력
            if errors:
                st.warning("아래 파일은 변환에 실패했습니다(0KB 방지로 ZIP에 넣지 않았습니다).")
                for e in errors:
                    st.code(e)
