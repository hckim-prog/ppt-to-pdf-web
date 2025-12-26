import re
import zipfile
import tempfile
import subprocess
from pathlib import Path

import streamlit as st

# -----------------------------
# 화면 기본 설정
# -----------------------------
st.set_page_config(page_title="PPT → PDF 일괄 변환", layout="centered")
st.title("PPT 폴더 → PDF로 한꺼번에 변환")

st.write("1) PPT 파일이 들어있는 폴더를 **ZIP으로 압축**해서 올리세요.")
st.write("2) 버튼을 누르면 PPT가 전부 PDF로 바뀌고, **PDFs.zip**으로 다운로드됩니다.")

# -----------------------------
# 파일 이름 정렬(선택: CH01, CH02 ... 순서로)
# -----------------------------
def chapter_key(path: Path):
    name = path.name.upper()

    # CH01, CH02 같은 패턴 우선
    m = re.search(r"CH\s*0*(\d+)", name)
    if m:
        return (0, int(m.group(1)), path.name)

    # 파일명에 숫자만 있어도 정렬
    m = re.search(r"0*(\d+)", name)
    if m:
        return (1, int(m.group(1)), path.name)

    # 그 외는 이름순
    return (2, 9999, path.name)


# -----------------------------
# PPT/PPTX -> PDF 변환 함수 (LibreOffice 사용)
# -----------------------------
def convert_ppt_to_pdf(input_path: Path, out_dir: Path) -> Path:
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(input_path),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"[변환 실패] {input_path.name}\n\n"
            f"stderr:\n{result.stderr}\n\nstdout:\n{result.stdout}"
        )

    pdf_path = out_dir / (input_path.stem + ".pdf")
    if not pdf_path.exists():
        # 혹시 파일명이 조금 다르게 나오는 경우 대비
        pdfs = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not pdfs:
            raise RuntimeError(f"PDF 결과 파일을 찾지 못했습니다: {input_path.name}")
        pdf_path = pdfs[0]

    return pdf_path


# -----------------------------
# 업로드: ZIP 1개
# -----------------------------
uploaded_zip = st.file_uploader("PPT 폴더를 ZIP으로 압축한 파일을 올리세요", type=["zip"])

if uploaded_zip is None:
    st.info("아직 ZIP 파일이 없어요. 폴더를 ZIP으로 압축해서 업로드하세요.")
    st.stop()


# -----------------------------
# 변환 버튼
# -----------------------------
if st.button("PDF로 일괄 변환 시작"):
    with st.spinner("변환 중입니다... 잠깐만요!"):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            in_dir = tmp_dir / "inputs"
            extract_dir = tmp_dir / "unzipped"
            pdf_dir = tmp_dir / "pdfs"

            in_dir.mkdir()
            extract_dir.mkdir()
            pdf_dir.mkdir()

            # 1) 업로드된 ZIP 저장
            zip_path = in_dir / uploaded_zip.name
            zip_path.write_bytes(uploaded_zip.getvalue())

            # 2) ZIP 풀기
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(extract_dir)

            # 3) PPT/PPTX 찾기(폴더 안 + 하위폴더까지)
            ppt_files = list(extract_dir.rglob("*.ppt")) + list(extract_dir.rglob("*.pptx"))

            if not ppt_files:
                st.error("ZIP 안에서 PPT/PPTX 파일을 찾지 못했습니다.")
                st.stop()

            # 4) CH01~ 순서로 정렬
            ppt_files.sort(key=chapter_key)

            st.write("변환할 파일 목록:")
            for p in ppt_files:
                st.write(f"- {p.name}")

            # 5) 변환 진행
            progress = st.progress(0)
            errors = []

            for i, ppt in enumerate(ppt_files, start=1):
                try:
                    convert_ppt_to_pdf(ppt, pdf_dir)
                except Exception as e:
                    errors.append(str(e))

                progress.progress(int(i / len(ppt_files) * 100))

            if errors:
                st.warning("일부 파일은 변환에 실패했습니다. 오류를 확인하세요.")
                for e in errors:
                    st.code(e)

            # 6) PDF를 ZIP으로 묶기
            out_zip = tmp_dir / "PDFs.zip"
            with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for pdf in sorted(pdf_dir.glob("*.pdf")):
                    z.write(pdf, arcname=pdf.name)

            st.success("완료! PDFs.zip을 다운로드하세요.")
            st.download_button(
                "PDFs.zip 다운로드",
                data=out_zip.read_bytes(),
                file_name="PDFs.zip",
                mime="application/zip",
            )
