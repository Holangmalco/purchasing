# -*- coding: utf-8 -*-
"""
구매부서 유틸리티 (PDF 종합 도구 & 파일명 복구) - Purchasing Department Utility & PDF Suite
- [탭 1] PDF 종합 도구 (오프라인 100% 안전 처리)
  * 사진(JPG, PNG 등) & 서류(Word, Excel, HWP, TXT) PDF 일괄 변환 (기본값: 각 파일별 개별 변환)
  * 여러 PDF 파일 하나로 깔끔하게 합치기 (병합)
  * PDF 페이지 분할 & 특정 페이지 추출
  * PDF 고화질 용량 압축 (다이어트)
  * 누워있는 PDF 똑바로 세우기 (일괄 90도/180도 회전)
- [탭 2] 한글 파일명 복구 (맥북 자소 분리 / 깨진 파일명 복구)
  * 맥북(macOS) NFD 자소 분리(ㅎㅏㄴㄱㅡㄹ) -> 윈도우 표준 NFC 결합 복구
  * 호환용 자모(Compatibility Jamo) 자모 결합 오토마타 완벽 지원
  * 압축 해제 인코딩(Mojibake) 오류 안전 복구
  * 우클릭 메뉴 등록 및 바탕화면 바로가기 지원
- 직관적인 2-탭 UI & 창 내부 드래그 앤 드롭(Drag & Drop) 완벽 지원
"""

import os
import sys
import unicodedata
import winreg
import ctypes
import winsound
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageOps, ImageDraw
import pypdfium2 as pdfium

try:
    import windnd
except ImportError:
    windnd = None

APP_TITLE = "구매부서 유틸리티 (파일명 복구 & PDF 종합 도구)"
REG_SHELL_NAME = "PurchasingDeptUtility"
MENU_LABEL = "[구매부서 유틸리티] 한글 파일명 복구"
MENU_BG_LABEL = "[구매부서 유틸리티] 폴더 내 한글 파일명 전체 복구"

# 지원 확장자 상수
IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.gif', '.tiff'}
DOC_EXTS = {'.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt', '.hwp', '.hwpx', '.txt', '.csv'}

# 한글 음절 조합 상수
CHOS = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']
JUNGS = ['ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ']
JONGS = ['', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']

COMPOUND_VOWELS = {
    ('ㅗ', 'ㅏ'): 'ㅘ', ('ㅗ', 'ㅐ'): 'ㅙ', ('ㅗ', 'ㅣ'): 'ㅚ',
    ('ㅜ', 'ㅓ'): 'ㅝ', ('ㅜ', 'ㅔ'): 'ㅞ', ('ㅜ', 'ㅣ'): 'ㅟ',
    ('ㅡ', 'ㅣ'): 'ㅢ'
}

COMPOUND_JONGS = {
    ('ㄱ', 'ㅅ'): 'ㄳ',
    ('ㄴ', 'ㅈ'): 'ㄵ', ('ㄴ', 'ㅎ'): 'ㄶ',
    ('ㄹ', 'ㄱ'): 'ㄺ', ('ㄹ', 'ㅁ'): 'ㄻ', ('ㄹ', 'ㅂ'): 'ㄼ',
    ('ㄹ', 'ㅅ'): 'ㄽ', ('ㄹ', 'ㅌ'): 'ㄾ', ('ㄹ', 'ㅍ'): 'ㄿ', ('ㄹ', 'ㅎ'): 'ㅀ',
    ('ㅂ', 'ㅅ'): 'ㅄ'
}


def is_cho(c: str) -> bool:
    return c in CHOS


def is_jung(c: str) -> bool:
    return c in JUNGS


def is_jong(c: str) -> bool:
    return c in JONGS and c != ''


def compose_compatibility_jamo(text: str) -> str:
    """
    브라우저, 메신저 등을 통해 다운로드되어 호환용 자모(U+3131~U+3163)로
    낱낱이 풀려버린 한글 자모들을 온전한 완성형 음절로 결합합니다.
    """
    res = []
    i = 0
    n = len(text)

    while i < n:
        c = text[i]
        if is_cho(c):
            cho = c
            if i + 1 < n and (is_jung(text[i + 1]) or (i + 2 < n and (text[i + 1], text[i + 2]) in COMPOUND_VOWELS)):
                i += 1
                jung = text[i]
                if i + 1 < n and (jung, text[i + 1]) in COMPOUND_VOWELS:
                    jung = COMPOUND_VOWELS[(jung, text[i + 1])]
                    i += 1

                jong = ''
                if i + 1 < n and is_jong(text[i + 1]):
                    next_is_vowel = (i + 2 < n and (is_jung(text[i + 2]) or (i + 3 < n and (text[i + 2], text[i + 3]) in COMPOUND_VOWELS)))
                    if not next_is_vowel:
                        if i + 2 < n and (text[i + 1], text[i + 2]) in COMPOUND_JONGS:
                            next_next_is_vowel = (i + 3 < n and is_jung(text[i + 3]))
                            if not next_next_is_vowel:
                                jong = COMPOUND_JONGS[(text[i + 1], text[i + 2])]
                                i += 2
                            else:
                                jong = text[i + 1]
                                i += 1
                        else:
                            jong = text[i + 1]
                            i += 1

                cho_idx = CHOS.index(cho)
                jung_idx = JUNGS.index(jung)
                jong_idx = JONGS.index(jong)
                syl = chr(0xAC00 + (cho_idx * 21 + jung_idx) * 28 + jong_idx)
                res.append(syl)
            else:
                res.append(cho)
        else:
            res.append(c)
        i += 1

    return ''.join(res)


def repair_name(name: str) -> str:
    """파일명의 한글 자소 분리 및 인코딩 오류를 안전하게 원복합니다."""
    step1 = unicodedata.normalize('NFC', name)
    step2 = compose_compatibility_jamo(step1)

    if step2 != name:
        return step2

    for enc in ['cp437', 'latin1', 'cp949']:
        try:
            cand = unicodedata.normalize('NFC', name.encode(enc).decode('utf-8'))
            orig_ko = sum(1 for c in name if '\uac00' <= c <= '\ud7a3')
            cand_ko = sum(1 for c in cand if '\uac00' <= c <= '\ud7a3')
            if cand_ko > orig_ko and '\ufffd' not in cand:
                return cand
        except Exception:
            pass

    return step2


def safe_rename(old_path: str, new_name: str) -> tuple[bool, str, str]:
    parent_dir = os.path.dirname(old_path)
    old_name = os.path.basename(old_path)

    if old_name == new_name:
        return False, old_name, old_name

    target_path = os.path.join(parent_dir, new_name)

    if os.path.exists(target_path) and os.path.abspath(target_path) != os.path.abspath(old_path):
        stem, ext = os.path.splitext(new_name)
        idx = 1
        while True:
            candidate_name = f"{stem} ({idx}){ext}"
            candidate_path = os.path.join(parent_dir, candidate_name)
            if not os.path.exists(candidate_path):
                target_path = candidate_path
                new_name = candidate_name
                break
            idx += 1

    try:
        temp_path = os.path.join(parent_dir, f"__tmp_fix_{idx if 'idx' in locals() else 0}_{old_name}")
        os.rename(old_path, temp_path)
        os.rename(temp_path, target_path)
        return True, old_name, new_name
    except Exception as e:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.rename(temp_path, old_path)
            except Exception:
                pass
        return False, old_name, f"실패 ({e})"


def process_path(target_path: str) -> list[tuple[str, str]]:
    results = []
    target_path = os.path.abspath(target_path)

    if not os.path.exists(target_path):
        return results

    if os.path.isfile(target_path):
        old_name = os.path.basename(target_path)
        new_name = repair_name(old_name)
        if new_name != old_name:
            success, old_n, new_n = safe_rename(target_path, new_name)
            if success:
                results.append((old_n, new_n))
        return results

    if os.path.isdir(target_path):
        for root, dirs, files in os.walk(target_path, topdown=False):
            for fname in files:
                fpath = os.path.join(root, fname)
                new_fname = repair_name(fname)
                if new_fname != fname:
                    success, old_n, new_n = safe_rename(fpath, new_fname)
                    if success:
                        results.append((old_n, new_n))

            for dname in dirs:
                dpath = os.path.join(root, dname)
                new_dname = repair_name(dname)
                if new_dname != dname:
                    success, old_n, new_n = safe_rename(dpath, new_dname)
                    if success:
                        results.append((old_n, new_n))

        root_parent = os.path.dirname(target_path)
        root_name = os.path.basename(target_path)
        new_root_name = repair_name(root_name)
        if new_root_name != root_name and root_parent:
            success, old_n, new_n = safe_rename(target_path, new_root_name)
            if success:
                results.append((old_n, new_root_name))

    return results


# =========================================================================
# PDF 도구 엔진 (변환, 병합, 분할/추출, 압축, 회전)
# =========================================================================

def get_unique_filename(target_path: str) -> str:
    if not os.path.exists(target_path):
        return target_path
    parent = os.path.dirname(target_path)
    stem, ext = os.path.splitext(os.path.basename(target_path))
    idx = 1
    while True:
        cand = os.path.join(parent, f"{stem} ({idx}){ext}")
        if not os.path.exists(cand):
            return cand
        idx += 1


def convert_images_to_pdf(image_paths: list[str], output_pdf: str = None, merge_into_single: bool = True) -> tuple[bool, list[str], str]:
    valid_paths = [p for p in image_paths if os.path.isfile(p) and os.path.splitext(p)[1].lower() in IMAGE_EXTS]
    if not valid_paths:
        return False, [], "변환할 수 있는 사진 파일(JPG, PNG 등)이 없습니다."

    created_pdfs = []
    try:
        if merge_into_single:
            pil_images = []
            for path in valid_paths:
                img = Image.open(path)
                img = ImageOps.exif_transpose(img)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                pil_images.append(img)

            if not output_pdf:
                first_dir = os.path.dirname(valid_paths[0])
                first_name = os.path.splitext(os.path.basename(valid_paths[0]))[0]
                first_name = repair_name(first_name)
                if len(valid_paths) > 1:
                    target = os.path.join(first_dir, f"{first_name}_외{len(valid_paths)-1}장_합본.pdf")
                else:
                    target = os.path.join(first_dir, f"{first_name}.pdf")
                output_pdf = get_unique_filename(target)

            pil_images[0].save(
                output_pdf,
                "PDF",
                resolution=100.0,
                save_all=True,
                append_images=pil_images[1:]
            )
            created_pdfs.append(output_pdf)
            return True, created_pdfs, f"총 {len(valid_paths)}장의 사진을 1개의 PDF로 묶었습니다."
        else:
            for path in valid_paths:
                img = Image.open(path)
                img = ImageOps.exif_transpose(img)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                elif img.mode != "RGB":
                    img = img.convert("RGB")

                p_dir = os.path.dirname(path)
                stem = os.path.splitext(os.path.basename(path))[0]
                stem = repair_name(stem)
                out_p = get_unique_filename(os.path.join(p_dir, f"{stem}.pdf"))
                img.save(out_p, "PDF", resolution=100.0)
                created_pdfs.append(out_p)

            return True, created_pdfs, f"총 {len(created_pdfs)}개의 사진을 각각 PDF로 개별 변환했습니다."
    except Exception as e:
        return False, created_pdfs, f"사진 변환 중 오류: {e}"


def convert_doc_to_pdf(doc_path: str, output_pdf: str = None, excel_fit_mode: str = "width") -> tuple[bool, str, str]:
    if not os.path.isfile(doc_path):
        return False, "", "파일이 존재하지 않습니다."

    ext = os.path.splitext(doc_path)[1].lower()
    p_dir = os.path.dirname(doc_path)
    stem = os.path.splitext(os.path.basename(doc_path))[0]
    stem = repair_name(stem)

    if not output_pdf:
        output_pdf = get_unique_filename(os.path.join(p_dir, f"{stem}.pdf"))

    abs_in = os.path.abspath(doc_path)
    abs_out = os.path.abspath(output_pdf)

    # 텍스트 문서 (.txt, .csv)
    if ext in ('.txt', '.csv'):
        try:
            with open(abs_in, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
            page_w, page_h = 1240, 1754
            margin = 60
            line_height = 28
            max_lines_per_page = (page_h - margin * 2) // line_height

            pages_img = []
            curr_lines = []
            for line in lines:
                curr_lines.append(line.rstrip('\r\n'))
                if len(curr_lines) >= max_lines_per_page:
                    img = Image.new("RGB", (page_w, page_h), color="white")
                    draw = ImageDraw.Draw(img)
                    y = margin
                    for l in curr_lines:
                        draw.text((margin, y), l, fill="black")
                        y += line_height
                    pages_img.append(img)
                    curr_lines = []

            if curr_lines or not pages_img:
                img = Image.new("RGB", (page_w, page_h), color="white")
                draw = ImageDraw.Draw(img)
                y = margin
                for l in curr_lines:
                    draw.text((margin, y), l, fill="black")
                    y += line_height
                pages_img.append(img)

            pages_img[0].save(
                abs_out,
                "PDF",
                resolution=150.0,
                save_all=True,
                append_images=pages_img[1:]
            )
            return True, abs_out, f"텍스트 문서 변환 완료 ({ext.upper()} -> PDF)"
        except Exception as e:
            return False, "", f"텍스트 변환 실패: {e}"

    # 오피스 및 한글 문서 PowerShell 자동화 변환
    ps_safe_in = abs_in.replace("'", "''")
    ps_safe_out = abs_out.replace("'", "''")

    if ext in ('.docx', '.doc'):
        ps = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{ps_safe_in}', $false, $true)
    $doc.ExportAsFixedFormat('{ps_safe_out}', 17)
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR: $($_.Exception.Message)"
}} finally {{
    if ($doc -ne $null) {{ $doc.Close([ref]$false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }}
    if ($word -ne $null) {{ $word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''
    elif ext in ('.xlsx', '.xls'):
        fit_ps = ""
        if excel_fit_mode == "width":
            fit_ps = """
    foreach ($ws in $wb.Worksheets) {
        try {
            $ws.PageSetup.Zoom = $false
            $ws.PageSetup.FitToPagesWide = 1
            $ws.PageSetup.FitToPagesTall = $false
        } catch {}
    }
"""
        elif excel_fit_mode == "sheet":
            fit_ps = """
    foreach ($ws in $wb.Worksheets) {
        try {
            $ws.PageSetup.Zoom = $false
            $ws.PageSetup.FitToPagesWide = 1
            $ws.PageSetup.FitToPagesTall = 1
        } catch {}
    }
"""

        ps = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$excel = $null
$wb = $null
try {{
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AlertBeforeOverwriting = $false
    $wb = $excel.Workbooks.Open('{ps_safe_in}', 0, $true)
{fit_ps}
    $wb.ExportAsFixedFormat(0, '{ps_safe_out}')
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR: $($_.Exception.Message)"
}} finally {{
    if ($wb -ne $null) {{ $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }}
    if ($excel -ne $null) {{ $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''
    elif ext in ('.pptx', '.ppt'):
        ps = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ppt = $null
$pres = $null
try {{
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open('{ps_safe_in}', $true, $false, $false)
    $pres.SaveAs('{ps_safe_out}', 32)
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR: $($_.Exception.Message)"
}} finally {{
    if ($pres -ne $null) {{ $pres.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }}
    if ($ppt -ne $null) {{ $ppt.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''
    elif ext in ('.hwp', '.hwpx'):
        ps = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$hwp = $null
try {{
    $hwp = New-Object -ComObject HWPFrame.HwpObject
    $hwp.Open('{ps_safe_in}', "HWP", "forceopen:true")
    $hwp.SaveAs('{ps_safe_out}', "PDF")
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR: $($_.Exception.Message)"
}} finally {{
    if ($hwp -ne $null) {{ $hwp.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''
    else:
        return False, "", f"지원되지 않는 문서 확장자입니다: {ext}"

    tf = os.path.join(tempfile.gettempdir(), f"_conv_{os.getpid()}_{hash(abs_in) & 0xffffffff}.ps1")
    try:
        with open(tf, 'w', encoding='utf-8-sig') as f:
            f.write(ps)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        res = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', tf],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            startupinfo=startupinfo,
            timeout=120
        )
        if "SUCCESS" in res.stdout and os.path.exists(abs_out):
            fit_desc = ""
            if ext in ('.xlsx', '.xls'):
                if excel_fit_mode == "width":
                    fit_desc = " [너비 1페이지 맞춤]"
                elif excel_fit_mode == "sheet":
                    fit_desc = " [시트 전체 1장 맞춤]"
            return True, abs_out, f"문서 변환 성공 ({ext.upper()} -> PDF{fit_desc})"
        else:
            err = res.stdout.strip() or res.stderr.strip()
            if "80040154" in err or "클래스가 등록되지 않았습니다" in err:
                detail = "PC에 해당 프로그램(MS Office/한글)이 설치되어 있지 않습니다."
            elif "800A03EC" in err:
                detail = "엑셀 인쇄/PDF 설정 오류 (Windows 'Microsoft Print to PDF' 가상 프린터 확인 필요)"
            else:
                detail = "해당 프로그램이 PC에 설치되어 있어야 합니다"
            return False, "", f"변환 실패: {err} ({detail})"
    except Exception as e:
        return False, "", f"변환 중 오류 발생: {e}"
    finally:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except Exception:
                pass


def merge_pdfs(pdf_paths: list[str], output_pdf: str = None, sort_by_name: bool = True) -> tuple[bool, str, str]:
    valid_pdfs = [p for p in pdf_paths if os.path.isfile(p) and p.lower().endswith('.pdf')]
    if len(valid_pdfs) < 2:
        return False, "", "PDF를 합치려면 최소 2개 이상의 PDF 파일이 필요합니다."

    if sort_by_name:
        valid_pdfs.sort(key=lambda x: os.path.basename(x).lower())

    try:
        dest_doc = pdfium.PdfDocument.new()
        total_pages = 0
        for p in valid_pdfs:
            src_doc = pdfium.PdfDocument(p)
            total_pages += len(src_doc)
            dest_doc.import_pages(src_doc)
            src_doc.close()

        if not output_pdf:
            p_dir = os.path.dirname(valid_pdfs[0])
            first_name = os.path.splitext(os.path.basename(valid_pdfs[0]))[0]
            first_name = repair_name(first_name)
            target = os.path.join(p_dir, f"{first_name}_외{len(valid_pdfs)-1}개_합본.pdf")
            output_pdf = get_unique_filename(target)

        dest_doc.save(output_pdf)
        dest_doc.close()
        return True, output_pdf, f"총 {len(valid_pdfs)}개 PDF(전체 {total_pages}페이지)를 성공적으로 하나로 합쳤습니다."
    except Exception as e:
        return False, "", f"PDF 병합 중 오류: {e}"


def parse_page_ranges(range_str: str, max_pages: int) -> list[int]:
    if not range_str or not range_str.strip():
        return []
    pages = set()
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            sub = part.split('-')
            if len(sub) == 2 and sub[0].strip().isdigit() and sub[1].strip().isdigit():
                start = int(sub[0].strip())
                end = int(sub[1].strip())
                for p in range(start, end + 1):
                    if 1 <= p <= max_pages:
                        pages.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= max_pages:
                pages.add(p - 1)
    return sorted(list(pages))


def split_or_extract_pdf(pdf_path: str, range_str: str = "") -> tuple[bool, list[str], str]:
    if not os.path.isfile(pdf_path) or not pdf_path.lower().endswith('.pdf'):
        return False, [], "유효한 PDF 파일이 아닙니다."

    created = []
    try:
        src_doc = pdfium.PdfDocument(pdf_path)
        total_pages = len(src_doc)
        p_dir = os.path.dirname(pdf_path)
        stem = os.path.splitext(os.path.basename(pdf_path))[0]
        stem = repair_name(stem)

        indices = parse_page_ranges(range_str, total_pages)
        if indices:
            dest_doc = pdfium.PdfDocument.new()
            dest_doc.import_pages(src_doc, indices)
            out_p = get_unique_filename(os.path.join(p_dir, f"{stem}_추출({len(indices)}p).pdf"))
            dest_doc.save(out_p)
            dest_doc.close()
            created.append(out_p)
            src_doc.close()
            return True, created, f"전체 {total_pages}페이지 중 {len(indices)}개 페이지 추출 완료!"
        else:
            for idx in range(total_pages):
                dest_doc = pdfium.PdfDocument.new()
                dest_doc.import_pages(src_doc, [idx])
                out_p = get_unique_filename(os.path.join(p_dir, f"{stem}_p{idx+1:03d}.pdf"))
                dest_doc.save(out_p)
                dest_doc.close()
                created.append(out_p)
            src_doc.close()
            return True, created, f"전체 {total_pages}페이지를 1장씩 낱개 PDF로 분할 완료!"
    except Exception as e:
        return False, created, f"분할/추출 중 오류: {e}"


def rotate_pdf(pdf_path: str, angle: int = 90) -> tuple[bool, str, str]:
    if not os.path.isfile(pdf_path) or not pdf_path.lower().endswith('.pdf'):
        return False, "", "유효한 PDF 파일이 아닙니다."

    try:
        doc = pdfium.PdfDocument(pdf_path)
        total_pages = len(doc)
        for page in doc:
            cur = page.get_rotation()
            page.set_rotation((cur + angle) % 360)

        p_dir = os.path.dirname(pdf_path)
        stem = os.path.splitext(os.path.basename(pdf_path))[0]
        stem = repair_name(stem)
        out_p = get_unique_filename(os.path.join(p_dir, f"{stem}_회전({angle}도).pdf"))
        doc.save(out_p)
        doc.close()
        return True, out_p, f"총 {total_pages}페이지를 {angle}도 똑바로 회전 완료!"
    except Exception as e:
        return False, "", f"PDF 회전 중 오류: {e}"


def compress_pdf(pdf_path: str, dpi: int = 150, quality: int = 75) -> tuple[bool, str, str]:
    if not os.path.isfile(pdf_path) or not pdf_path.lower().endswith('.pdf'):
        return False, "", "유효한 PDF 파일이 아닙니다."

    try:
        orig_size = os.path.getsize(pdf_path)
        doc = pdfium.PdfDocument(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            doc.close()
            return False, "", "빈 PDF 문서입니다."

        pil_pages = []
        scale = dpi / 72.0
        for page in doc:
            img = page.render(scale=scale).to_pil()
            if img.mode != 'RGB':
                img = img.convert('RGB')
            pil_pages.append(img)
        doc.close()

        p_dir = os.path.dirname(pdf_path)
        stem = os.path.splitext(os.path.basename(pdf_path))[0]
        stem = repair_name(stem)
        out_p = get_unique_filename(os.path.join(p_dir, f"{stem}_압축.pdf"))

        pil_pages[0].save(
            out_p,
            "PDF",
            resolution=float(dpi),
            quality=quality,
            save_all=True,
            append_images=pil_pages[1:]
        )

        new_size = os.path.getsize(out_p)
        ratio = (1.0 - (new_size / orig_size)) * 100 if orig_size > 0 else 0
        orig_mb = orig_size / (1024 * 1024)
        new_mb = new_size / (1024 * 1024)

        msg = f"압축 완료 ({total_pages}p): {orig_mb:.2f}MB -> {new_mb:.2f}MB ({ratio:.1f}% 절감)"
        return True, out_p, msg
    except Exception as e:
        return False, "", f"PDF 압축 중 오류: {e}"


# =========================================================================
# 윈도우 우클릭 메뉴 및 바로가기 관련 함수
# =========================================================================

def get_current_executable_command() -> str:
    current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    exe_candidate = os.path.join(current_dir, "구매부서_유틸리티.exe")
    if not os.path.exists(exe_candidate):
        exe_candidate = os.path.join(current_dir, "맥북한글복구기.exe")
    if os.path.exists(exe_candidate):
        return f'"{exe_candidate}" "%1"'
    if getattr(sys, 'frozen', False):
        exe_path = os.path.abspath(sys.executable)
        return f'"{exe_path}" "%1"'
    else:
        py_dir = os.path.dirname(sys.executable)
        pyw = os.path.join(py_dir, "pythonw.exe")
        if not os.path.exists(pyw):
            pyw = sys.executable
        script_path = os.path.abspath(__file__)
        return f'"{pyw}" "{script_path}" "%1"'


def get_current_bg_command() -> str:
    current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    exe_candidate = os.path.join(current_dir, "구매부서_유틸리티.exe")
    if not os.path.exists(exe_candidate):
        exe_candidate = os.path.join(current_dir, "맥북한글복구기.exe")
    if os.path.exists(exe_candidate):
        return f'"{exe_candidate}" "%V"'
    if getattr(sys, 'frozen', False):
        exe_path = os.path.abspath(sys.executable)
        return f'"{exe_path}" "%V"'
    else:
        py_dir = os.path.dirname(sys.executable)
        pyw = os.path.join(py_dir, "pythonw.exe")
        if not os.path.exists(pyw):
            pyw = sys.executable
        script_path = os.path.abspath(__file__)
        return f'"{pyw}" "{script_path}" "%V"'


def register_context_menu() -> tuple[bool, str]:
    try:
        cmd_single = get_current_executable_command()
        cmd_bg = get_current_bg_command()
        current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        exe_candidate = os.path.join(current_dir, "구매부서_유틸리티.exe")
        if not os.path.exists(exe_candidate):
            exe_candidate = os.path.join(current_dir, "맥북한글복구기.exe")
        icon_path = f'"{exe_candidate}",0' if os.path.exists(exe_candidate) else "imageres.dll,112"

        targets = [
            (r"Software\Classes\*\shell\\" + REG_SHELL_NAME, MENU_LABEL, cmd_single),
            (r"Software\Classes\Directory\shell\\" + REG_SHELL_NAME, MENU_LABEL, cmd_single),
            (r"Software\Classes\Directory\Background\shell\\" + REG_SHELL_NAME, MENU_BG_LABEL, cmd_bg),
        ]

        for subkey, label, cmd in targets:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, subkey)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, label)
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon_path)
            winreg.CloseKey(key)

            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, subkey + r"\command")
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(cmd_key)

        return True, "윈도우 마우스 우클릭 메뉴에 성공적으로 등록되었습니다!"
    except Exception as e:
        return False, f"등록 중 오류가 발생했습니다: {e}"


def unregister_context_menu() -> tuple[bool, str]:
    try:
        targets = [
            r"Software\Classes\*\shell\\" + REG_SHELL_NAME,
            r"Software\Classes\Directory\shell\\" + REG_SHELL_NAME,
            r"Software\Classes\Directory\Background\shell\\" + REG_SHELL_NAME,
            r"Software\Classes\*\shell\FixMacHangul",
            r"Software\Classes\Directory\shell\FixMacHangul",
            r"Software\Classes\Directory\Background\shell\FixMacHangul",
        ]

        def delete_key_recursive(root, subkey):
            try:
                key = winreg.OpenKey(root, subkey, 0, winreg.KEY_ALL_ACCESS)
                while True:
                    try:
                        child = winreg.EnumKey(key, 0)
                        delete_key_recursive(root, f"{subkey}\\{child}")
                    except OSError:
                        break
                winreg.CloseKey(key)
                winreg.DeleteKey(root, subkey)
            except FileNotFoundError:
                pass

        for subkey in targets:
            delete_key_recursive(winreg.HKEY_CURRENT_USER, subkey)

        return True, "우클릭 메뉴가 깔끔하게 제거되었습니다."
    except Exception as e:
        return False, f"제거 중 오류가 발생했습니다: {e}"


def create_desktop_shortcut() -> tuple[bool, str]:
    try:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        lnk_path = os.path.join(desktop, '구매부서 유틸리티.lnk')
        legacy_lnk = os.path.join(desktop, '맥북 한글 파일명 복구.lnk')
        if os.path.exists(legacy_lnk):
            try:
                os.remove(legacy_lnk)
            except Exception:
                pass
        current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        exe_path = os.path.join(current_dir, "구매부서_유틸리티.exe")
        if not os.path.exists(exe_path):
            legacy_exe = os.path.join(current_dir, "맥북한글복구기.exe")
            if os.path.exists(legacy_exe):
                exe_path = legacy_exe
        if not os.path.exists(exe_path):
            exe_path = sys.executable

        vbs = f'''Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = "{lnk_path}"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "{exe_path}"
oLink.WorkingDirectory = "{current_dir}"
oLink.Description = "구매부서 유틸리티 (파일명 복구 & PDF 종합 도구)"
oLink.Save
'''
        vbs_path = os.path.join(tempfile.gettempdir(), '_create_shortcut.vbs')
        with open(vbs_path, 'w', encoding='cp949', errors='ignore') as f:
            f.write(vbs)

        subprocess.run(['cscript', '//nologo', vbs_path], check=True)
        if os.path.exists(vbs_path):
            try:
                os.remove(vbs_path)
            except Exception:
                pass

        return True, "바탕화면에 바로가기가 생성되었습니다!\n바탕화면 아이콘 위로 파일을 끌어다 놓아도 바로 복구됩니다."
    except Exception as e:
        return False, f"바로가기 생성 중 오류: {e}"


def show_auto_alert(title: str, message: str, auto_close_sec: int = 3):
    root = tk.Tk()
    root.title(title)
    root.attributes("-topmost", True)
    root.resizable(False, False)

    w, h = 440, 190
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.configure(bg="#F8F9FA")

    frame = tk.Frame(root, bg="#F8F9FA", padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    lbl_msg = tk.Label(
        frame,
        text=message,
        font=("Malgun Gothic", 10),
        bg="#F8F9FA",
        fg="#212529",
        justify=tk.CENTER,
        wraplength=400
    )
    lbl_msg.pack(expand=True)

    btn_ok = tk.Button(
        frame,
        text=f"확인 ({auto_close_sec}초 후 자동 닫힘)",
        font=("Malgun Gothic", 9, "bold"),
        bg="#004797",
        fg="white",
        activebackground="#003366",
        activeforeground="white",
        relief=tk.FLAT,
        padx=15,
        pady=5,
        command=root.destroy,
        cursor="hand2"
    )
    btn_ok.pack(side=tk.BOTTOM)

    remaining = [auto_close_sec]
    def countdown():
        if remaining[0] <= 1:
            root.destroy()
        else:
            remaining[0] -= 1
            btn_ok.config(text=f"확인 ({remaining[0]}초 후 자동 닫힘)")
            root.after(1000, countdown)

    root.after(1000, countdown)
    root.bind("<Return>", lambda e: root.destroy())
    root.bind("<Escape>", lambda e: root.destroy())

    winsound.MessageBeep(winsound.MB_ICONASTERISK)
    root.mainloop()


# =========================================================================
# 메인 GUI (2-탭 구조: 한글 복구 / PDF 도구)
# =========================================================================

def launch_main_gui():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("740x680")
    root.minsize(680, 600)
    root.configure(bg="#F4F6F9")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", font=("Malgun Gothic", 10), padding=6)
    style.configure("TNotebook", background="#F4F6F9")
    style.configure("TNotebook.Tab", font=("Malgun Gothic", 10, "bold"), padding=[20, 8])

    # 1. 상단 공통 헤더
    header = tk.Frame(root, bg="#004797", padx=20, pady=14)
    header.pack(fill=tk.X)

    lbl_title = tk.Label(
        header,
        text="💼 구매부서 유틸리티",
        font=("Malgun Gothic", 13, "bold"),
        fg="white",
        bg="#004797"
    )
    lbl_title.pack(anchor="w")

    # 공통 로그 출력 함수
    def log(msg):
        txt_log.config(state=tk.NORMAL)
        txt_log.insert(tk.END, msg + "\n")
        txt_log.see(tk.END)
        txt_log.config(state=tk.DISABLED)

    # 2. 탭 컨테이너 (Notebook)
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=False, padx=15, pady=(10, 5))

    tab_fixer = tk.Frame(notebook, bg="#F4F6F9", padx=10, pady=10)
    tab_pdf = tk.Frame(notebook, bg="#F4F6F9", padx=10, pady=10)

    # 첫 화면을 PDF 종합 도구로 배치하고, 파일명 복구를 뒤 탭으로 배치
    notebook.add(tab_pdf, text="  📑 PDF 도구  ")
    notebook.add(tab_fixer, text="  🛠️ 파일명 복구  ")
    notebook.select(tab_pdf)

    # ---------------------------------------------------------------------
    # [탭 1] 맥북 한글 파일명 복구 탭 내용
    # ---------------------------------------------------------------------
    def handle_fixer_items(item_list, source_label="선택"):
        all_res = []
        for item in item_list:
            if isinstance(item, bytes):
                try:
                    item = item.decode('utf-8')
                except Exception:
                    item = item.decode('mbcs', errors='ignore')
            if os.path.exists(item):
                res = process_path(item)
                all_res.extend(res)

        if all_res:
            log(f"✅ [파일명 복구: {source_label}] 총 {len(all_res)}개 항목 복구 완료:")
            for old_n, new_n in all_res:
                log(f"   • {old_n}  ->  {new_n}")
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("복구 완료", f"총 {len(all_res)}개 파일/폴더명이 정상 복구되었습니다!")
        else:
            log(f"✨ [파일명 복구: {source_label}] 선택하신 항목의 한글 이름이 이미 모두 정상입니다.")
            winsound.MessageBeep(winsound.MB_ICONINFORMATION)
            messagebox.showinfo("확인", "깨진 한글 파일명이 없거나 이미 정상입니다.")

    drop_frame_fixer = tk.Frame(tab_fixer, bg="#EBF3FB", bd=2, relief=tk.GROOVE, padx=20, pady=16)
    drop_frame_fixer.pack(fill=tk.X, pady=(4, 6))

    lbl_drop_icon1 = tk.Label(drop_frame_fixer, text="📥", font=("Segoe UI Emoji", 20), bg="#EBF3FB")
    lbl_drop_icon1.pack()

    lbl_drop_title1 = tk.Label(
        drop_frame_fixer,
        text="파일 또는 폴더를 여기에 끌어다 놓으세요",
        font=("Malgun Gothic", 10, "bold"),
        fg="#004797",
        bg="#EBF3FB"
    )
    lbl_drop_title1.pack(pady=(3, 0))

    # 폴더/파일 수동 선택 버튼
    btn_frame_fixer = tk.Frame(tab_fixer, bg="#F4F6F9")
    btn_frame_fixer.pack(fill=tk.X, pady=(4, 6))

    def on_pick_folder():
        folder = filedialog.askdirectory(title="복구할 폴더를 선택하세요")
        if folder:
            handle_fixer_items([folder], "폴더 검사")

    def on_pick_files():
        files = filedialog.askopenfilenames(title="복구할 파일들을 선택하세요")
        if files:
            handle_fixer_items(files, "파일 검사")

    btn_folder = tk.Button(
        btn_frame_fixer,
        text="📁 폴더 선택 복구",
        font=("Malgun Gothic", 9, "bold"),
        bg="#004797",
        fg="white",
        activebackground="#003366",
        activeforeground="white",
        padx=10,
        pady=6,
        relief=tk.FLAT,
        cursor="hand2",
        command=on_pick_folder
    )
    btn_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

    btn_files = tk.Button(
        btn_frame_fixer,
        text="📄 개별 파일 선택 복구",
        font=("Malgun Gothic", 9, "bold"),
        bg="#2B5B84",
        fg="white",
        activebackground="#1E3E5C",
        activeforeground="white",
        padx=10,
        pady=6,
        relief=tk.FLAT,
        cursor="hand2",
        command=on_pick_files
    )
    btn_files.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

    # 편의 기능
    quick_frame = tk.LabelFrame(
        tab_fixer,
        text=" 원클릭 환경 설정 ",
        font=("Malgun Gothic", 9, "bold"),
        bg="#F4F6F9",
        fg="#333",
        padx=10,
        pady=6
    )
    quick_frame.pack(fill=tk.X, pady=(2, 4))

    qbtn_frame = tk.Frame(quick_frame, bg="#F4F6F9")
    qbtn_frame.pack(fill=tk.X)

    def on_create_shortcut():
        ok, msg = create_desktop_shortcut()
        if ok:
            log("📌 " + msg.split("\n")[0])
            messagebox.showinfo("바탕화면 바로가기", msg)
        else:
            log("❌ " + msg)
            messagebox.showerror("오류", msg)

    def on_register():
        ok, msg = register_context_menu()
        if ok:
            log("✨ " + msg)
            messagebox.showinfo("우클릭 메뉴 등록", msg)
        else:
            log("❌ " + msg)
            messagebox.showerror("오류", msg)

    def on_unregister():
        ok, msg = unregister_context_menu()
        if ok:
            log("🗑️ " + msg)
            messagebox.showinfo("우클릭 메뉴 제거", msg)
        else:
            log("❌ " + msg)
            messagebox.showerror("오류", msg)

    btn_sc = tk.Button(
        qbtn_frame,
        text="📌 바탕화면 바로가기 생성",
        font=("Malgun Gothic", 8, "bold"),
        bg="#17A2B8",
        fg="white",
        relief=tk.FLAT,
        cursor="hand2",
        command=on_create_shortcut
    )
    btn_sc.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

    btn_reg = tk.Button(
        qbtn_frame,
        text="✨ 우클릭 메뉴 등록",
        font=("Malgun Gothic", 8, "bold"),
        bg="#28A745",
        fg="white",
        relief=tk.FLAT,
        cursor="hand2",
        command=on_register
    )
    btn_reg.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 3))

    btn_unreg = tk.Button(
        qbtn_frame,
        text="🗑️ 우클릭 메뉴 제거",
        font=("Malgun Gothic", 8),
        bg="#6C757D",
        fg="white",
        relief=tk.FLAT,
        cursor="hand2",
        command=on_unregister
    )
    btn_unreg.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))

    # ---------------------------------------------------------------------
    # [탭 2] PDF 만능 도구 탭 내용
    # ---------------------------------------------------------------------
    pdf_mode = tk.StringVar(value="conv")
    img_merge_var = tk.BooleanVar(value=False)
    excel_fit_var = tk.StringVar(value="width")
    pdf_sort_var = tk.BooleanVar(value=True)
    split_range_var = tk.StringVar(value="")
    compress_level_var = tk.StringVar(value="normal")
    rotate_angle_var = tk.IntVar(value=90)

    # 1) 작업 모드 선택 버튼 바
    mode_frame = tk.Frame(tab_pdf, bg="#F4F6F9")
    mode_frame.pack(fill=tk.X, pady=(0, 6))

    mode_buttons = {}

    modes_def = [
        ("conv", "📷 PDF 변환"),
        ("merge", "📑 PDF 병합"),
        ("split", "✂️ 페이지 분할"),
        ("compress", "📉 용량 압축"),
        ("rotate", "🔄 방향 회전"),
    ]

    def update_pdf_ui():
        m = pdf_mode.get()
        for k, btn in mode_buttons.items():
            if k == m:
                btn.config(bg="#004797", fg="white", font=("Malgun Gothic", 9, "bold"))
            else:
                btn.config(bg="#E9ECEF", fg="#495057", font=("Malgun Gothic", 9))

        opt_conv.pack_forget()
        opt_merge.pack_forget()
        opt_split.pack_forget()
        opt_compress.pack_forget()
        opt_rotate.pack_forget()

        if m == "conv":
            opt_conv.pack(fill=tk.X, pady=(2, 4))
            lbl_drop_icon2.config(text="📷")
            lbl_drop_title2.config(text="변환할 사진/서류를 여기에 끌어다 놓으세요")
            lbl_drop_desc2.config(text="")
            btn_pick_pdf.config(text="📄 파일 직접 선택")
        elif m == "merge":
            opt_merge.pack(fill=tk.X, pady=(2, 4))
            lbl_drop_icon2.config(text="📑")
            lbl_drop_title2.config(text="합칠 PDF들을 여기에 끌어다 놓으세요")
            lbl_drop_desc2.config(text="")
            btn_pick_pdf.config(text="📑 PDF 직접 선택")
        elif m == "split":
            opt_split.pack(fill=tk.X, pady=(2, 4))
            lbl_drop_icon2.config(text="✂️")
            lbl_drop_title2.config(text="분할할 PDF를 여기에 끌어다 놓으세요")
            lbl_drop_desc2.config(text="")
            btn_pick_pdf.config(text="✂️ PDF 직접 선택")
        elif m == "compress":
            opt_compress.pack(fill=tk.X, pady=(2, 4))
            lbl_drop_icon2.config(text="📉")
            lbl_drop_title2.config(text="압축할 PDF를 여기에 끌어다 놓으세요")
            lbl_drop_desc2.config(text="")
            btn_pick_pdf.config(text="📉 PDF 직접 선택")
        elif m == "rotate":
            opt_rotate.pack(fill=tk.X, pady=(2, 4))
            lbl_drop_icon2.config(text="🔄")
            lbl_drop_title2.config(text="회전할 PDF를 여기에 끌어다 놓으세요")
            lbl_drop_desc2.config(text="")
            btn_pick_pdf.config(text="🔄 PDF 직접 선택")

    for key, text in modes_def:
        b = tk.Button(
            mode_frame,
            text=text,
            relief=tk.FLAT,
            padx=6,
            pady=6,
            cursor="hand2",
            command=lambda k=key: [pdf_mode.set(k), update_pdf_ui()]
        )
        b.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        mode_buttons[key] = b

    # 2) 모드별 세부 옵션 프레임 (동적 전환)
    opt_container = tk.Frame(tab_pdf, bg="#F4F6F9", height=32)
    opt_container.pack(fill=tk.X)

    # (1) 사진/문서 옵션
    opt_conv = tk.Frame(opt_container, bg="#F4F6F9")
    chk_img_merge = tk.Checkbutton(
        opt_conv,
        text="사진들을 1개 PDF로 묶어서 저장",
        variable=img_merge_var,
        bg="#F4F6F9",
        font=("Malgun Gothic", 9, "bold"),
        fg="#004797"
    )
    chk_img_merge.pack(side=tk.LEFT, padx=(0, 10))

    lbl_excel_fit = tk.Label(
        opt_conv,
        text="|  📊 엑셀 맞춤:",
        font=("Malgun Gothic", 9, "bold"),
        fg="#333",
        bg="#F4F6F9"
    )
    lbl_excel_fit.pack(side=tk.LEFT, padx=(5, 3))

    rb_fit_width = tk.Radiobutton(
        opt_conv,
        text="너비 1페이지 맞춤 (가로 잘림 방지, 기본)",
        variable=excel_fit_var,
        value="width",
        bg="#F4F6F9",
        font=("Malgun Gothic", 9)
    )
    rb_fit_width.pack(side=tk.LEFT, padx=3)

    rb_fit_sheet = tk.Radiobutton(
        opt_conv,
        text="시트 전체 1장 맞춤",
        variable=excel_fit_var,
        value="sheet",
        bg="#F4F6F9",
        font=("Malgun Gothic", 9)
    )
    rb_fit_sheet.pack(side=tk.LEFT, padx=3)

    rb_fit_none = tk.Radiobutton(
        opt_conv,
        text="원본 유지",
        variable=excel_fit_var,
        value="none",
        bg="#F4F6F9",
        font=("Malgun Gothic", 9)
    )
    rb_fit_none.pack(side=tk.LEFT, padx=3)

    # (2) 합치기 옵션
    opt_merge = tk.Frame(opt_container, bg="#F4F6F9")
    chk_pdf_sort = tk.Checkbutton(
        opt_merge,
        text="파일명 가나다순으로 정렬",
        variable=pdf_sort_var,
        bg="#F4F6F9",
        font=("Malgun Gothic", 9, "bold"),
        fg="#004797"
    )
    chk_pdf_sort.pack(anchor="w")

    # (3) 분할/추출 옵션
    opt_split = tk.Frame(opt_container, bg="#F4F6F9")
    lbl_split_range = tk.Label(opt_split, text="추출할 페이지:", font=("Malgun Gothic", 9, "bold"), bg="#F4F6F9")
    lbl_split_range.pack(side=tk.LEFT)
    ent_split_range = tk.Entry(opt_split, textvariable=split_range_var, font=("Malgun Gothic", 9), width=18)
    ent_split_range.pack(side=tk.LEFT, padx=6)
    lbl_split_hint = tk.Label(
        opt_split,
        text="예: 1-3, 5  (💡 빈칸으로 비워두면 1장씩 낱개로 모두 분할됩니다)",
        font=("Malgun Gothic", 8),
        fg="#666",
        bg="#F4F6F9"
    )
    lbl_split_hint.pack(side=tk.LEFT)

    # (4) 압축 옵션
    opt_compress = tk.Frame(opt_container, bg="#F4F6F9")
    lbl_comp = tk.Label(opt_compress, text="압축 강도:", font=("Malgun Gothic", 9, "bold"), bg="#F4F6F9")
    lbl_comp.pack(side=tk.LEFT, padx=(0, 8))
    rb_comp1 = tk.Radiobutton(
        opt_compress,
        text="보통 압축 (150 DPI - 추천 / 선명함 유지)",
        variable=compress_level_var,
        value="normal",
        bg="#F4F6F9",
        font=("Malgun Gothic", 9)
    )
    rb_comp1.pack(side=tk.LEFT, padx=4)
    rb_comp2 = tk.Radiobutton(
        opt_compress,
        text="강한 압축 (100 DPI - 메일 첨부 및 웹 업로드용)",
        variable=compress_level_var,
        value="high",
        bg="#F4F6F9",
        font=("Malgun Gothic", 9)
    )
    rb_comp2.pack(side=tk.LEFT, padx=4)

    # (5) 회전 옵션
    opt_rotate = tk.Frame(opt_container, bg="#F4F6F9")
    lbl_rot = tk.Label(opt_rotate, text="회전 각도:", font=("Malgun Gothic", 9, "bold"), bg="#F4F6F9")
    lbl_rot.pack(side=tk.LEFT, padx=(0, 8))
    rb_rot1 = tk.Radiobutton(opt_rotate, text="시계방향 90° ↻", variable=rotate_angle_var, value=90, bg="#F4F6F9", font=("Malgun Gothic", 9))
    rb_rot1.pack(side=tk.LEFT, padx=4)
    rb_rot2 = tk.Radiobutton(opt_rotate, text="180° 반전 ⟲", variable=rotate_angle_var, value=180, bg="#F4F6F9", font=("Malgun Gothic", 9))
    rb_rot2.pack(side=tk.LEFT, padx=4)
    rb_rot3 = tk.Radiobutton(opt_rotate, text="반시계방향 90° ↺", variable=rotate_angle_var, value=270, bg="#F4F6F9", font=("Malgun Gothic", 9))
    rb_rot3.pack(side=tk.LEFT, padx=4)

    # 3) PDF 대형 드롭존
    drop_frame_pdf = tk.Frame(tab_pdf, bg="#EBF3FB", bd=2, relief=tk.GROOVE, padx=20, pady=16)
    drop_frame_pdf.pack(fill=tk.X, pady=(4, 6))

    lbl_drop_icon2 = tk.Label(drop_frame_pdf, text="📷", font=("Segoe UI Emoji", 22), bg="#EBF3FB")
    lbl_drop_icon2.pack()

    lbl_drop_title2 = tk.Label(
        drop_frame_pdf,
        text="여기에 파일을 끌어다 놓으세요 (Drag & Drop)",
        font=("Malgun Gothic", 11, "bold"),
        fg="#004797",
        bg="#EBF3FB"
    )
    lbl_drop_title2.pack(pady=(2, 2))

    lbl_drop_desc2 = tk.Label(
        drop_frame_pdf,
        text="선택한 모드에 맞춰 즉시 처리됩니다.",
        font=("Malgun Gothic", 9),
        fg="#555",
        bg="#EBF3FB"
    )
    lbl_drop_desc2.pack()

    # 4) PDF 실제 처리 디스패처
    def handle_pdf_items(item_list):
        mode = pdf_mode.get()
        paths = []
        for it in item_list:
            if isinstance(it, bytes):
                try:
                    it = it.decode('utf-8')
                except Exception:
                    it = it.decode('mbcs', errors='ignore')
            if os.path.exists(it):
                if os.path.isfile(it):
                    paths.append(it)
                elif os.path.isdir(it):
                    for r, _, fnames in os.walk(it):
                        for fn in fnames:
                            paths.append(os.path.join(r, fn))

        if not paths:
            log("⚠️ 처리할 유효한 파일이 없습니다.")
            return

        # 모드별 처리
        if mode == "conv":
            img_files = [p for p in paths if os.path.splitext(p)[1].lower() in IMAGE_EXTS]
            doc_files = [p for p in paths if os.path.splitext(p)[1].lower() in DOC_EXTS]

            if not img_files and not doc_files:
                log("⚠️ 지원되는 사진(JPG, PNG 등) 또는 문서(Word, Excel, HWP 등) 파일이 없습니다.")
                messagebox.showwarning("파일 확인", "선택하신 파일 중 지원되는 사진이나 문서 파일이 없습니다.")
                return

            if img_files:
                merge_flag = img_merge_var.get()
                ok, out_pdfs, msg = convert_images_to_pdf(img_files, merge_into_single=merge_flag)
                if ok:
                    log(f"📷 {msg}")
                    for op in out_pdfs:
                        log(f"   -> 저장: {os.path.basename(op)}")
                else:
                    log(f"❌ {msg}")

            if doc_files:
                current_fit = excel_fit_var.get()
                fit_label = "너비 1페이지 맞춤" if current_fit == "width" else ("시트 전체 1장 맞춤" if current_fit == "sheet" else "원본 설정 유지")
                log(f"📄 총 {len(doc_files)}개의 문서를 PDF로 변환 시작 (엑셀 맞춤: {fit_label})...")
                succ_cnt = 0
                for df in doc_files:
                    ok, out_pdf, msg = convert_doc_to_pdf(df, excel_fit_mode=current_fit)
                    if ok:
                        succ_cnt += 1
                        log(f"   • {os.path.basename(df)} -> {os.path.basename(out_pdf)}")
                    else:
                        log(f"   ❌ {os.path.basename(df)}: {msg}")
                log(f"📄 문서 변환 완료: 총 {len(doc_files)}개 중 {succ_cnt}개 성공")

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("변환 완료", "선택하신 사진 및 문서의 PDF 변환이 완료되었습니다!")

        elif mode == "merge":
            pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
            if len(pdf_files) < 2:
                log("⚠️ 합치려면 최소 2개 이상의 PDF 파일이 필요합니다.")
                messagebox.showwarning("PDF 합치기", "합치려면 최소 2개 이상의 PDF 파일을 선택해 주세요.")
                return

            sort_flag = pdf_sort_var.get()
            ok, out_pdf, msg = merge_pdfs(pdf_files, sort_by_name=sort_flag)
            if ok:
                log(f"📑 {msg}")
                log(f"   -> 저장: {out_pdf}")
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("병합 완료", f"총 {len(pdf_files)}개 PDF가 하나로 합쳐졌습니다!\n\n저장: {os.path.basename(out_pdf)}")
            else:
                log(f"❌ {msg}")
                messagebox.showerror("오류", msg)

        elif mode == "split":
            pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
            if not pdf_files:
                log("⚠️ 분할할 PDF 파일이 없습니다.")
                return

            range_text = split_range_var.get().strip()
            total_created = 0
            for pf in pdf_files:
                ok, created, msg = split_or_extract_pdf(pf, range_text)
                if ok:
                    total_created += len(created)
                    log(f"✂️ [{os.path.basename(pf)}] {msg}")
                    for cp in created:
                        log(f"   -> {os.path.basename(cp)}")
                else:
                    log(f"❌ [{os.path.basename(pf)}] {msg}")

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("분할/추출 완료", f"총 {total_created}개의 PDF 파일이 생성되었습니다!")

        elif mode == "compress":
            pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
            if not pdf_files:
                log("⚠️ 압축할 PDF 파일이 없습니다.")
                return

            dpi_val = 100 if compress_level_var.get() == "high" else 150
            quality_val = 70 if compress_level_var.get() == "high" else 78

            log(f"📉 PDF 용량 다이어트 진행 중 ({len(pdf_files)}개 파일)...")
            succ_cnt = 0
            for pf in pdf_files:
                ok, out_p, msg = compress_pdf(pf, dpi=dpi_val, quality=quality_val)
                if ok:
                    succ_cnt += 1
                    log(f"   • [{os.path.basename(pf)}] {msg}")
                else:
                    log(f"   ❌ [{os.path.basename(pf)}] {msg}")

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("압축 완료", f"총 {succ_cnt}개 PDF의 용량 압축이 완료되었습니다!")

        elif mode == "rotate":
            pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
            if not pdf_files:
                log("⚠️ 회전할 PDF 파일이 없습니다.")
                return

            ang = rotate_angle_var.get()
            log(f"🔄 PDF {ang}도 회전 진행 중 ({len(pdf_files)}개 파일)...")
            succ_cnt = 0
            for pf in pdf_files:
                ok, out_p, msg = rotate_pdf(pf, angle=ang)
                if ok:
                    succ_cnt += 1
                    log(f"   • [{os.path.basename(pf)}] {msg}")
                else:
                    log(f"   ❌ [{os.path.basename(pf)}] {msg}")

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("회전 완료", f"총 {succ_cnt}개 PDF가 {ang}도 회전되어 저장되었습니다!")

    # 5) PDF 수동 파일 선택 버튼
    def on_pick_pdf_files():
        mode = pdf_mode.get()
        if mode == "conv":
            flt = [("사진 및 문서", "*.jpg;*.jpeg;*.png;*.bmp;*.webp;*.docx;*.xlsx;*.pptx;*.hwp;*.hwpx;*.txt;*.csv"), ("모든 파일", "*.*")]
        else:
            flt = [("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]

        files = filedialog.askopenfilenames(title="파일을 선택하세요", filetypes=flt)
        if files:
            handle_pdf_items(files)

    btn_pick_pdf = tk.Button(
        tab_pdf,
        text="📄 사진 및 서류 파일 직접 선택",
        font=("Malgun Gothic", 9, "bold"),
        bg="#004797",
        fg="white",
        activebackground="#003366",
        activeforeground="white",
        padx=10,
        pady=6,
        relief=tk.FLAT,
        cursor="hand2",
        command=on_pick_pdf_files
    )
    btn_pick_pdf.pack(fill=tk.X, pady=(2, 4))

    # 초기 모드 UI 반영
    update_pdf_ui()

    # ---------------------------------------------------------------------
    # 드래그 앤 드롭 통합 핸들러 (현재 선택된 탭에 따라 동작)
    # ---------------------------------------------------------------------
    def on_drop_files(dropped_items):
        clean_items = []
        for item in dropped_items:
            if isinstance(item, bytes):
                try:
                    item = item.decode('utf-8')
                except Exception:
                    item = item.decode('mbcs', errors='ignore')
            clean_items.append(item)

        selected_tab = notebook.select()
        if notebook.nametowidget(selected_tab) == tab_pdf:
            handle_pdf_items(clean_items)
        else:
            handle_fixer_items(clean_items, "드래그앤드롭")

    if windnd:
        windnd.hook_dropfiles(root, func=on_drop_files, force_unicode=True)

    # ---------------------------------------------------------------------
    # 3. 하단 공통 처리 결과 로그 창
    # ---------------------------------------------------------------------
    log_frame = tk.LabelFrame(
        root,
        text=" 처리 결과 로그 ",
        font=("Malgun Gothic", 9, "bold"),
        bg="#F4F6F9",
        fg="#333",
        padx=10,
        pady=6
    )
    log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(2, 10))

    txt_log = tk.Text(
        log_frame,
        font=("Consolas", 9),
        bg="#FFFFFF",
        fg="#212529",
        wrap=tk.WORD,
        relief=tk.SOLID,
        borderwidth=1
    )
    scrollbar = tk.Scrollbar(log_frame, command=txt_log.yview)
    txt_log.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 로그 제어 버튼
    log_ctrl_frame = tk.Frame(log_frame, bg="#F4F6F9")
    log_ctrl_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

    def on_clear_log():
        txt_log.config(state=tk.NORMAL)
        txt_log.delete("1.0", tk.END)
        txt_log.config(state=tk.DISABLED)

    btn_clear_log = tk.Button(
        log_ctrl_frame,
        text="🗑️ 로그 지우기",
        font=("Malgun Gothic", 8),
        bg="#E9ECEF",
        relief=tk.FLAT,
        cursor="hand2",
        command=on_clear_log
    )
    btn_clear_log.pack(side=tk.RIGHT)

    log("📌 구매부서 유틸리티 (PDF 종합 도구 & 파일명 복구) 준비 완료.")
    log("💡 [탭 1: PDF 종합 도구] 사진/문서 PDF 변환(개별 변환 기본), 합치기, 분할, 압축, 회전을 지원합니다.")
    log("💡 [탭 2: 파일명 복구] 깨진 한글 파일명을 윈도우 완성형으로 1초 만에 복구합니다.\n" + "-" * 70)

    root.mainloop()


def main():
    if len(sys.argv) > 1:
        arg1 = sys.argv[1].lower()
        if arg1 == "--register":
            ok, msg = register_context_menu()
            print(msg)
            ctypes.windll.user32.MessageBoxW(0, msg, APP_TITLE, 0x40 if ok else 0x10)
            return
        elif arg1 == "--unregister":
            ok, msg = unregister_context_menu()
            print(msg)
            ctypes.windll.user32.MessageBoxW(0, msg, APP_TITLE, 0x40 if ok else 0x10)
            return
        elif arg1 == "--shortcut":
            ok, msg = create_desktop_shortcut()
            print(msg)
            ctypes.windll.user32.MessageBoxW(0, msg, APP_TITLE, 0x40 if ok else 0x10)
            return

        targets = sys.argv[1:]
        all_results = []

        for target in targets:
            if os.path.exists(target):
                res = process_path(target)
                all_results.extend(res)

        count = len(all_results)
        if count > 0:
            msg = f"✅ 한글 파일명 복구 완료!\n\n총 {count}개 항목이 정상 복구되었습니다.\n(영문, 숫자, 특수문자 등은 100% 보존됨)"
        else:
            msg = "선택하신 파일/폴더의 한글 이름이\n이미 모두 정상 완성형 상태입니다."

        show_auto_alert(APP_TITLE, msg, auto_close_sec=3)
    else:
        launch_main_gui()


if __name__ == "__main__":
    main()
