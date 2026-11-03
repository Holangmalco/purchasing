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
import customtkinter as ctk
from PIL import Image, ImageOps, ImageDraw
import pypdfium2 as pdfium

try:
    import windnd
except ImportError:
    windnd = None

APP_TITLE = "구매부서 유틸리티"
REG_SHELL_NAME = "PurchasingDeptUtility"
MENU_LABEL = "[구매부서 유틸리티] 파일명 복구"
MENU_BG_LABEL = "[구매부서 유틸리티] 폴더 파일명 복구"

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
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    class ModernPurchasingUtility(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title(APP_TITLE)
            self.geometry("900x700")
            self.minsize(820, 620)
            self.configure(fg_color="#F8F9FA")

            self.current_mode = "conv"

            # 작업 옵션 변수
            self.img_merge_var = ctk.BooleanVar(value=False)
            self.excel_fit_var = ctk.StringVar(value="너비 1페이지")
            self.pdf_sort_var = ctk.BooleanVar(value=True)
            self.split_range_var = ctk.StringVar(value="")
            self.compress_level_var = ctk.StringVar(value="보통 (150 DPI)")
            self.rotate_angle_var = ctk.StringVar(value="90° ↻")

            # 2열 그리드 레이아웃
            self.grid_columnconfigure(0, weight=0, minsize=220)
            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self._build_sidebar()
            self._build_content_area()

            # 창 전체 드래그 앤 드롭 연결
            if windnd:
                try:
                    windnd.hook_dropfiles(self, func=self._on_drop_files, force_unicode=True)
                except Exception as e:
                    self.log(f"⚠️ 드래그 앤 드롭 오류: {e}")

            self.log("구매부서 유틸리티 준비 완료")

        def _build_sidebar(self):
            self.sidebar = ctk.CTkFrame(
                self,
                fg_color="#FFFFFF",
                corner_radius=0,
                border_width=1,
                border_color="#E9ECEF"
            )
            self.sidebar.grid(row=0, column=0, sticky="nsew")

            # 앱 브랜드 헤더 (간결하게 "구매부서 유틸리티"만)
            brand_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
            brand_frame.pack(fill="x", padx=18, pady=(20, 14))

            lbl_logo = ctk.CTkLabel(
                brand_frame,
                text="구매부서 유틸리티",
                font=("Segoe UI", 16, "bold"),
                text_color="#191F28"
            )
            lbl_logo.pack(anchor="w")

            # 섹션 1: PDF 도구
            lbl_sec1 = ctk.CTkLabel(
                self.sidebar,
                text="PDF 도구",
                font=("Segoe UI", 11, "bold"),
                text_color="#8B95A1"
            )
            lbl_sec1.pack(anchor="w", padx=18, pady=(8, 4))

            self.nav_buttons = {}
            modes = [
                ("conv", "📷  사진·문서 변환"),
                ("merge", "📑  PDF 병합"),
                ("split", "✂️  페이지 분할·추출"),
                ("compress", "📉  용량 압축"),
                ("rotate", "🔄  방향 회전"),
            ]

            for key, label in modes:
                btn = ctk.CTkButton(
                    self.sidebar,
                    text=label,
                    font=("Segoe UI", 12),
                    height=36,
                    corner_radius=8,
                    anchor="w",
                    fg_color="transparent",
                    text_color="#4E5968",
                    hover_color="#F2F4F6",
                    command=lambda k=key: self.set_mode(k)
                )
                btn.pack(fill="x", padx=10, pady=2)
                self.nav_buttons[key] = btn

            # 섹션 2: 파일명 복구
            lbl_sec2 = ctk.CTkLabel(
                self.sidebar,
                text="파일명 복구",
                font=("Segoe UI", 11, "bold"),
                text_color="#8B95A1"
            )
            lbl_sec2.pack(anchor="w", padx=18, pady=(12, 4))

            btn_fixer = ctk.CTkButton(
                self.sidebar,
                text="🛠️  자소 분리 복구",
                font=("Segoe UI", 12),
                height=36,
                corner_radius=8,
                anchor="w",
                fg_color="transparent",
                text_color="#4E5968",
                hover_color="#F2F4F6",
                command=lambda: self.set_mode("fixer")
            )
            btn_fixer.pack(fill="x", padx=10, pady=2)
            self.nav_buttons["fixer"] = btn_fixer

            # 하단 시스템 설정 카드
            sys_frame = ctk.CTkFrame(self.sidebar, fg_color="#F8F9FA", corner_radius=10)
            sys_frame.pack(side="bottom", fill="x", padx=10, pady=14)

            lbl_sys = ctk.CTkLabel(
                sys_frame,
                text="시스템 설정",
                font=("Segoe UI", 11, "bold"),
                text_color="#4E5968"
            )
            lbl_sys.pack(anchor="w", padx=10, pady=(8, 4))

            btn_reg = ctk.CTkButton(
                sys_frame,
                text="우클릭 메뉴 등록",
                font=("Segoe UI", 11),
                height=28,
                corner_radius=6,
                fg_color="#FFFFFF",
                text_color="#191F28",
                hover_color="#E9ECEF",
                border_width=1,
                border_color="#E5E8EB",
                command=self._on_register_menu
            )
            btn_reg.pack(fill="x", padx=8, pady=2)

            btn_unreg = ctk.CTkButton(
                sys_frame,
                text="우클릭 메뉴 삭제",
                font=("Segoe UI", 11),
                height=28,
                corner_radius=6,
                fg_color="#FFFFFF",
                text_color="#6B7684",
                hover_color="#E9ECEF",
                border_width=1,
                border_color="#E5E8EB",
                command=self._on_unregister_menu
            )
            btn_unreg.pack(fill="x", padx=8, pady=2)

            btn_sc = ctk.CTkButton(
                sys_frame,
                text="바탕화면 바로가기",
                font=("Segoe UI", 11),
                height=28,
                corner_radius=6,
                fg_color="#FFFFFF",
                text_color="#191F28",
                hover_color="#E9ECEF",
                border_width=1,
                border_color="#E5E8EB",
                command=self._on_create_shortcut
            )
            btn_sc.pack(fill="x", padx=8, pady=(2, 8))

        def _build_content_area(self):
            self.main_content = ctk.CTkFrame(self, fg_color="#F8F9FA", corner_radius=0)
            self.main_content.grid(row=0, column=1, sticky="nsew", padx=20, pady=16)
            self.main_content.grid_columnconfigure(0, weight=1)
            self.main_content.grid_rowconfigure(3, weight=1)

            # 1. 헤더 (모드 제목 및 요약)
            self.header_frame = ctk.CTkFrame(self.main_content, fg_color="transparent")
            self.header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

            self.lbl_mode_title = ctk.CTkLabel(
                self.header_frame,
                text="사진·문서 변환",
                font=("Segoe UI", 18, "bold"),
                text_color="#191F28"
            )
            self.lbl_mode_title.pack(anchor="w")

            self.lbl_mode_desc = ctk.CTkLabel(
                self.header_frame,
                text="이미지(JPG, PNG) 및 문서(Word, Excel, HWP) ➔ PDF 변환",
                font=("Segoe UI", 12),
                text_color="#6B7684"
            )
            self.lbl_mode_desc.pack(anchor="w", pady=(2, 0))

            # 2. 드롭존 카드
            self.drop_card = ctk.CTkFrame(
                self.main_content,
                fg_color="#FFFFFF",
                corner_radius=14,
                border_width=1.5,
                border_color="#E5E8EB"
            )
            self.drop_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))

            drop_inner = ctk.CTkFrame(self.drop_card, fg_color="transparent")
            drop_inner.pack(fill="both", expand=True, padx=20, pady=20)

            self.lbl_drop_icon = ctk.CTkLabel(drop_inner, text="📷", font=("Segoe UI Emoji", 34))
            self.lbl_drop_icon.pack(pady=(0, 2))

            self.lbl_drop_title = ctk.CTkLabel(
                drop_inner,
                text="변환할 파일을 여기에 놓으세요",
                font=("Segoe UI", 14, "bold"),
                text_color="#191F28"
            )
            self.lbl_drop_title.pack()

            self.lbl_drop_sub = ctk.CTkLabel(
                drop_inner,
                text="드래그 앤 드롭 시 즉시 실행",
                font=("Segoe UI", 11),
                text_color="#8B95A1"
            )
            self.lbl_drop_sub.pack(pady=(2, 14))

            # 작업 버튼 바
            self.action_btn_frame = ctk.CTkFrame(drop_inner, fg_color="transparent")
            self.action_btn_frame.pack()

            self.btn_pick_primary = ctk.CTkButton(
                self.action_btn_frame,
                text="파일 선택",
                font=("Segoe UI", 12, "bold"),
                height=36,
                corner_radius=18,
                fg_color="#3182F6",
                hover_color="#1B64DA",
                command=self._on_pick_files
            )
            self.btn_pick_primary.pack(side="left", padx=4)

            self.btn_pick_folder = ctk.CTkButton(
                self.action_btn_frame,
                text="폴더 선택",
                font=("Segoe UI", 12, "bold"),
                height=36,
                corner_radius=18,
                fg_color="#E8F3FF",
                text_color="#1B64DA",
                hover_color="#D4E8FC",
                command=self._on_pick_folder
            )

            # 3. 옵션 카드
            self.options_card = ctk.CTkFrame(
                self.main_content,
                fg_color="#FFFFFF",
                corner_radius=12,
                border_width=1,
                border_color="#E5E8EB"
            )
            self.options_card.grid(row=2, column=0, sticky="ew", pady=(0, 12))

            self._build_options()

            # 4. 하단 콘솔 로그 카드
            self.log_card = ctk.CTkFrame(
                self.main_content,
                fg_color="#FFFFFF",
                corner_radius=12,
                border_width=1,
                border_color="#E5E8EB"
            )
            self.log_card.grid(row=3, column=0, sticky="nsew")

            log_header = ctk.CTkFrame(self.log_card, fg_color="transparent")
            log_header.pack(fill="x", padx=14, pady=(8, 4))

            lbl_log_title = ctk.CTkLabel(
                log_header,
                text="처리 내역",
                font=("Segoe UI", 11, "bold"),
                text_color="#333D4B"
            )
            lbl_log_title.pack(side="left")

            btn_clear = ctk.CTkButton(
                log_header,
                text="비우기",
                width=50,
                height=22,
                corner_radius=10,
                font=("Segoe UI", 10),
                fg_color="#F2F4F6",
                text_color="#6B7684",
                hover_color="#E5E8EB",
                command=self.clear_log
            )
            btn_clear.pack(side="right")

            self.txt_log = ctk.CTkTextbox(
                self.log_card,
                font=("Consolas", 10),
                fg_color="#F8F9FA",
                text_color="#333D4B",
                corner_radius=8,
                border_width=1,
                border_color="#E9ECEF"
            )
            self.txt_log.pack(fill="both", expand=True, padx=14, pady=(0, 10))

            # 초기 모드 적용
            self.set_mode("conv")

        def _build_options(self):
            # (1) 사진/문서 변환 옵션
            self.opt_conv = ctk.CTkFrame(self.options_card, fg_color="transparent")
            self.sw_img_merge = ctk.CTkSwitch(
                self.opt_conv,
                text="사진 1개 PDF로 병합",
                variable=self.img_merge_var,
                font=("Segoe UI", 11, "bold"),
                text_color="#191F28",
                progress_color="#3182F6"
            )
            self.sw_img_merge.pack(side="left", padx=14, pady=10)

            lbl_excel = ctk.CTkLabel(
                self.opt_conv,
                text="엑셀 맞춤:",
                font=("Segoe UI", 11, "bold"),
                text_color="#4E5968"
            )
            lbl_excel.pack(side="left", padx=(10, 6))

            self.seg_excel = ctk.CTkSegmentedButton(
                self.opt_conv,
                values=["너비 1페이지", "시트 전체 1장", "원본 유지"],
                variable=self.excel_fit_var,
                font=("Segoe UI", 11),
                corner_radius=6,
                selected_color="#3182F6",
                selected_hover_color="#1B64DA"
            )
            self.seg_excel.pack(side="left", padx=(0, 14))

            # (2) 합치기 옵션
            self.opt_merge = ctk.CTkFrame(self.options_card, fg_color="transparent")
            self.sw_sort = ctk.CTkSwitch(
                self.opt_merge,
                text="가나다순 자동 정렬",
                variable=self.pdf_sort_var,
                font=("Segoe UI", 11, "bold"),
                text_color="#191F28",
                progress_color="#3182F6"
            )
            self.sw_sort.pack(side="left", padx=14, pady=10)

            # (3) 분할/추출 옵션
            self.opt_split = ctk.CTkFrame(self.options_card, fg_color="transparent")
            lbl_split = ctk.CTkLabel(
                self.opt_split,
                text="추출 범위:",
                font=("Segoe UI", 11, "bold"),
                text_color="#191F28"
            )
            lbl_split.pack(side="left", padx=(14, 6), pady=10)

            self.ent_split = ctk.CTkEntry(
                self.opt_split,
                textvariable=self.split_range_var,
                placeholder_text="예: 1-3, 5 (미입력 시 낱개 분할)",
                width=260,
                height=32,
                corner_radius=6,
                font=("Segoe UI", 11)
            )
            self.ent_split.pack(side="left", padx=(0, 14), pady=10)

            # (4) 압축 옵션
            self.opt_compress = ctk.CTkFrame(self.options_card, fg_color="transparent")
            lbl_comp = ctk.CTkLabel(
                self.opt_compress,
                text="압축 강도:",
                font=("Segoe UI", 11, "bold"),
                text_color="#191F28"
            )
            lbl_comp.pack(side="left", padx=(14, 6), pady=10)

            self.seg_comp = ctk.CTkSegmentedButton(
                self.opt_compress,
                values=["보통 (150 DPI)", "고압축 (100 DPI)"],
                variable=self.compress_level_var,
                font=("Segoe UI", 11),
                corner_radius=6,
                selected_color="#3182F6"
            )
            self.seg_comp.pack(side="left", padx=(0, 14), pady=10)

            # (5) 회전 옵션
            self.opt_rotate = ctk.CTkFrame(self.options_card, fg_color="transparent")
            lbl_rot = ctk.CTkLabel(
                self.opt_rotate,
                text="회전 각도:",
                font=("Segoe UI", 11, "bold"),
                text_color="#191F28"
            )
            lbl_rot.pack(side="left", padx=(14, 6), pady=10)

            self.seg_rot = ctk.CTkSegmentedButton(
                self.opt_rotate,
                values=["90° ↻", "180°", "270° ↺"],
                variable=self.rotate_angle_var,
                font=("Segoe UI", 11),
                corner_radius=6,
                selected_color="#3182F6"
            )
            self.seg_rot.pack(side="left", padx=(0, 14), pady=10)

            # (6) 한글 복구 옵션
            self.opt_fixer = ctk.CTkFrame(self.options_card, fg_color="transparent")
            lbl_fixer_info = ctk.CTkLabel(
                self.opt_fixer,
                text="맥북(NFD) 자소 분리 및 인코딩 오류 복구",
                font=("Segoe UI", 11),
                text_color="#4E5968"
            )
            lbl_fixer_info.pack(side="left", padx=14, pady=10)

        def set_mode(self, mode):
            self.current_mode = mode

            # 사이드바 활성화 스타일
            for k, btn in self.nav_buttons.items():
                if k == mode:
                    btn.configure(
                        fg_color="#E8F3FF",
                        text_color="#1B64DA",
                        font=("Segoe UI", 12, "bold")
                    )
                else:
                    btn.configure(
                        fg_color="transparent",
                        text_color="#4E5968",
                        font=("Segoe UI", 12)
                    )

            # 기존 옵션 숨기기
            self.opt_conv.pack_forget()
            self.opt_merge.pack_forget()
            self.opt_split.pack_forget()
            self.opt_compress.pack_forget()
            self.opt_rotate.pack_forget()
            self.opt_fixer.pack_forget()
            self.btn_pick_folder.pack_forget()

            if mode == "conv":
                self.lbl_mode_title.configure(text="사진·문서 변환")
                self.lbl_mode_desc.configure(text="이미지 및 문서 파일 ➔ PDF 일괄 변환")
                self.lbl_drop_icon.configure(text="📷")
                self.lbl_drop_title.configure(text="변환할 파일을 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.opt_conv.pack(fill="x")
            elif mode == "merge":
                self.lbl_mode_title.configure(text="PDF 병합")
                self.lbl_mode_desc.configure(text="여러 PDF 파일을 하나로 병합")
                self.lbl_drop_icon.configure(text="📑")
                self.lbl_drop_title.configure(text="병합할 PDF 파일을 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.opt_merge.pack(fill="x")
            elif mode == "split":
                self.lbl_mode_title.configure(text="페이지 분할·추출")
                self.lbl_mode_desc.configure(text="지정 페이지 추출 또는 1장씩 분할")
                self.lbl_drop_icon.configure(text="✂️")
                self.lbl_drop_title.configure(text="분할할 PDF 파일을 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.opt_split.pack(fill="x")
            elif mode == "compress":
                self.lbl_mode_title.configure(text="용량 압축")
                self.lbl_mode_desc.configure(text="문서 용량 최적화 및 압축")
                self.lbl_drop_icon.configure(text="📉")
                self.lbl_drop_title.configure(text="압축할 PDF 파일을 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.opt_compress.pack(fill="x")
            elif mode == "rotate":
                self.lbl_mode_title.configure(text="방향 회전")
                self.lbl_mode_desc.configure(text="문서 방향 회전 (90° / 180° / 270°)")
                self.lbl_drop_icon.configure(text="🔄")
                self.lbl_drop_title.configure(text="회전할 PDF 파일을 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.opt_rotate.pack(fill="x")
            elif mode == "fixer":
                self.lbl_mode_title.configure(text="파일명 복구")
                self.lbl_mode_desc.configure(text="맥북(NFD) 자소 분리 및 인코딩 깨짐 복구")
                self.lbl_drop_icon.configure(text="🛠️")
                self.lbl_drop_title.configure(text="복구할 파일 또는 폴더를 여기에 놓으세요")
                self.btn_pick_primary.configure(text="파일 선택")
                self.btn_pick_folder.pack(side="left", padx=4)
                self.opt_fixer.pack(fill="x")

        def log(self, message):
            self.txt_log.insert("end", str(message) + "\n")
            self.txt_log.see("end")

        def clear_log(self):
            self.txt_log.delete("1.0", "end")

        def _on_drop_files(self, dropped_items):
            clean_items = []
            for item in dropped_items:
                if isinstance(item, bytes):
                    try: item = item.decode('utf-8')
                    except Exception: item = item.decode('mbcs', errors='ignore')
                clean_items.append(item)

            if self.current_mode == "fixer":
                self.handle_fixer_items(clean_items, "드래그앤드롭")
            else:
                self.handle_pdf_items(clean_items)

        def _on_pick_files(self):
            mode = self.current_mode
            if mode == "fixer":
                flt = [("모든 파일", "*.*")]
                files = filedialog.askopenfilenames(title="파일 선택", filetypes=flt)
                if files: self.handle_fixer_items(files, "파일 선택")
            elif mode == "conv":
                flt = [("사진 및 문서", "*.jpg;*.jpeg;*.png;*.bmp;*.webp;*.docx;*.xlsx;*.pptx;*.hwp;*.hwpx;*.txt;*.csv"), ("모든 파일", "*.*")]
                files = filedialog.askopenfilenames(title="파일 선택", filetypes=flt)
                if files: self.handle_pdf_items(files)
            else:
                flt = [("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
                files = filedialog.askopenfilenames(title="PDF 선택", filetypes=flt)
                if files: self.handle_pdf_items(files)

        def _on_pick_folder(self):
            folder = filedialog.askdirectory(title="폴더 선택")
            if folder:
                self.handle_fixer_items([folder], "폴더 선택")

        def handle_fixer_items(self, item_list, source_label="선택"):
            all_res = []
            for item in item_list:
                if os.path.exists(item):
                    res = process_path(item)
                    all_res.extend(res)

            if all_res:
                self.log(f"✅ [복구 완료] {len(all_res)}개 항목:")
                for old_n, new_n in all_res:
                    self.log(f"   • {old_n} -> {new_n}")
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("복구 완료", f"총 {len(all_res)}개 파일/폴더명이 복구되었습니다.")
            else:
                self.log(f"✨ [확인] 선택한 항목이 이미 정상이거나 복구 대상이 없습니다.")
                winsound.MessageBeep(winsound.MB_ICONINFORMATION)
                messagebox.showinfo("확인", "깨진 파일명이 없거나 이미 정상입니다.")

        def handle_pdf_items(self, item_list):
            mode = self.current_mode
            paths = []
            for it in item_list:
                if os.path.exists(it):
                    if os.path.isfile(it):
                        paths.append(it)
                    elif os.path.isdir(it):
                        for r, _, fnames in os.walk(it):
                            for fn in fnames:
                                paths.append(os.path.join(r, fn))

            if not paths:
                self.log("⚠️ 처리할 유효한 파일이 없습니다.")
                return

            if mode == "conv":
                img_files = [p for p in paths if os.path.splitext(p)[1].lower() in IMAGE_EXTS]
                doc_files = [p for p in paths if os.path.splitext(p)[1].lower() in DOC_EXTS]

                if not img_files and not doc_files:
                    self.log("⚠️ 변환 가능한 파일(사진, Word, Excel, HWP 등)이 없습니다.")
                    messagebox.showwarning("파일 확인", "변환 가능한 사진이나 문서 파일이 없습니다.")
                    return

                if img_files:
                    merge_flag = self.img_merge_var.get()
                    ok, out_pdfs, msg = convert_images_to_pdf(img_files, merge_into_single=merge_flag)
                    if ok:
                        self.log(f"📷 {msg}")
                        for op in out_pdfs:
                            self.log(f"   -> {os.path.basename(op)}")
                    else:
                        self.log(f"❌ {msg}")

                if doc_files:
                    cur_fit_val = self.excel_fit_var.get()
                    fit_mode = "width" if "너비" in cur_fit_val else ("sheet" if "전체" in cur_fit_val else "none")
                    self.log(f"📄 문서 {len(doc_files)}개 변환 시작...")
                    succ_cnt = 0
                    for df in doc_files:
                        ok, out_pdf, msg = convert_doc_to_pdf(df, excel_fit_mode=fit_mode)
                        if ok:
                            succ_cnt += 1
                            self.log(f"   • {os.path.basename(df)} -> {os.path.basename(out_pdf)}")
                        else:
                            self.log(f"   ❌ {os.path.basename(df)}: {msg}")
                    self.log(f"📄 문서 변환 완료: {len(doc_files)}개 중 {succ_cnt}개 성공")

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("완료", "PDF 변환이 완료되었습니다.")

            elif mode == "merge":
                pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
                if len(pdf_files) < 2:
                    self.log("⚠️ 병합하려면 2개 이상의 PDF 파일이 필요합니다.")
                    messagebox.showwarning("PDF 병합", "2개 이상의 PDF 파일을 선택해 주세요.")
                    return

                sort_flag = self.pdf_sort_var.get()
                ok, out_pdf, msg = merge_pdfs(pdf_files, sort_by_name=sort_flag)
                if ok:
                    self.log(f"📑 {msg}")
                    self.log(f"   -> {os.path.basename(out_pdf)}")
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    messagebox.showinfo("완료", f"PDF {len(pdf_files)}개가 병합되었습니다.\n저장: {os.path.basename(out_pdf)}")
                else:
                    self.log(f"❌ {msg}")
                    messagebox.showerror("오류", msg)

            elif mode == "split":
                pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
                if not pdf_files:
                    self.log("⚠️ 분할할 PDF 파일이 없습니다.")
                    return

                range_text = self.split_range_var.get().strip()
                total_created = 0
                for pf in pdf_files:
                    ok, created, msg = split_or_extract_pdf(pf, range_text)
                    if ok:
                        total_created += len(created)
                        self.log(f"✂️ [{os.path.basename(pf)}] {msg}")
                        for cp in created:
                            self.log(f"   -> {os.path.basename(cp)}")
                    else:
                        self.log(f"❌ [{os.path.basename(pf)}] {msg}")

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("완료", f"총 {total_created}개 PDF 파일이 생성되었습니다.")

            elif mode == "compress":
                pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
                if not pdf_files:
                    self.log("⚠️ 압축할 PDF 파일이 없습니다.")
                    return

                is_high = "고압축" in self.compress_level_var.get()
                dpi_val = 100 if is_high else 150
                quality_val = 70 if is_high else 78

                self.log(f"📉 PDF 압축 진행 중 ({len(pdf_files)}개)...")
                succ_cnt = 0
                for pf in pdf_files:
                    ok, out_p, msg = compress_pdf(pf, dpi=dpi_val, quality=quality_val)
                    if ok:
                        succ_cnt += 1
                        self.log(f"   • [{os.path.basename(pf)}] {msg}")
                    else:
                        self.log(f"   ❌ [{os.path.basename(pf)}] {msg}")

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("완료", f"{succ_cnt}개 PDF의 압축이 완료되었습니다.")

            elif mode == "rotate":
                pdf_files = [p for p in paths if p.lower().endswith('.pdf')]
                if not pdf_files:
                    self.log("⚠️ 회전할 PDF 파일이 없습니다.")
                    return

                rot_val = self.rotate_angle_var.get()
                ang = 180 if "180" in rot_val else (270 if "270" in rot_val else 90)

                self.log(f"🔄 PDF {ang}도 회전 진행 중 ({len(pdf_files)}개)...")
                succ_cnt = 0
                for pf in pdf_files:
                    ok, out_p, msg = rotate_pdf(pf, angle=ang)
                    if ok:
                        succ_cnt += 1
                        self.log(f"   • [{os.path.basename(pf)}] {msg}")
                    else:
                        self.log(f"   ❌ [{os.path.basename(pf)}] {msg}")

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("완료", f"{succ_cnt}개 PDF가 회전되었습니다.")

        def _on_register_menu(self):
            ok, msg = register_context_menu()
            if ok:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("우클릭 등록", msg)
                self.log(f"✨ {msg}")
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                messagebox.showerror("등록 실패", msg)
                self.log(f"❌ {msg}")

        def _on_unregister_menu(self):
            ok, msg = unregister_context_menu()
            if ok:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("우클릭 삭제", msg)
                self.log(f"🗑️ {msg}")
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                messagebox.showerror("삭제 실패", msg)
                self.log(f"❌ {msg}")

        def _on_create_shortcut(self):
            ok, msg = create_desktop_shortcut()
            if ok:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("바로가기", msg)
                self.log(f"📌 {msg}")
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
                messagebox.showerror("생성 실패", msg)
                self.log(f"❌ {msg}")

    app = ModernPurchasingUtility()
    app.mainloop()


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
            msg = f"파일명 복구 완료\n\n총 {count}개 항목이 정상 복구되었습니다."
        else:
            msg = "선택하신 항목의 파일명이 이미 정상입니다."

        show_auto_alert(APP_TITLE, msg, auto_close_sec=3)
    else:
        launch_main_gui()


if __name__ == "__main__":
    main()
