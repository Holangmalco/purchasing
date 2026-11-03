# -*- coding: utf-8 -*-
"""
이폼싸인(eformsign) 워드 모듈 자동 실행기
워드 문서(.docx)를 열고 eformsign 탭의 [실행] 또는 [업로드] 버튼을 자동으로 클릭하여
이폼싸인 전자문서 작성/전송 창을 즉시 호출합니다.
"""

import os
import sys
import subprocess
import tempfile

def trigger_eformsign_in_word(docx_path: str, button_name: str = "실행") -> bool:
    """
    지정된 워드 문서(.docx)를 열고, eformsign 리본 탭의 지정된 버튼(기본값: '실행')을 자동 클릭합니다.
    
    :param docx_path: 열고자 하는 .docx 파일의 절대 경로
    :param button_name: 클릭할 버튼 이름 ('실행' 또는 '업로드')
    :return: 성공 여부 (bool)
    """
    abs_path = os.path.abspath(docx_path)
    if not os.path.exists(abs_path):
        print(f"[!] 파일을 찾을 수 없습니다: {abs_path}")
        return False

    ps_script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$targetPath = "{abs_path}"
$targetButton = "{button_name}"

Write-Host "[*] Word 실행 및 파일 로딩 중: $targetPath"
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Open($targetPath)

try {{
    # Word 창 로딩 대기
    Start-Sleep -Seconds 2
    $processes = Get-Process -Name WINWORD
    if ($processes.Count -eq 0) {{
        Write-Host "[!] WINWORD 프로세스를 찾을 수 없습니다."
        exit 1
    }}
    
    $hwnd = $processes[0].MainWindowHandle
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    if ($root -eq $null) {{
        Write-Host "[!] Word 창 AutomationElement를 가져올 수 없습니다."
        exit 1
    }}
    
    # 1. eformsign 탭 찾기
    Write-Host "[*] eformsign 탭 검색 중..."
    $nameProp = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, "eformsign"
    )
    $eformTab = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $nameProp)
    
    if ($eformTab -eq $null) {{
        Write-Host "[!] Word 상단에 'eformsign' 탭이 없습니다."
        exit 1
    }}
    
    # 2. 탭 활성화 (선택)
    $selectPattern = $eformTab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    if ($selectPattern -ne $null) {{
        $selectPattern.Select()
    }} else {{
        $invokePattern = $eformTab.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        if ($invokePattern -ne $null) {{
            $invokePattern.Invoke()
        }}
    }}
    Write-Host "[*] eformsign 탭 활성화 완료"
    Start-Sleep -Milliseconds 800
    
    # 3. 목표 버튼(실행/업로드) 찾기 및 클릭
    Write-Host "[*] '$targetButton' 버튼 탐색 및 클릭 시도..."
    $btnProp = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button
    )
    $buttons = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnProp)
    $foundBtn = $null
    
    foreach ($b in $buttons) {{
        $n = $b.Current.Name.Trim()
        if ($n -eq $targetButton -or $n -like "*$targetButton*") {{
            $foundBtn = $b
            break
        }}
    }}
    
    if ($foundBtn -ne $null) {{
        Write-Host "[*] 버튼 '$($foundBtn.Current.Name)' 발견! 클릭을 수행합니다."
        $bInvoke = $foundBtn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        if ($bInvoke -ne $null) {{
            $bInvoke.Invoke()
            Write-Host "[+] '$targetButton' 버튼이 성공적으로 자동 클릭되었습니다!"
            exit 0
        }} else {{
            Write-Host "[!] 버튼 InvokePattern을 지원하지 않습니다."
            exit 1
        }}
    }} else {{
        Write-Host "[!] '$targetButton' 버튼을 찾지 못했습니다."
        exit 1
    }}
}} catch {{
    Write-Host "[!] 에러 발생: $_"
    exit 1
}}
"""

    with tempfile.NamedTemporaryFile("w", encoding="utf-8-sig", suffix=".ps1", delete=False) as f:
        f.write(ps_script)
        temp_ps1 = f.name

    try:
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", temp_ps1
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(res.stdout)
        if res.stderr:
            print("STDERR:", res.stderr)
        return res.returncode == 0
    finally:
        if os.path.exists(temp_ps1):
            try:
                os.remove(temp_ps1)
            except Exception:
                pass

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
        btn = sys.argv[2] if len(sys.argv) > 2 else "실행"
        trigger_eformsign_in_word(target_file, btn)
    else:
        test_path = os.path.join(os.path.dirname(__file__), "..", "template_under10mil.docx")
        print(f"[*] 테스트 실행: {test_path}")
        trigger_eformsign_in_word(test_path, "실행")
