<#
naver-works-mcp 설치 자동화 스크립트 (Windows PowerShell)

이 스크립트가 있는 폴더를 기준으로:
  1) 가상환경(.venv) 생성
  2) pip 업그레이드
  3) requirements.txt 설치
  4) .env 파일 준비 (없으면 .env.example을 복사 후 notepad로 오픈)
  5) 설치 확인 + Claude 데스크톱 config에 등록할 JSON 조각 출력

사용법 (PowerShell에서, 이 스크립트가 있는 폴더로 이동한 뒤):
  .\setup.ps1
#>

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

Write-Host "[1/5] 가상환경 생성 중..." -ForegroundColor Cyan
if (-not (Test-Path "$root\.venv")) {
    py -m venv "$root\.venv"
} else {
    Write-Host "  이미 .venv가 있어 건너뜁니다." -ForegroundColor DarkGray
}

Write-Host "[2/5] pip 업그레이드 중..." -ForegroundColor Cyan
& "$root\.venv\Scripts\python.exe" -m pip install --upgrade pip | Out-Null

Write-Host "[3/5] 패키지 설치 중 (requirements.txt)..." -ForegroundColor Cyan
& "$root\.venv\Scripts\python.exe" -m pip install -r "$root\requirements.txt"

Write-Host "[4/5] .env 파일 준비 중..." -ForegroundColor Cyan
if (-not (Test-Path "$root\.env")) {
    Copy-Item "$root\.env.example" "$root\.env"
    Write-Host "  .env를 새로 만들었습니다. notepad에서 NAVER_EMAIL / NAVER_APP_PASSWORD를 입력하고 저장 후 닫아주세요." -ForegroundColor Yellow
    Write-Host "  (외부 앱 비밀번호: 네이버웍스 웹 -> 프로필 -> 개인 설정 -> 보안 -> 외부 앱 비밀번호)" -ForegroundColor Yellow
    notepad "$root\.env"
} else {
    Write-Host "  .env가 이미 있어 건너뜁니다." -ForegroundColor DarkGray
}

Write-Host "[5/5] 설치 확인 중..." -ForegroundColor Cyan
$check = & "$root\.venv\Scripts\python.exe" -c "import mcp, caldav, icalendar, pypdf, docx, openpyxl, dotenv; print('OK')" 2>&1
if ($check -match "OK") {
    Write-Host "  필수 패키지 임포트 확인 완료." -ForegroundColor Green
} else {
    Write-Host "  패키지 임포트 확인 실패. 아래 오류를 확인하세요:" -ForegroundColor Red
    Write-Host $check
    exit 1
}

Write-Host ""
Write-Host "설치가 끝났습니다. Claude 데스크톱 설정 파일의 mcpServers에 아래 항목을 추가하세요:" -ForegroundColor Green
Write-Host ""
$jsonPath = $root.Replace('\', '\\')
Write-Host @"
"naver-works": {
  "command": "$jsonPath\\.venv\\Scripts\\python.exe",
  "args": ["$jsonPath\\server.py"]
}
"@
Write-Host ""
Write-Host "설정 파일 저장 후 'taskkill /IM Claude.exe /F'로 앱을 완전히 종료하고 다시 실행하세요." -ForegroundColor Cyan
