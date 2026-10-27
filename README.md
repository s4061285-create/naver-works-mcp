# naver-works-mcp (네이버웍스 메일 + 캘린더 통합 MCP)

> v2.1 · MIT License · 설치자마다 `.env`(자격증명)와 서명 파일을 직접 만들어야 합니다.
>
> 네이버웍스 메일(IMAP/SMTP) + 캘린더(CalDAV)를 외부 앱 비밀번호 하나로 다루는 MCP 서버입니다.
> 초기 구조는 [jaykim81/naver-works-mail-mcp](https://github.com/jaykim81/naver-works-mail-mcp)에서
> 출발했으며, 캘린더·전달·답장 등을 대폭 확장했습니다.

**v2.1 변경사항**: 메일 **답장(reply)** 기능과 **연락처 검색** 기능이 새로 추가됐습니다.
- `reply_email(mail_id, body, folder="inbox", reply_all=False, cc=None, include_attachments=False)` 추가
  — 메일 19개 도구 체제로 확장(총 29개).
- `reply_all=False`(기본값)면 원본 발신자에게만, `reply_all=True`면 원본 To/Cc에 있던
  사람들까지 참조로 포함해 **전체회신**합니다 (내 계정 주소는 자동 제외). `cc`로 참조자를
  추가로 더 넣을 수도 있습니다 (전체회신과 조합 가능).
- 제목은 자동으로 `RE: 원본제목`으로 붙고(이미 `RE:`면 중복 추가 안 함), 본문은
  `forward_email`과 동일한 아웃웍 스타일 `-----Original Message-----` 인용 형식과
  원본 HTML·인라인 이미지 보존 로직을 그대로 공유합니다 (내부적으로 `_build_quoted_mime`
  헬퍼로 통합).
- 원본 메일의 `Message-ID`를 읽어 `In-Reply-To`/`References` 헤더를 채워주므로, 메일
  클라이언트에서 실제 답장 스레드로 묶입니다.
- 답장에는 개인 서명이 붙지 않고, 원본 첨부파일은 기본적으로 다시 첨부하지 않습니다
  (`include_attachments=True`로 켤 수 있음).
- `search_contacts(name, limit=10, scan_per_folder=300)` 추가 — 네이버웍스 정식 조직도
  API 연동(개발자 콘솔 등록 + 관리자 승인 필요) 없이, 받은/보낸메일함 최근 메일들의
  From/To/Cc 헤더를 이름으로 뒤져 이메일 주소를 추정합니다. 결과가 여러 건이면 사용자
  확인 후 사용을 권장합니다.

**v2.0 변경사항**: 메일 **전달(forward)** 기능이 새로 추가됐습니다.
- `forward_email(mail_id, to, folder="inbox", cc=None, comment="", include_attachments=True)` 추가
  — 메일 17개 도구 체제로 확장(총 27개).
- 제목은 자동으로 `FW: 원본제목`으로 붙습니다 (이미 `FW:`로 시작하면 중복 추가 안 함).
- 본문은 실제 메일 솔루션(아웃룩 등)이 만드는 전달 인용 형식을 그대로 재현합니다:
  ```
  -----Original Message-----
  From: "이름"<이메일>
  To: <이메일>;
  Cc: 
  Sent: 2026. 9. 17. (목) 17:18 (GMT+09:00)
  Subject: 원본 제목
  ```
  `From:`/`To:`/`Cc:`/`Sent:`/`Subject:` 레이블만 굵게 표시되고, `Cc:`는 참조자가 없어도
  항상 표시됩니다. 날짜는 `Sent:` 원본 `Date` 헤더를 한국어 아웃룩 스타일로 변환해서 보여줍니다.
- 원본이 HTML 메일이면 plain-text가 아니라 **원본 HTML 본문을 그대로 재사용**해서
  이미지가 포함된 서식이 전달 메일에서도 깨지지 않고 보입니다. 원본에 `cid:`로 참조하는
  인라인 이미지가 있으면 `Content-ID`를 유지한 채 함께 실어 보냅니다(`include_attachments`로
  일반 첨부파일과 함께 껐다 켤 수 있음).
- 전달 메일에는 개인 서명(`signature`)이 붙지 않습니다.
- 전달 인용 헤더/코멘트 영역은 `Sans Serif` 폰트로 표시됩니다(원본 본문 자체의 폰트에는
  영향 없음).

**v1.9 변경사항**: `send_email` 개인 서명(v1.8)의 렌더링 품질 개선.
- 이메일 클라이언트(특히 Outlook 등 Word 렌더링 엔진 계열)가 `<table>`에만 지정된
  `font-family`를 안쪽 `<td>`/`<b>`/`<span>`/`<a>`까지 상속시키지 않아 기본 대체
  폰트로 깨져 보이던 문제를 수정 — 텍스트가 들어가는 요소마다 `font-family`를
  각각 명시하도록 변경.
- 서명 폰트 크기 조정: 이름 13px 굵게, 나머지 항목(직책·부서·연락처 등) 12px.
- 서명 레이아웃 간격 조정: 텍스트 영역과 QR코드 쪽 경계선 사이 간격을 로고↔경계선과
  동일한 14px로 맞춤 (경계선↔QR코드 간격은 기존 그대로 유지).

**v1.8 변경사항**: `send_email`에 개인 서명 옵션 추가.
- `send_email(to, subject, body, cc=None, signature=False)` — `signature=True`로
  호출하면 본문 끝에 개인 서명(`signature.html` — 로고·이름/직책/연락처·QR코드 등)을
  자동으로 붙여서 보낸다.
- 로고·QR코드는 파일 첨부가 아니라 네이버웍스가 호스팅하는 외부 이미지 URL을
  `<img src="...">`로 그대로 참조하는 방식이라 별도 첨부 파일이 필요 없다.
- 일반 텍스트 클라이언트용 대체 본문(plain 파트)에는 이미지 없이 텍스트 서명만 붙는다.
- 신규로 작성하는 메일에는 `signature=True`, 회신/답장/전달에는 `signature=False`
  (기본값)로 호출하는 것을 권장.

**v1.7 변경사항**: 메일 발송 줄바꿈 버그 수정.
- `send_email` 본문을 HTML로 바꾸는 `_build_html_body()`에서 줄바꿈을
  `"\n"` → `"<br>\n"`로 변환하고 있었는데, 본문 스타일에 걸린
  `white-space:pre-wrap` 때문에 뒤에 남은 개행문자(`\n`)까지 실제 줄바꿈으로
  한 번 더 렌더링되어 엔터 한 번이 화면에서는 두 번처럼(줄간격이 두 배로) 보이던
  문제를 수정했습니다.
- `"\n"` → `"<br>\n"` 를 `"\n"` → `"<br>"` 로 변경. 문단 사이에 빈 줄을 하나
  넣으면 이제 실제로 한 줄만 띄워서 보입니다.

**v1.6 변경사항**: 폴더 지정이 전면 확장됐습니다.
- `move_email(mail_id, to_folder, folder)` 추가 — 메일을 **임의의 폴더로 이동**합니다.
  휴지통에 넣은 메일을 받은편지함으로 되돌릴 때 사용합니다.
- `list_emails` / `list_unread_emails` / `search_emails`에 `folder` 인자 추가 — 이제
  휴지통·임시보관함·스팸은 물론 직접 만든 폴더까지 조회·검색할 수 있습니다.
- 폴더 별칭 지원: `inbox` / `sent` / `trash` / `drafts` / `junk` 와 한글 이름
  (`받은편지함`, `보낸메일함`, `휴지통`, `임시보관함`, `스팸`). 그 밖의 값은
  `list_mail_folders()`가 알려주는 IMAP 폴더명으로 보고 그대로 사용합니다.
- 기존에는 `folder` 인자가 `sent`가 아니면 무조건 받은편지함으로 처리됐는데,
  이제 오타는 조용히 무시되지 않고 해당 폴더를 찾지 못했다는 오류로 드러납니다.

**v1.5 변경사항**: 메일 정리용 도구 3종이 추가됐습니다.
- `move_email_to_trash` — 메일을 **휴지통 폴더로 이동**합니다. IMAP `MOVE`를 먼저 시도하고
  서버가 지원하지 않으면 `COPY` + `\Deleted` 플래그 + `EXPUNGE`로 대체합니다. 휴지통
  폴더명은 특수용도 플래그(`\Trash`)와 이름 힌트로 자동 탐색하며, 실패 시 `.env`의
  `TRASH_FOLDER`로 직접 지정할 수 있습니다.
- `delete_local_file` — `save_attachment`로 내려받은 첨부파일을 삭제합니다. 안전장치로
  `ATTACHMENT_DIR` 또는 `.env`의 `DELETE_ALLOWED_DIRS` 아래 파일만 삭제됩니다.
- `list_mail_folders` — 서버의 실제 IMAP 폴더 목록을 조회합니다(휴지통 폴더명 확인용).

또한 `search_emails` / `search_sent_emails`의 **한글 검색 오류**를 고쳤습니다. 기존에는
한글 키워드를 넣으면 `'ascii' codec can't encode characters` 오류가 났는데, CHARSET UTF-8
리터럴 전송 + 실패 시 제목 헤더 직접 대조 폴백으로 동작하도록 바꿨습니다.

**v1.4 변경사항**: `send_email` 본문이 **맑은 고딕 14px** 스타일이 적용된 HTML 메일로
전송되도록 변경했습니다 (일반 텍스트 클라이언트를 위한 대체 텍스트 파트도 함께 포함).
줄바꿈은 그대로 유지되고, 본문에 `<`, `>`, `&` 같은 문자가 있어도 안전하게 이스케이프
처리됩니다.

**v1.3 변경사항**: `send_email`에 `cc`(참조) 파라미터 추가. 메일 목록/상세 조회 결과에도
`cc` 필드가 추가됩니다. 이와 함께 제목/발신자/수신자 헤더 디코딩 버그도 수정됐습니다 —
기존에는 인코딩된 이름이 여러 개일 때(예: "홍길동 <a@x>, 김철수 <b@x>") 첫 번째 이후가
잘려나갔는데, 전체를 이어붙이도록 고쳤습니다.

**v1.2 변경사항**: `attendees` 참석자 초대 기능의 실제 동작이 테스트로 확인됐습니다.
`create_my_event`/`create_team_event`에 `attendees=["이메일"]`을 넘기면 **참석자 캘린더에도
일정이 조용히 등록**되지만, **초대 알림(수락/거절)은 가지 않습니다.** 참석자에게는 별도로
메일/메시지로 알려주는 것을 권장합니다.

Claude 데스크톱에서 네이버웍스 메일(IMAP/SMTP)과 캘린더(CalDAV)를 함께 쓸 수 있는
MCP 서버입니다. 개발자 콘솔 등록이나 관리자 승인 없이, 네이버웍스 **외부 앱 비밀번호**
하나만으로 동작합니다.

## 포함 파일

| 파일 | 설명 |
|---|---|
| `server.py` | MCP 서버 본체 (메일 19개 + 캘린더 10개, 총 29개 도구) |
| `requirements.txt` | 필요 패키지 목록 |
| `.env.example` | 자격증명 입력 템플릿 (`.env`는 포함되어 있지 않습니다 — 각자 발급받아야 함) |
| `setup.ps1` | 설치 자동화 스크립트 (venv 생성 + 패키지 설치 + `.env` 준비를 한 번에) |
| `signature.example.html` / `.txt` | `send_email(signature=True)`용 서명 예시 (복사해서 `signature.html` / `signature.txt`로 사용) |
| `.gitignore` | `.env`/`.venv`/개인 서명 파일 등을 커밋에서 제외 |
| `LICENSE` | MIT |

## 제공 도구 (총 29개)

**메일 (19개)**

| 도구 | 설명 |
|---|---|
| `list_emails` | 받은편지함 최근 메일 조회 |
| `list_unread_emails` | 읽지 않은 메일 조회 |
| `list_emails_from` | 특정 발신자 메일 조회 |
| `get_email` | 메일 ID로 전체 내용 조회 |
| `search_emails` | 제목 키워드로 검색 |
| `list_sent_emails` / `list_sent_emails_to` / `search_sent_emails` | 보낸메일함 조회/검색 |
| `search_contacts` | 이름 일부로 이메일 주소 검색 (받은/보낸메일함 From·To·Cc 헤더 기반, 정식 조직도 API 아님, v2.1 신규) |
| `list_attachments` / `save_attachment` / `read_attachment` | 첨부파일 목록/저장/텍스트 추출 |
| `send_email` | 메일 발송 (`cc`로 참조자 지정, 본문은 맑은 고딕 14px HTML로 자동 전송) |
| `forward_email` | 메일 전달 — 원본 발신자/날짜/제목/수신자를 아웃룩 스타일로 인용, 원본 HTML·인라인 이미지·첨부파일까지 그대로 함께 전달 (v2.0 신규) |
| `reply_email` | 메일 답장/전체회신 — `reply_all`로 발신자만/전체 선택, 아웃룩 스타일 인용 + 원본 HTML·인라인 이미지 보존, `In-Reply-To`/`References`로 스레딩 유지 (v2.1 신규) |
| `list_mail_folders` | 서버의 IMAP 폴더 목록 조회 (휴지통/보낸메일함 실제 폴더명 확인용) |
| `move_email_to_trash` | 메일을 휴지통 폴더로 이동 (`folder`로 출발 폴더 지정) |
| `move_email` | 메일을 임의 폴더로 이동 — 휴지통 → 받은편지함 복구 포함 |
| `delete_local_file` | 저장된 첨부파일 삭제 (허용 폴더 안에서만 동작) |

**캘린더 (10개)**

| 도구 | 설명 |
|---|---|
| `list_calendars` | 캘린더 목록 |
| `list_today_events` / `list_upcoming_events` / `list_events_by_date_range` | 일정 조회 |
| `get_event` / `search_events` | 일정 상세/검색 |
| `create_my_event` / `create_team_event` | 일정 등록 (본인/팀 캘린더, `attendees`로 참석자 추가 가능 — 아래 참고) |
| `update_event` / `delete_event` | 일정 수정/삭제 |

## 설치 (Windows, PowerShell)

### 0. 사전 준비

- Python 3.10 이상, Git (없으면 `winget install -e --id Python.Python.3.12`,
  `winget install -e --id Git.Git` 후 PowerShell을 새로 엽니다)
- 네이버웍스 **외부 앱 비밀번호** 발급: 네이버웍스 웹 → 프로필 → 개인 설정 → 보안 → 외부 앱 비밀번호
  (메일의 IMAP/SMTP, 캘린더 CalDAV 사용이 계정/관리자 설정에서 허용돼 있어야 합니다)

### 1. 저장소 받기

```powershell
git clone https://github.com/s4061285-create/naver-works-mcp.git
cd naver-works-mcp
```

### 2. 설치 — 방법 A — 자동 설치 스크립트 (권장)

```powershell
.\setup.ps1
```

venv 생성 → 패키지 설치 → `.env` 파일 준비(메모장 자동 오픈) → 설치 확인까지 한 번에
진행되고, 마지막에 Claude 데스크톱 config에 넣을 JSON을 화면에 출력해줍니다.

> PowerShell 실행 정책 때문에 스크립트가 막히면 `powershell -ExecutionPolicy Bypass -File .\setup.ps1`로 실행하세요.

### 2. 설치 — 방법 B — 수동 설치

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

copy .env.example .env
notepad .env
```

`requirements.txt`에 `mcp[cli]<2`를 고정해뒀습니다. 그대로 설치하면 `mcp` 2.x가 잡혀
`ModuleNotFoundError: mcp.server.fastmcp` 오류가 나므로 반드시 이 버전을 유지하세요.

## .env 설정

`.env`는 배포판에 포함되어 있지 않습니다 (비밀번호가 담기기 때문에 각자 발급받아야
합니다). `.env.example`을 참고해 아래 값을 채워주세요.

- `NAVER_EMAIL`: 본인 네이버웍스 로그인 이메일
- `NAVER_APP_PASSWORD`: 네이버웍스 웹 → 프로필 → 개인 설정 → 보안 → 외부 앱 비밀번호에서
  발급받은 값 (로그인 비밀번호 아님, 재확인 불가하니 발급 즉시 복사)
- `NAVER_CALDAV_SERVER`: 보통 그대로 둠 (`https://caldav.worksmobile.com`)
- `ATTACHMENT_DIR`: 첨부파일 기본 저장 폴더 (비우면 시스템 임시 폴더)
- `TRASH_FOLDER`: 휴지통 폴더명. **비워두면 자동 탐색**하며, 자동 탐색이 실패할 때만
  `list_mail_folders()`로 확인한 실제 폴더명을 넣어주세요
- `DELETE_ALLOWED_DIRS`: `delete_local_file`이 삭제해도 되는 폴더 (`;` 또는 `,`로 여러 개).
  여기에 없는 경로는 삭제 요청이 거부됩니다

- `SIGNATURE_HTML_FILE` / `SIGNATURE_TEXT_FILE`: (선택) 서명 파일 경로. 기본값은 서버 폴더의
  `signature.html` / `signature.txt`

이 비밀번호 하나로 메일과 캘린더가 모두 동작합니다.

### 서명 설정 (선택)

`send_email(signature=True)`를 쓰려면 예시 파일을 복사해 본인 정보로 수정합니다.
(`signature.html` / `signature.txt`는 `.gitignore`로 제외되어 커밋되지 않습니다.)

```powershell
copy signature.example.html signature.html
copy signature.example.txt signature.txt
```

파일이 없는 상태에서 `signature=True`로 호출하면 안내 오류가 반환됩니다. 로고·QR 이미지는
첨부가 아니라 `<img src="https://...">` 외부 URL 참조 방식입니다.

## 단독 실행 확인

```powershell
.\.venv\Scripts\python.exe server.py
```

에러 없이 멈춰 있으면 정상입니다 (stdio 대기). `Ctrl+C`로 종료하세요.

## Claude 데스크톱 등록

설정 파일 경로 (일반 설치 버전과 Microsoft Store 버전이 다릅니다):

```
일반 설치: %APPDATA%\Claude\claude_desktop_config.json
Store 버전: %LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json
```

`mcpServers`에 아래 항목을 추가합니다 (경로는 실제 설치 위치로 교체 — `setup.ps1`
마지막 출력에 본인 경로가 그대로 나옵니다).

```json
{
  "mcpServers": {
    "naver-works": {
      "command": "<설치 폴더>\\.venv\\Scripts\\python.exe",
      "args": ["<설치 폴더>\\server.py"]
    }
  }
}
```

비밀번호는 `.env` 파일에서 읽으므로 config에는 넣지 않습니다.

Claude Code를 쓴다면 아래처럼 등록할 수도 있습니다.

```powershell
claude mcp add naver-works -- "<설치 폴더>\.venv\Scripts\python.exe" "<설치 폴더>\server.py"
```

저장 후 앱을 완전히 종료하고 재시작합니다.

```powershell
taskkill /IM Claude.exe /F
```

## 사용 예시

```
최근 메일 5개 보여줘
안 읽은 메일 있어?
kim@example.com 이 보낸 메일 찾아줘
오늘 일정 알려줘
이번 주 일정 등록해줘: 내일 오후 3시 팀 회의
```

## 문제 해결

| 증상 | 원인/조치 |
|---|---|
| `ModuleNotFoundError: mcp.server.fastmcp` | `mcp` 2.x 설치됨 → `pip install "mcp[cli]<2"` 로 재설치 |
| `spawn ... python.exe ENOENT` | config 경로에 `naver-works-mcp` 폴더 단계가 빠짐 → 경로 재확인 |
| 401 / 인증 실패 | 로그인 비밀번호를 넣었을 가능성 → 외부 앱 비밀번호로 교체 |
| "캘린더를 찾지 못했습니다" | 이메일/비밀번호 오타, 또는 계정에 CalDAV 접근이 막혀 있음 |
| Claude에서 도구가 안 보임 | config 경로/JSON 문법 확인, `taskkill` 후 재시작, 아래 로그 확인 |

로그 파일 위치 (설정 파일과 같은 계층의 `logs` 폴더):
```
<설정 파일 경로와 동일한 상위 폴더>\logs\mcp-server-naver-works.log
```

## 참고

- 외부 앱 비밀번호는 계정 전체 권한을 가지므로 `.env` 파일에만 두고, 다른 사람과 공유하거나
  git에 커밋하지 마세요 (`.gitignore`에 이미 제외 처리되어 있습니다). 사용 중단 시
  네이버웍스 보안 설정에서 폐기하세요.
- 각 PC마다 자기 계정으로 발급받은 외부 앱 비밀번호로 `.env`를 개별 구성해야 합니다.

## 라이선스

[MIT](LICENSE)
