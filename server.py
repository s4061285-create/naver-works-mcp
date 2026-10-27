"""
Naver Works(LINE WORKS) 통합 MCP 서버 — 메일(IMAP/SMTP) + 캘린더(CalDAV).

기존에 별도로 운영하던 naver-works-mail-mcp와 naver-works-calendar-mcp를
하나의 MCP 서버로 통합한 버전입니다. 두 서버 모두 개발자 콘솔 등록/관리자
승인 없이, 네이버웍스 "외부 앱 비밀번호" 하나만으로 동작합니다
(외부 앱 비밀번호는 IMAP/SMTP뿐 아니라 CalDAV 로그인에도 사용 가능).

제공 도구 (총 29개)

[메일 - 19개]
- list_emails                : 받은편지함 최근 메일 조회
- list_unread_emails         : 읽지 않은 메일만 조회
- list_emails_from           : 특정 발신자 메일 조회
- get_email                  : 메일 ID로 특정 메일 전체 내용 조회
- search_emails              : 제목 키워드로 메일 검색
- list_sent_emails           : 보낸메일함 최근 메일 조회
- list_sent_emails_to        : 보낸메일함에서 특정 수신자에게 보낸 메일 조회
- search_sent_emails         : 보낸메일함에서 제목 키워드로 검색
- search_contacts            : 이름 일부로 이메일 주소 검색 (받은/보낸메일함 From·To·Cc 헤더 기반, 정식 조직도 API 아님)
- list_attachments           : 메일 첨부파일 목록 조회
- save_attachment            : 메일 첨부파일을 로컬에 저장
- read_attachment            : 메일 첨부파일 내용을 텍스트로 추출
- send_email                 : 메일 발송 (cc로 참조자 지정, attachments로 파일 첨부 가능)
- forward_email              : 메일 전달 (원본 발신자/날짜/제목/본문을 인용해 포함, 원본 첨부파일도 함께 전달 가능)
- reply_email                : 메일 답장/전체회신 (reply_all로 발신자만/전체 선택, 원본 인용 포함, In-Reply-To/References로 스레딩 유지)
- list_mail_folders          : 서버의 IMAP 폴더 목록 조회 (휴지통 폴더명 확인용)
- move_email_to_trash        : 메일을 휴지통으로 이동
- move_email                 : 메일을 임의의 폴더로 이동 (휴지통 → 받은편지함 복구 포함)
- delete_local_file          : save_attachment로 저장한 로컬 파일 삭제 (허용 폴더 내에서만)

[캘린더 - 10개]
- list_calendars             : 내 캘린더 목록 (이름/개수 확인용)
- list_today_events          : 오늘 일정
- list_upcoming_events       : 오늘부터 N일간 일정 (기본 7일)
- list_events_by_date_range  : 특정 기간 일정 (YYYY-MM-DD ~ YYYY-MM-DD)
- get_event                  : uid로 특정 일정 상세 조회
- search_events              : 제목/설명에 키워드가 포함된 일정 검색
- create_my_event            : 본인(기본) 캘린더에 일정 등록 (attendees로 참석자 초대 가능, 실험적)
- create_team_event          : 팀 공유 캘린더에 일정 등록 (list_calendars로 이름 확인 후 사용, attendees 지원)
- update_event               : uid로 기존 일정 수정
- delete_event               : uid로 기존 일정 삭제

참고: attendees(참석자 초대) 동작 방식 — 실제 테스트로 확인됨 (2026-09-16 테스트).
iCalendar ATTENDEE 속성을 추가해 CalDAV 서버로 저장하면, 네이버웍스가 참석자 본인의
캘린더에도 그 일정을 조용히 등록해준다. 다만 초대 알림(수락/거절 팝업, 메일 등)은
전혀 가지 않는다 — 참석자는 사전 동의 없이 자기 캘린더에 일정이 생긴 것을 나중에
발견하게 된다. 따라서 attendees로 초대할 때는 상대방에게 별도로 메일/메시지로
"일정 등록해뒀다"고 알려주는 것을 권장한다.
"""

import html as html_lib
import imaplib
import smtplib
import email
import io
import mimetypes
import os
import re
import tempfile
import uuid
from datetime import datetime, timedelta, timezone
from email import encoders
from email.header import decode_header
from email.utils import parsedate_to_datetime, getaddresses
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

import caldav
import icalendar
from caldav.lib.error import PutError
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

# 스크립트 위치 기준으로 .env를 로드한다 (Claude 데스크톱이 서버를 다른 작업
# 디렉터리에서 띄워도 항상 이 폴더의 .env를 찾도록 고정 경로 사용).
load_dotenv(Path(__file__).resolve().parent / ".env")

NAVER_EMAIL = os.getenv("NAVER_EMAIL", "")
NAVER_APP_PASSWORD = os.getenv("NAVER_APP_PASSWORD", "")
CALDAV_SERVER = os.getenv("NAVER_CALDAV_SERVER", "https://caldav.worksmobile.com").rstrip("/")

IMAP_SERVER = "imap.worksmobile.com"
IMAP_PORT = 993
SMTP_SERVER = "smtp.worksmobile.com"
SMTP_PORT = 465

ATTACHMENT_DIR = os.getenv("ATTACHMENT_DIR") or os.path.join(
    tempfile.gettempdir(), "naver_works_mail_attachments"
)

# 휴지통 폴더명. 비워두면 IMAP LIST 결과에서 자동 탐색한다.
# 자동 탐색이 실패하면 list_mail_folders()로 실제 폴더명을 확인해 .env에 지정한다.
TRASH_FOLDER = os.getenv("TRASH_FOLDER", "").strip()

# delete_local_file()이 삭제를 허용할 폴더 목록(OS 경로 구분자 ; 또는 , 로 구분).
# 안전장치이므로 비워두면 ATTACHMENT_DIR 아래 파일만 삭제할 수 있다.
DELETE_ALLOWED_DIRS = os.getenv("DELETE_ALLOWED_DIRS", "").strip()

MAX_EXTRACT_CHARS = 8000
TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".json", ".log", ".html", ".htm", ".xml", ".yaml", ".yml"}

KST = timezone(timedelta(hours=9))

mcp = FastMCP("naver-works")

_calendars_cache: Optional[list] = None


def _require_credentials() -> None:
    if not NAVER_EMAIL or not NAVER_APP_PASSWORD:
        raise RuntimeError(
            "NAVER_EMAIL / NAVER_APP_PASSWORD가 .env에 없습니다. "
            ".env.example을 .env로 복사하고 네이버웍스 외부 앱 비밀번호를 넣어주세요."
        )


# ---------------------------------------------------------------------------
# 메일 (IMAP/SMTP)
# ---------------------------------------------------------------------------


def _decode_str(value, encoding=None):
    if isinstance(value, bytes):
        return value.decode(encoding or "utf-8", errors="ignore")
    return value or ""


def _decode_header_full(value: str) -> str:
    """헤더 전체를 디코딩한다.

    decode_header()는 헤더 안에 인코딩된 단어(예: 한글 이름)가 여러 개 있으면
    각각을 별도 튜플로 쪼개서 반환한다. 첫 번째 튜플([0])만 쓰면 두 번째 수신자
    이후 이름/주소가 통째로 잘려나가므로(예: "박찬미 <a@x>, 전인우 <b@x>"에서
    전인우 누락), 모든 조각을 순서대로 이어붙여야 한다.
    """
    if not value:
        return ""
    parts = []
    for part, enc in decode_header(value):
        if isinstance(part, bytes):
            parts.append(part.decode(enc or "utf-8", errors="ignore"))
        else:
            parts.append(part)
    return "".join(parts)


def _connect_imap(mailbox: str = "INBOX"):
    _require_credentials()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(NAVER_EMAIL, NAVER_APP_PASSWORD)
    mail.select(f'"{mailbox}"')
    return mail


_SENT_FOLDER_CACHE = None


def _get_sent_folder() -> str:
    """보낸메일함 폴더명을 찾아 캐시합니다 (특수용도 플래그 \\Sent 우선)."""
    return _get_special_folder("sent")


# 폴더 별칭 → (IMAP 특수용도 플래그, 폴더명 힌트). 힌트는 소문자로 비교한다.
_SPECIAL_FOLDERS = {
    "sent":   ("\\Sent",   ("sent", "보낸")),
    "trash":  ("\\Trash",  ("trash", "deleted", "휴지통", "지운편지함")),
    "drafts": ("\\Drafts", ("draft", "임시")),
    "junk":   ("\\Junk",   ("junk", "spam", "스팸")),
}

# 사용자가 쓰는 말 → _SPECIAL_FOLDERS 키
_FOLDER_ALIASES = {
    "sent": "sent", "보낸메일함": "sent", "보낸편지함": "sent", "보낸": "sent",
    "trash": "trash", "휴지통": "trash", "지운편지함": "trash", "삭제": "trash",
    "deleted": "trash", "deleted messages": "trash",
    "drafts": "drafts", "draft": "drafts", "임시보관함": "drafts",
    "junk": "junk", "spam": "junk", "스팸": "junk",
}

_INBOX_ALIASES = ("", "inbox", "받은편지함", "받은메일함", "받은")

_SPECIAL_FOLDER_CACHE: dict = {}


def _get_special_folder(kind: str) -> str:
    """특수용도 폴더명을 찾아 캐시한다.

    우선순위: .env 지정값(휴지통의 TRASH_FOLDER) > IMAP 특수용도 플래그 > 폴더명 힌트.
    찾지 못하면 RuntimeError. list_mail_folders()로 실제 폴더명을 확인할 수 있다.
    """
    if kind == "trash" and TRASH_FOLDER:
        return TRASH_FOLDER
    if kind in _SPECIAL_FOLDER_CACHE:
        return _SPECIAL_FOLDER_CACHE[kind]

    flag, hints = _SPECIAL_FOLDERS[kind]
    boxes = list_mailboxes()

    for box in boxes:
        if flag in box["flags"]:
            _SPECIAL_FOLDER_CACHE[kind] = box["name"]
            return box["name"]

    for box in boxes:
        haystack = (box["name"] + " " + box["display"]).lower()
        if any(hint in haystack for hint in hints):
            _SPECIAL_FOLDER_CACHE[kind] = box["name"]
            return box["name"]

    raise RuntimeError(
        f"'{kind}' 폴더를 찾지 못했습니다. list_mail_folders()로 실제 폴더명을 확인한 뒤 "
        f"folder 인자에 그 폴더명을 그대로 넣어주세요."
    )


def _resolve_mailbox(folder: str) -> str:
    """folder 인자를 실제 IMAP 폴더명으로 바꾼다.

    "inbox"/"sent"/"trash"/"drafts"/"junk"와 한글 별칭을 지원하고,
    그 밖의 값은 IMAP 폴더명(list_mail_folders의 name)으로 보고 그대로 사용한다.
    """
    key = (folder or "inbox").strip()
    low = key.lower()
    if low in _INBOX_ALIASES:
        return "INBOX"
    if low in _FOLDER_ALIASES:
        return _get_special_folder(_FOLDER_ALIASES[low])
    return key


_MAILBOX_LINE_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.+)$')


def _parse_mailbox_line(line) -> tuple[str, str]:
    """IMAP LIST 응답 한 줄에서 (플래그 문자열, 폴더명)을 뽑아낸다."""
    decoded = line.decode(errors="ignore") if isinstance(line, bytes) else str(line)
    m = _MAILBOX_LINE_RE.match(decoded.strip())
    if not m:
        return "", decoded.strip().split()[-1].strip('"') if decoded.strip() else ""
    name = m.group("name").strip()
    if name.startswith('"') and name.endswith('"'):
        name = name[1:-1]
    return m.group("flags"), name


def _decode_mutf7(name: str) -> str:
    """IMAP modified UTF-7 폴더명을 사람이 읽는 문자열로 되돌린다.

    네이버웍스는 보통 영문 폴더명을 쓰지만, 한글 폴더명(예: 휴지통)을 쓰는
    계정에서도 자동 탐색이 동작하도록 디코딩을 시도한다. 실패하면 원문을 그대로 반환.
    """
    try:
        out, i = [], 0
        while i < len(name):
            ch = name[i]
            if ch != "&":
                out.append(ch)
                i += 1
                continue
            end = name.find("-", i + 1)
            if end == -1:
                out.append(ch)
                i += 1
                continue
            chunk = name[i + 1:end]
            if chunk == "":
                out.append("&")
            else:
                b64 = chunk.replace(",", "/")
                b64 += "=" * (-len(b64) % 4)
                import base64
                out.append(base64.b64decode(b64).decode("utf-16-be", errors="ignore"))
            i = end + 1
        return "".join(out)
    except Exception:
        return name


def list_mailboxes() -> list[dict]:
    """서버의 IMAP 폴더 목록을 [{flags, name, display}] 형태로 반환한다."""
    _require_credentials()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(NAVER_EMAIL, NAVER_APP_PASSWORD)
    typ, folders = mail.list()
    mail.logout()

    results = []
    if typ == "OK" and folders:
        for line in folders:
            flags, name = _parse_mailbox_line(line)
            if not name:
                continue
            results.append({"flags": flags, "name": name, "display": _decode_mutf7(name)})
    return results


def _get_trash_folder() -> str:
    """휴지통 폴더명을 반환한다 (.env의 TRASH_FOLDER > \\Trash 플래그 > 이름 힌트)."""
    return _get_special_folder("trash")


def _search_ids(mail, *criteria) -> list[bytes]:
    """IMAP SEARCH 래퍼.

    imaplib은 검색어를 ASCII로 인코딩하기 때문에 한글 키워드를 그대로 넘기면
    'ascii codec can't encode characters' 오류가 난다. 비ASCII가 섞여 있으면
    CHARSET UTF-8 + IMAP 리터럴로 보내고, 서버가 이를 거부하면 최근 메일의
    제목 헤더만 받아 파이썬에서 직접 대조하는 방식으로 대체한다.
    """
    non_ascii = [c for c in criteria if isinstance(c, str) and not c.isascii()]
    if not non_ascii:
        typ, data = mail.search(None, *criteria)
        return data[0].split() if typ == "OK" and data and data[0] else []

    if len(non_ascii) == 1:
        try:
            keyword = non_ascii[0]
            head = [c for c in criteria if c is not keyword]
            mail.literal = keyword.encode("utf-8")
            typ, data = mail.search("UTF-8", *head)
            if typ == "OK" and data and data[0] is not None:
                return data[0].split()
        except Exception:
            pass

    return _search_ids_fallback(mail, non_ascii[0])


def _search_ids_fallback(mail, keyword: str, scan: int = 200) -> list[bytes]:
    """서버 검색이 안 될 때: 최근 scan건의 제목 헤더만 받아 파이썬에서 필터링."""
    typ, data = mail.search(None, "ALL")
    if typ != "OK" or not data or not data[0]:
        return []
    ids = data[0].split()[-scan:]
    if not ids:
        return []

    typ, resp = mail.fetch(b",".join(ids), "(BODY.PEEK[HEADER.FIELDS (SUBJECT)])")
    if typ != "OK" or not resp:
        return []

    needle = keyword.lower()
    hits = []
    for item in resp:
        if not isinstance(item, tuple) or len(item) < 2:
            continue
        prefix = item[0].decode(errors="ignore") if isinstance(item[0], bytes) else str(item[0])
        m = re.match(r"^\s*(\d+)\s", prefix)
        if not m:
            continue
        raw = item[1].decode("utf-8", errors="ignore") if isinstance(item[1], bytes) else str(item[1])
        raw = raw.replace("\r\n", " ").replace("\n", " ").strip()
        if ":" in raw:
            raw = raw.split(":", 1)[1]
        if needle in _decode_header_full(raw.strip()).lower():
            hits.append(m.group(1).encode())
    return hits


def _find_contacts_in_mailbox(mail, keyword: str, scan: int = 300) -> list[tuple[str, str]]:
    """최근 scan건의 From/To/Cc 헤더에서 keyword(이름 일부)와 일치하는 (이름, 이메일)을 찾는다.

    정식 조직도 API 없이, 실제로 메일을 주고받은 이력에서 연락처를 추정하기 위한
    보조 함수. 이름이든 이메일 주소든 keyword가 부분 일치하면 후보로 포함한다.
    """
    typ, data = mail.search(None, "ALL")
    if typ != "OK" or not data or not data[0]:
        return []
    ids = data[0].split()[-scan:]
    if not ids:
        return []

    typ, resp = mail.fetch(b",".join(ids), "(BODY.PEEK[HEADER.FIELDS (FROM TO CC)])")
    if typ != "OK" or not resp:
        return []

    needle = keyword.lower()
    matches = []
    for item in resp:
        if not isinstance(item, tuple) or len(item) < 2:
            continue
        raw = item[1]
        raw = raw.decode("utf-8", errors="ignore") if isinstance(raw, bytes) else str(raw)
        header_msg = email.message_from_string(raw)
        for field in ("From", "To", "Cc"):
            value = _decode_header_full(header_msg.get(field, ""))
            if not value:
                continue
            for display_name, addr in getaddresses([value]):
                if not addr:
                    continue
                if needle in display_name.lower() or needle in addr.lower():
                    matches.append((display_name, addr))
    return matches


def _fetch_raw_message(mail, mail_id: str):
    _, data = mail.fetch(mail_id.encode(), "(RFC822)")
    for part in data:
        if isinstance(part, tuple):
            return email.message_from_bytes(part[1])
    return None


def _decode_filename(part) -> str:
    filename = part.get_filename()
    if not filename:
        return ""
    decoded, enc = decode_header(filename)[0]
    return _decode_str(decoded, enc) if isinstance(decoded, bytes) else decoded


def _list_attachment_parts(msg):
    attachments = []
    if not msg or not msg.is_multipart():
        return attachments
    for part in msg.walk():
        disposition = part.get_content_disposition()
        filename = _decode_filename(part)
        if disposition == "attachment" or (filename and disposition != "inline"):
            attachments.append(part)
    return attachments


def _list_inline_parts(msg):
    """본문 HTML에서 cid:로 참조하는 인라인 이미지 등을 반환한다(첨부파일 목록과는 별개)."""
    inline = []
    if not msg or not msg.is_multipart():
        return inline
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        disposition = part.get_content_disposition()
        if disposition == "attachment":
            continue
        if part.get("Content-ID") or disposition == "inline":
            inline.append(part)
    return inline


def _get_html_body(msg) -> str:
    """메일의 text/html 파트를 그대로(이스케이프 없이) 반환한다. 없으면 빈 문자열."""
    if not msg:
        return ""
    if msg.is_multipart():
        for p in msg.walk():
            if p.get_content_type() == "text/html" and p.get_content_disposition() != "attachment":
                return (p.get_payload(decode=True) or b"").decode("utf-8", errors="ignore")
        return ""
    if msg.get_content_type() == "text/html":
        return (msg.get_payload(decode=True) or b"").decode("utf-8", errors="ignore")
    return ""


def _parse_message(msg) -> dict:
    if not msg:
        return {}

    subject = _decode_header_full(msg.get("Subject", ""))
    from_ = _decode_header_full(msg.get("From", ""))
    to_ = _decode_header_full(msg.get("To", ""))
    cc_ = _decode_header_full(msg.get("Cc", ""))

    date = msg.get("Date", "")
    body = ""
    if msg.is_multipart():
        for p in msg.walk():
            if p.get_content_type() == "text/plain" and p.get_content_disposition() != "attachment":
                body = p.get_payload(decode=True).decode("utf-8", errors="ignore")
                break
    else:
        body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

    attachments = [_decode_filename(p) for p in _list_attachment_parts(msg)]

    return {
        "from": from_,
        "to": to_,
        "cc": cc_,
        "subject": subject,
        "date": date,
        "body": body.strip(),
        "attachments": attachments,
    }


@mcp.tool()
def list_emails(limit: int = 5, folder: str = "inbox") -> list[dict]:
    """최근 메일을 가져옵니다. folder: "inbox"(기본값) / "sent" / "trash" / "drafts" / "junk",
    또는 list_mail_folders()가 알려주는 IMAP 폴더명을 그대로 넣어도 됩니다."""
    mail = _connect_imap(_resolve_mailbox(folder))
    _, messages = mail.search(None, "ALL")
    ids = messages[0].split()[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def list_unread_emails(limit: int = 5, folder: str = "inbox") -> list[dict]:
    """읽지 않은 메일만 가져옵니다. folder 지정 방식은 list_emails와 같습니다."""
    mail = _connect_imap(_resolve_mailbox(folder))
    _, messages = mail.search(None, "UNSEEN")
    ids = messages[0].split()[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def list_emails_from(from_email: str, limit: int = 5) -> list[dict]:
    """특정 발신자가 보낸 메일을 가져옵니다."""
    mail = _connect_imap()
    _, messages = mail.search(None, f'FROM "{from_email}"')
    ids = messages[0].split()[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def get_email(mail_id: str, folder: str = "inbox") -> dict:
    """메일 ID로 특정 메일 전체 내용을 가져옵니다. folder: "inbox"(받은편지함, 기본값) 또는 "sent"(보낸메일함)."""
    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    parsed = _parse_message(msg)
    mail.close()
    mail.logout()
    return parsed


@mcp.tool()
def search_emails(query: str, limit: int = 5, folder: str = "inbox") -> list[dict]:
    """제목 키워드로 메일을 검색합니다(한글 가능). folder 지정 방식은 list_emails와 같습니다."""
    mail = _connect_imap(_resolve_mailbox(folder))
    ids = _search_ids(mail, "SUBJECT", query)[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def list_sent_emails(limit: int = 5) -> list[dict]:
    """보낸메일함에서 최근 메일을 가져옵니다."""
    mail = _connect_imap(_get_sent_folder())
    _, messages = mail.search(None, "ALL")
    ids = messages[0].split()[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def list_sent_emails_to(to_email: str, limit: int = 5) -> list[dict]:
    """보낸메일함에서 특정 수신자에게 보낸 메일을 가져옵니다."""
    mail = _connect_imap(_get_sent_folder())
    _, messages = mail.search(None, f'TO "{to_email}"')
    ids = messages[0].split()[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def search_sent_emails(query: str, limit: int = 5) -> list[dict]:
    """보낸메일함에서 제목 키워드로 메일을 검색합니다."""
    mail = _connect_imap(_get_sent_folder())
    ids = _search_ids(mail, "SUBJECT", query)[-limit:][::-1]

    results = []
    for mid in ids:
        msg = _fetch_raw_message(mail, mid.decode())
        parsed = _parse_message(msg)
        if parsed:
            parsed["id"] = mid.decode()
            results.append(parsed)

    mail.close()
    mail.logout()
    return results


@mcp.tool()
def search_contacts(name: str, limit: int = 10, scan_per_folder: int = 300) -> list[dict]:
    """이름(일부)으로 이메일 주소를 찾습니다. 네이버웍스 정식 조직도 API 연동 없이,
    받은편지함/보낸메일함에서 최근 메일들의 From/To/Cc 헤더를 뒤져 실제로 메일을
    주고받은 이력 기반으로 연락처를 추정합니다. 정식 주소록이 아니므로 결과가 여러 건이면
    사용자에게 어느 주소가 맞는지 확인한 뒤 사용하세요.

    Args:
        name: 찾으려는 사람 이름 또는 이름 일부(한글 가능). 이메일 주소 일부로도 검색됩니다.
        limit: 반환할 최대 후보 수 (이메일 기준 중복 제거, 기본 10)
        scan_per_folder: 폴더별로 최근 몇 건의 메일을 뒤질지 (기본 300)
    """
    _require_credentials()

    found: dict[str, dict] = {}
    for folder in ("inbox", "sent"):
        try:
            mail = _connect_imap(_resolve_mailbox(folder))
        except Exception:
            continue
        try:
            for display_name, addr in _find_contacts_in_mailbox(mail, name, scan=scan_per_folder):
                key = addr.lower()
                if key not in found:
                    found[key] = {"name": display_name, "email": addr, "found_in": folder}
        finally:
            try:
                mail.close()
            except Exception:
                pass
            try:
                mail.logout()
            except Exception:
                pass
        if len(found) >= limit:
            break

    return list(found.values())[:limit]


def _truncate(text: str, limit: int = MAX_EXTRACT_CHARS) -> str:
    if len(text) > limit:
        return text[:limit] + f"\n\n... (내용이 길어 {limit}자까지만 표시됨, 전체 {len(text)}자)"
    return text


def _extract_pdf_text(content: bytes) -> str:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(content))
    pages = [f"--- Page {i + 1} ---\n{page.extract_text() or ''}" for i, page in enumerate(reader.pages)]
    return "\n\n".join(pages)


def _extract_docx_text(content: bytes) -> str:
    from docx import Document

    doc = Document(io.BytesIO(content))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def _extract_xlsx_text(content: bytes) -> str:
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(content), data_only=True)
    sheets_text = []
    for ws in wb.worksheets:
        rows = [
            ", ".join("" if v is None else str(v) for v in row)
            for row in ws.iter_rows(values_only=True)
        ]
        sheets_text.append(f"--- Sheet: {ws.title} ---\n" + "\n".join(rows))
    return "\n\n".join(sheets_text)


def _extract_attachment_text(filename: str, content: bytes) -> str:
    ext = os.path.splitext(filename)[1].lower()

    try:
        if ext in TEXT_EXTENSIONS:
            text = content.decode("utf-8", errors="ignore")
        elif ext == ".pdf":
            text = _extract_pdf_text(content)
        elif ext == ".docx":
            text = _extract_docx_text(content)
        elif ext in (".xlsx", ".xlsm"):
            text = _extract_xlsx_text(content)
        else:
            return (
                f"⚠️ '{ext or '알 수 없음'}' 형식은 텍스트 추출을 지원하지 않습니다. "
                "(지원: txt/md/csv/json/log/html, pdf, docx, xlsx)"
            )
    except ImportError as e:
        return f"❌ 텍스트 추출에 필요한 라이브러리가 없습니다: {e}"
    except Exception as e:
        return f"❌ 텍스트 추출 실패: {e}"

    return _truncate(text.strip())


@mcp.tool()
def list_attachments(mail_id: str, folder: str = "inbox") -> list[dict]:
    """메일에 포함된 첨부파일 목록(파일명, 크기, 타입)을 가져옵니다. folder: "inbox"(받은편지함, 기본값) 또는 "sent"(보낸메일함)."""
    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    mail.close()
    mail.logout()

    results = []
    for part in _list_attachment_parts(msg):
        payload = part.get_payload(decode=True) or b""
        results.append(
            {
                "filename": _decode_filename(part),
                "content_type": part.get_content_type(),
                "size": len(payload),
            }
        )
    return results


@mcp.tool()
def save_attachment(mail_id: str, filename: str, save_dir: str = "", folder: str = "inbox") -> str:
    """메일의 첨부파일을 지정한 폴더(기본: 임시 폴더)에 저장하고 저장 경로를 반환합니다. folder: "inbox"(받은편지함, 기본값) 또는 "sent"(보낸메일함)."""
    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    mail.close()
    mail.logout()

    for part in _list_attachment_parts(msg):
        if _decode_filename(part) == filename:
            target_dir = save_dir or ATTACHMENT_DIR
            os.makedirs(target_dir, exist_ok=True)
            safe_name = os.path.basename(filename)
            path = os.path.join(target_dir, safe_name)
            with open(path, "wb") as f:
                f.write(part.get_payload(decode=True) or b"")
            return path

    return f"❌ 첨부파일 '{filename}'을(를) 찾을 수 없습니다."


@mcp.tool()
def read_attachment(mail_id: str, filename: str, folder: str = "inbox") -> dict:
    """메일 첨부파일 내용을 텍스트로 추출해서 반환합니다. Claude가 이 내용을 바로 분석/요약할 수 있습니다.
    지원 형식: txt/md/csv/json/log/html, pdf, docx, xlsx. 그 외 형식은 추출 불가 메시지를 반환합니다.
    folder: "inbox"(받은편지함, 기본값) 또는 "sent"(보낸메일함)."""
    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    mail.close()
    mail.logout()

    for part in _list_attachment_parts(msg):
        if _decode_filename(part) == filename:
            content = part.get_payload(decode=True) or b""
            text = _extract_attachment_text(filename, content)
            return {
                "filename": filename,
                "content_type": part.get_content_type(),
                "size": len(content),
                "text": text,
            }

    return {"filename": filename, "error": f"❌ 첨부파일 '{filename}'을(를) 찾을 수 없습니다."}


@mcp.tool()
def list_mail_folders() -> list[dict]:
    """서버의 IMAP 폴더 목록을 반환합니다. 휴지통/보낸메일함의 실제 폴더명을 확인할 때 사용합니다.
    name은 IMAP에 그대로 넘기는 값, display는 사람이 읽는 이름(한글 폴더명 디코딩)입니다."""
    return list_mailboxes()


def _move_message(mail_id: str, src_mailbox: str, dst_mailbox: str) -> dict:
    """메일 한 통을 src_mailbox에서 dst_mailbox로 옮긴다.

    IMAP MOVE를 먼저 시도하고, 서버가 MOVE를 지원하지 않으면
    COPY + \\Deleted 플래그 + EXPUNGE 방식으로 대체한다.
    """
    if src_mailbox.strip('"').lower() == dst_mailbox.strip('"').lower():
        return {"ok": False, "error": f"출발 폴더와 도착 폴더가 같습니다: {src_mailbox}"}

    mail = _connect_imap(src_mailbox)
    try:
        # 옮기고 나면 조회가 안 되므로, 어떤 메일이었는지 먼저 확보해 둔다.
        info = _parse_message(_fetch_raw_message(mail, mail_id)) or {}
        if not info:
            return {
                "ok": False,
                "error": f"'{src_mailbox}'에서 id={mail_id} 메일을 찾지 못했습니다. "
                         f"id는 삭제/이동으로 바뀔 수 있으니 목록을 다시 조회해주세요.",
            }

        quoted = f'"{dst_mailbox}"'
        method = ""
        moved = False

        try:
            typ, _ = mail._simple_command("MOVE", mail_id, quoted)
            if typ == "OK":
                moved, method = True, "MOVE"
        except Exception:
            moved = False

        if not moved:
            typ, _ = mail.copy(mail_id, quoted)
            if typ != "OK":
                return {
                    "ok": False,
                    "error": f"'{dst_mailbox}'(으)로 복사하지 못했습니다. 폴더명을 확인해주세요.",
                }
            mail.store(mail_id, "+FLAGS", "\\Deleted")
            mail.expunge()
            method = "COPY+EXPUNGE"

        return {
            "ok": True,
            "mail_id": mail_id,
            "from_folder": src_mailbox,
            "to_folder": dst_mailbox,
            "method": method,
            "subject": info.get("subject", ""),
            "from": info.get("from", ""),
        }
    finally:
        try:
            mail.close()
        except Exception:
            pass
        try:
            mail.logout()
        except Exception:
            pass


@mcp.tool()
def move_email_to_trash(mail_id: str, folder: str = "inbox") -> dict:
    """메일을 휴지통 폴더로 이동합니다(완전 삭제가 아니라 휴지통으로 옮기는 것).

    mail_id: list_emails 등이 돌려준 id.
    folder : 출발 폴더. "inbox"(받은편지함, 기본값) 또는 "sent"(보낸메일함).

    휴지통 폴더명은 자동 탐색하며, .env의 TRASH_FOLDER로 직접 지정할 수도 있습니다.
    """
    result = _move_message(mail_id, _resolve_mailbox(folder), _get_trash_folder())
    if result.get("ok"):
        result["trash_folder"] = result["to_folder"]
    return result


@mcp.tool()
def move_email(mail_id: str, to_folder: str, folder: str = "inbox") -> dict:
    """메일을 지정한 폴더로 옮깁니다. 휴지통에서 받은편지함으로 되돌릴 때도 씁니다.

    mail_id  : 해당 폴더를 조회했을 때 나온 id (폴더마다 번호가 다릅니다).
    to_folder: 도착 폴더. "inbox"/"sent"/"trash"/"drafts"/"junk" 또는 IMAP 폴더명.
    folder   : 출발 폴더. 기본값은 "inbox".

    예) 휴지통의 메일을 받은편지함으로 복구:
        move_email(mail_id="12", to_folder="inbox", folder="trash")

    주의: mail_id는 IMAP 시퀀스 번호라서 메일이 지워지면 앞당겨집니다.
    옮기기 직전에 해당 폴더를 조회해 얻은 id를 바로 사용하세요."""
    return _move_message(mail_id, _resolve_mailbox(folder), _resolve_mailbox(to_folder))


def _delete_allowed_roots() -> list[Path]:
    roots = [Path(ATTACHMENT_DIR)]
    if DELETE_ALLOWED_DIRS:
        for chunk in re.split(r"[;,]", DELETE_ALLOWED_DIRS):
            chunk = chunk.strip()
            if chunk:
                roots.append(Path(chunk))
    resolved = []
    for root in roots:
        try:
            resolved.append(root.expanduser().resolve())
        except Exception:
            continue
    return resolved


@mcp.tool()
def delete_local_file(path: str) -> dict:
    """로컬에 저장된 파일을 삭제합니다. save_attachment로 내려받은 첨부파일 정리에 사용합니다.

    안전장치: ATTACHMENT_DIR 또는 .env의 DELETE_ALLOWED_DIRS에 지정된 폴더(하위 폴더 포함)
    안의 파일만 삭제할 수 있습니다. 그 밖의 경로, 디렉터리, 심볼릭 링크는 거부합니다.
    """
    try:
        target = Path(path).expanduser().resolve()
    except Exception as e:
        return {"ok": False, "error": f"경로를 해석할 수 없습니다: {e}"}

    roots = _delete_allowed_roots()
    if not any(target == root or root in target.parents for root in roots):
        return {
            "ok": False,
            "error": "허용되지 않은 경로입니다. .env의 DELETE_ALLOWED_DIRS에 상위 폴더를 추가해주세요.",
            "allowed_dirs": [str(r) for r in roots],
        }

    if not target.exists():
        return {"ok": False, "error": f"파일이 없습니다: {target}"}
    if target.is_dir():
        return {"ok": False, "error": "디렉터리는 삭제하지 않습니다. 파일 경로를 지정해주세요."}
    if target.is_symlink():
        return {"ok": False, "error": "심볼릭 링크는 삭제하지 않습니다."}

    size = target.stat().st_size
    try:
        target.unlink()
    except Exception as e:
        return {"ok": False, "error": f"삭제 실패: {e}"}

    return {"ok": True, "deleted": str(target), "size": size}


MAIL_FONT_FAMILY = "'맑은 고딕', '맑은고딕', 'Malgun Gothic', sans-serif"
MAIL_FONT_SIZE = "14px"

# forward_email의 "-----Original Message-----" 인용 헤더 전용 폰트.
# 아웃룩 등 메일 솔루션이 자체 생성하는 전달 헤더가 Sans Serif 계열을 쓰는 것과
# 맞추기 위한 것으로, 원본 본문 자체의 폰트에는 영향을 주지 않는다.
FORWARD_HEADER_FONT_FAMILY = "sans-serif"

# 개인 서명: 코드에 하드코딩하지 않고 서버 폴더의 파일에서 읽는다.
#   signature.html : HTML 서명 (로고/QR 등은 <img src="https://...">로 외부 URL 참조)
#   signature.txt  : 같은 내용의 텍스트 서명 (multipart/alternative의 plain 파트용)
# 경로는 .env의 SIGNATURE_HTML_FILE / SIGNATURE_TEXT_FILE로 바꿀 수 있다.
# 예시는 signature.example.html / signature.example.txt 참고.
_BASE_DIR = Path(__file__).resolve().parent


def _load_signature_file(env_name: str, default_name: str) -> str:
    path = Path(os.getenv(env_name, "").strip() or default_name).expanduser()
    if not path.is_absolute():
        path = _BASE_DIR / path
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8").strip()


def _get_signature() -> tuple[str, str]:
    """(html, text) 서명을 반환한다. 서명 파일이 없으면 예외."""
    sig_html = _load_signature_file("SIGNATURE_HTML_FILE", "signature.html")
    sig_text = _load_signature_file("SIGNATURE_TEXT_FILE", "signature.txt")
    if not sig_html:
        raise RuntimeError(
            "서명 파일(signature.html)이 없습니다. signature.example.html을 "
            "signature.html로 복사해 내용을 수정하거나, signature=False로 호출하세요."
        )
    return sig_html, sig_text


def _build_html_body(body: str, signature: bool = False) -> str:
    """일반 텍스트 본문을 맑은 고딕 14px 스타일이 적용된 HTML로 감싼다.

    줄바꿈은 <br>로 변환하고, 본문에 들어있는 <, >, & 등은 이스케이프해서
    HTML이 깨지지 않도록 한다. signature=True면 본문 뒤에 로고/QR코드가 포함된
    HTML 서명을 그대로(이스케이프하지 않고) 덧붙인다.
    """
    escaped = html_lib.escape(body).replace("\n", "<br>")
    html = (
        "<div style=\"font-family:{font}; font-size:{size}; "
        "white-space:pre-wrap;\">{content}</div>"
    ).format(font=MAIL_FONT_FAMILY, size=MAIL_FONT_SIZE, content=escaped)
    if signature:
        html += "<br><br>" + _get_signature()[0]
    return html


def _attach_file(msg: MIMEMultipart, file_path: str) -> None:
    path = Path(file_path).expanduser()
    if not path.is_file():
        raise FileNotFoundError(file_path)

    ctype, encoding = mimetypes.guess_type(str(path))
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)

    part = MIMEBase(maintype, subtype)
    with open(path, "rb") as f:
        part.set_payload(f.read())
    encoders.encode_base64(part)
    # filename을 (charset, language, value) 3튜플로 넘기면 한글 등 비ASCII
    # 파일명도 RFC 2231 규격으로 자동 인코딩된다.
    part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", path.name))
    msg.attach(part)


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    signature: bool = False,
    attachments: Optional[list[str]] = None,
) -> str:
    """메일을 발송합니다. 본문은 맑은 고딕(맑은고딕) 14px 스타일이 자동 적용된
    HTML 메일로 전송됩니다 (일반 텍스트 클라이언트를 위한 대체 텍스트도 함께 포함).

    Args:
        to: 받는 사람 이메일 주소 (여러 명이면 쉼표로 구분)
        subject: 제목
        body: 본문 (줄바꿈은 그대로 유지되어 HTML에 반영됩니다)
        cc: 참조자 이메일 주소 (선택, 여러 명이면 쉼표로 구분)
        signature: True면 본문 끝에 개인 서명(signature.html, 로고/QR코드 포함 가능)을 자동으로
            붙인다. 신규로 작성하는 메일일 때만 True로 주고, 회신/답장/전달일
            때는 False(기본값)로 둔다.
        attachments: 첨부할 로컬 파일의 절대 경로 목록 (선택, 예:
            ["C:/Users/me/Desktop/보고서.pdf"]). save_attachment로 내려받은
            파일을 그대로 다시 첨부해서 보낼 때도 사용할 수 있습니다.
    """
    _require_credentials()

    alt = MIMEMultipart("alternative")
    plain_body = body + ("\n\n" + _get_signature()[1] if signature else "")
    # 텍스트 파트를 먼저 붙이고 HTML 파트를 나중에 붙인다 (MIME 관례상 클라이언트가
    # 마지막 파트, 즉 더 선호되는 표현을 우선 렌더링함).
    alt.attach(MIMEText(plain_body, "plain", "utf-8"))
    alt.attach(MIMEText(_build_html_body(body, signature=signature), "html", "utf-8"))

    if attachments:
        msg = MIMEMultipart("mixed")
        msg.attach(alt)
        for file_path in attachments:
            try:
                _attach_file(msg, file_path)
            except FileNotFoundError:
                return f"❌ 첨부파일을 찾을 수 없습니다: {file_path}"
    else:
        msg = alt

    msg["From"] = NAVER_EMAIL
    msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject

    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(NAVER_EMAIL, NAVER_APP_PASSWORD)
    server.send_message(msg)
    server.quit()

    result = f"✅ '{to}'로 메일 전송 성공"
    if cc:
        result += f" (참조: {cc})"
    if attachments:
        result += f" (첨부: {', '.join(Path(p).name for p in attachments)})"
    return result


def _format_forward_date(date_str: str) -> str:
    """이메일 Date 헤더를 아웃룩 전달 형식으로 변환한다.

    예: 'Thu, 17 Sep 2026 09:54:00 +0900' -> '2026. 9. 17. (목) 09:54 (GMT+09:00)'
    파싱에 실패하면 원본 문자열을 그대로 반환한다.
    """
    if not date_str:
        return date_str
    try:
        dt = parsedate_to_datetime(date_str)
    except (TypeError, ValueError):
        return date_str
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=KST)
    dt_kst = dt.astimezone(KST)
    weekday = "월화수목금토일"[dt_kst.weekday()]
    return f"{dt_kst.year}. {dt_kst.month}. {dt_kst.day}. ({weekday}) {dt_kst.strftime('%H:%M')} (GMT+09:00)"


def _format_forward_addresses(header_value: str, trailing_semicolon: bool = True) -> str:
    """From/To/Cc 헤더를 메일 솔루션의 전달 인용 형식으로 변환한다.

    예: 'Choi Yeon Ung <a@x.com>, Kim <b@x.com>'
        -> '"Choi Yeon Ung"<a@x.com>; "Kim"<b@x.com>;'
    이름이 없는 주소는 <addr> 형태만 남긴다. 값이 없으면 빈 문자열을 반환한다.
    From 줄은 세미콜론을 붙이지 않으므로 trailing_semicolon=False로 호출한다.
    """
    if not header_value:
        return ""
    formatted = []
    for name, addr in getaddresses([header_value]):
        if not addr:
            continue
        formatted.append(f'"{name}"<{addr}>' if name else f"<{addr}>")
    if not formatted:
        return ""
    joined = "; ".join(formatted)
    return joined + ";" if trailing_semicolon else joined


def _build_quoted_mime(msg, parsed: dict, new_text: str, include_attachments: bool):
    """원본 메일(msg/parsed)을 인용해서 전달/답장용 MIME 메시지를 만든다.

    forward_email과 reply_email이 공유하는 본문 조립 로직 — 아웃룩 스타일
    "-----Original Message-----" 인용 헤더, 원본 HTML/인라인 이미지 재사용,
    원본 첨부파일 재첨부까지 처리한다. From/To/Cc/Subject 헤더는 호출자가
    반환된 out_msg에 직접 설정한다.

    Returns:
        (out_msg, attached_names) 튜플.
    """
    header_rows = [
        ("From", _format_forward_addresses(parsed["from"], trailing_semicolon=False)),
        ("To", _format_forward_addresses(parsed["to"])),
        ("Cc", _format_forward_addresses(parsed["cc"])),
        ("Sent", _format_forward_date(parsed["date"])),
        ("Subject", parsed["subject"]),
    ]
    quoted_header = "-----Original Message-----\n" + "\n".join(
        f"{label}: {value}" for label, value in header_rows
    )

    body_sections = ([new_text] if new_text else []) + [quoted_header, parsed["body"]]
    body = "\n\n".join(body_sections)

    # 원본에 HTML 파트가 있으면 그대로 재사용해서(이스케이프하지 않음) 이미지 등
    # 서식이 살아있는 원본 본문을 그대로 보여준다.
    original_html = _get_html_body(msg)
    new_text_html = html_lib.escape(new_text).replace("\n", "<br>") if new_text else ""
    header_html = "\n".join(
        ["-----Original Message-----"]
        + [f"<b>{label}:</b> {html_lib.escape(value)}" for label, value in header_rows]
    )
    body_html_content = original_html or _build_html_body(parsed["body"])

    html_pieces = []
    if new_text_html:
        html_pieces.append(
            f'<div style="font-family:{FORWARD_HEADER_FONT_FAMILY}; font-size:{MAIL_FONT_SIZE}; white-space:pre-wrap;">{new_text_html}</div>'
        )
    html_pieces.append(
        f'<div style="font-family:{FORWARD_HEADER_FONT_FAMILY}; font-size:{MAIL_FONT_SIZE}; '
        f'white-space:pre-wrap;">{header_html}</div>'
    )
    html_pieces.append(body_html_content)
    html_body = "<br>".join(html_pieces)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(body, "plain", "utf-8"))
    alt.attach(MIMEText(html_body, "html", "utf-8"))

    # 원본 HTML이 cid:로 참조하는 인라인 이미지가 있으면 Content-ID를 유지한 채
    # 함께 실어 보내야 전달/답장 메일에서도 이미지가 깨지지 않고 보인다.
    inline_parts = _list_inline_parts(msg) if include_attachments else []
    if inline_parts:
        related = MIMEMultipart("related")
        related.attach(alt)
        for part in inline_parts:
            content = part.get_payload(decode=True) or b""
            ctype = part.get_content_type() or "application/octet-stream"
            maintype, subtype = ctype.split("/", 1) if "/" in ctype else ("application", "octet-stream")

            inline_part = MIMEBase(maintype, subtype)
            inline_part.set_payload(content)
            encoders.encode_base64(inline_part)
            content_id = part.get("Content-ID")
            if content_id:
                inline_part.add_header("Content-ID", content_id)
            inline_filename = _decode_filename(part) or "image"
            inline_part.add_header("Content-Disposition", "inline", filename=("utf-8", "", inline_filename))
            related.attach(inline_part)
        content_root = related
    else:
        content_root = alt

    out_msg = MIMEMultipart("mixed")
    out_msg.attach(content_root)

    attached_names = []
    if include_attachments:
        for part in _list_attachment_parts(msg):
            filename = _decode_filename(part) or "attachment"
            content = part.get_payload(decode=True) or b""
            ctype = part.get_content_type() or "application/octet-stream"
            maintype, subtype = ctype.split("/", 1) if "/" in ctype else ("application", "octet-stream")

            new_part = MIMEBase(maintype, subtype)
            new_part.set_payload(content)
            encoders.encode_base64(new_part)
            new_part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", filename))
            out_msg.attach(new_part)
            attached_names.append(filename)

    return out_msg, attached_names


@mcp.tool()
def forward_email(
    mail_id: str,
    to: str,
    folder: str = "inbox",
    cc: Optional[str] = None,
    comment: str = "",
    include_attachments: bool = True,
) -> str:
    """메일을 다른 사람에게 전달합니다. 원본 발신자/날짜/제목/수신자와 본문을
    인용문 형태로 포함해서 보내고, 원한다면 원본 첨부파일도 그대로 함께 전달합니다.
    원본이 HTML 메일이면 본문의 이미지(인라인 이미지 포함)도 최대한 그대로 살려서 보냅니다.
    (전달 메일에는 서명을 붙이지 않습니다.)

    Args:
        mail_id: 전달할 메일의 id (list_emails/get_email 등이 돌려준 값)
        to: 받는 사람 이메일 주소 (여러 명이면 쉼표로 구분)
        folder: 원본 메일이 있는 폴더. "inbox"(받은편지함, 기본값) 또는 "sent" 등
        cc: 참조자 이메일 주소 (선택, 여러 명이면 쉼표로 구분)
        comment: 전달하면서 맨 위에 덧붙일 한마디 메모 (선택)
        include_attachments: True(기본값)면 원본 메일의 첨부파일을 그대로 함께 첨부
    """
    _require_credentials()

    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    mail.close()
    mail.logout()

    parsed = _parse_message(msg)
    if not parsed:
        return f"❌ '{folder}'에서 id={mail_id} 메일을 찾지 못했습니다."

    original_subject = parsed["subject"]
    subject = original_subject if original_subject.lower().startswith("fw:") else f"FW: {original_subject}"

    out_msg, attached_names = _build_quoted_mime(msg, parsed, comment, include_attachments)

    out_msg["From"] = NAVER_EMAIL
    out_msg["To"] = to
    if cc:
        out_msg["Cc"] = cc
    out_msg["Subject"] = subject

    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(NAVER_EMAIL, NAVER_APP_PASSWORD)
    server.send_message(out_msg)
    server.quit()

    result = f"✅ '{to}'로 메일 전달 성공 (원본: {original_subject})"
    if cc:
        result += f" (참조: {cc})"
    if attached_names:
        result += f" (첨부: {', '.join(attached_names)})"
    return result


@mcp.tool()
def reply_email(
    mail_id: str,
    body: str,
    folder: str = "inbox",
    reply_all: bool = False,
    cc: Optional[str] = None,
    include_attachments: bool = False,
) -> str:
    """메일에 답장합니다. 원본 발신자/날짜/제목/수신자를 아웃룩 스타일 인용문으로
    포함해서 보내고, 가능하면 메일 스레딩(In-Reply-To/References)도 맞춥니다.
    (답장에는 서명을 붙이지 않습니다. 인사말/마무리 인사 등은 body에 직접 작성하세요.)

    Args:
        mail_id: 답장할 메일의 id (list_emails/get_email 등이 돌려준 값)
        body: 답장 본문 — 인용문 위에 붙는 새 내용 (인사말·마무리 인사 포함해서 직접 작성)
        folder: 원본 메일이 있는 폴더. "inbox"(받은편지함, 기본값) 또는 "sent" 등
        reply_all: True면 전체회신 — 원본 To/Cc에 있던 사람들도 참조에 넣는다
            (내 계정 주소는 자동으로 제외). False(기본값)면 원본 발신자에게만 답장.
        cc: 추가로 참조에 넣을 이메일 주소 (선택, 여러 명이면 쉼표로 구분).
            reply_all과 함께 쓰면 원본 To/Cc 목록에 추가로 합쳐진다.
        include_attachments: True면 원본 메일의 첨부파일을 그대로 다시 첨부한다
            (기본 False — 보통 답장에는 원본 첨부를 다시 붙이지 않음).
    """
    _require_credentials()

    mail = _connect_imap(_resolve_mailbox(folder))
    msg = _fetch_raw_message(mail, mail_id)
    mail.close()
    mail.logout()

    parsed = _parse_message(msg)
    if not parsed:
        return f"❌ '{folder}'에서 id={mail_id} 메일을 찾지 못했습니다."

    original_subject = parsed["subject"]
    subject = original_subject if original_subject.lower().startswith("re:") else f"RE: {original_subject}"

    to_addr = ""
    for _, addr in getaddresses([parsed["from"]]):
        if addr:
            to_addr = addr
            break
    if not to_addr:
        return f"❌ 원본 메일(id={mail_id})에서 회신 대상 주소를 찾지 못했습니다."

    my_addr = NAVER_EMAIL.strip().lower()
    cc_candidates = []
    if reply_all:
        for header_value in (parsed["to"], parsed["cc"]):
            if not header_value:
                continue
            for _, addr in getaddresses([header_value]):
                if addr and addr.lower() not in (my_addr, to_addr.lower()):
                    cc_candidates.append(addr)
    if cc:
        cc_candidates.extend(part.strip() for part in cc.split(",") if part.strip())

    seen = set()
    deduped_cc = []
    for addr in cc_candidates:
        key = addr.lower()
        if key not in seen:
            seen.add(key)
            deduped_cc.append(addr)
    final_cc = ", ".join(deduped_cc) if deduped_cc else None

    out_msg, attached_names = _build_quoted_mime(msg, parsed, body, include_attachments)

    out_msg["From"] = NAVER_EMAIL
    out_msg["To"] = to_addr
    if final_cc:
        out_msg["Cc"] = final_cc
    out_msg["Subject"] = subject

    message_id = msg.get("Message-ID") if msg else None
    if message_id:
        out_msg["In-Reply-To"] = message_id
        references = msg.get("References") if msg else None
        out_msg["References"] = f"{references} {message_id}" if references else message_id

    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(NAVER_EMAIL, NAVER_APP_PASSWORD)
    server.send_message(out_msg)
    server.quit()

    action = "전체회신" if reply_all else "답장"
    result = f"✅ '{to_addr}'로 {action} 성공 (원본: {original_subject})"
    if final_cc:
        result += f" (참조: {final_cc})"
    if attached_names:
        result += f" (첨부: {', '.join(attached_names)})"
    return result


# ---------------------------------------------------------------------------
# 캘린더 (CalDAV)
# ---------------------------------------------------------------------------


def _get_calendars() -> list:
    """캘린더 목록을 가져온다 (최초 1회만 서버에 접속, 이후 캐시)."""
    global _calendars_cache
    if _calendars_cache is not None:
        return _calendars_cache

    _require_credentials()
    client = caldav.DAVClient(
        url=CALDAV_SERVER,
        username=NAVER_EMAIL,
        password=NAVER_APP_PASSWORD,
    )  # type: ignore

    # 서버가 well-known 방식 자동 탐색을 지원하지 않을 수 있어, 실패하면
    # 포럼에 보고된 고정 경로(/principals/users/{email})로 다시 시도한다.
    try:
        principal = client.principal()
    except Exception:
        principal_url = f"{CALDAV_SERVER}/principals/users/{NAVER_EMAIL}/"
        principal = caldav.Principal(client=client, url=principal_url)  # type: ignore

    calendars = principal.calendars()
    if not calendars:
        raise RuntimeError(
            "캘린더를 찾지 못했습니다. 이메일/앱 비밀번호가 맞는지, "
            "네이버웍스에서 CalDAV 접근이 허용되어 있는지 확인해주세요."
        )
    _calendars_cache = calendars
    return calendars


def _comp_get(comp, key: str):
    val = comp.get(key)
    if val is None:
        return None
    return getattr(val, "dt", None) or str(val)


def _event_to_dict(obj, calendar_name: str) -> dict:
    comp = obj.icalendar_component
    start = _comp_get(comp, "dtstart")
    end = _comp_get(comp, "dtend")
    is_all_day = bool(start) and not isinstance(start, datetime)

    def _fmt(v):
        if v is None:
            return None
        if isinstance(v, datetime):
            if v.tzinfo is None:
                v = v.replace(tzinfo=KST)
            return v.astimezone(KST).strftime("%Y-%m-%dT%H:%M:%S+09:00")
        return str(v)  # date (전일 일정)

    attendees_raw = comp.get("attendee")
    if attendees_raw is None:
        attendees = []
    elif isinstance(attendees_raw, list):
        attendees = [str(a).replace("mailto:", "") for a in attendees_raw]
    else:
        attendees = [str(attendees_raw).replace("mailto:", "")]

    organizer = comp.get("organizer")
    organizer_email = str(organizer).replace("mailto:", "") if organizer else None

    return {
        "uid": str(comp.get("uid")) if comp.get("uid") else None,
        "summary": str(comp.get("summary")) if comp.get("summary") else None,
        "description": str(comp.get("description")) if comp.get("description") else None,
        "location": str(comp.get("location")) if comp.get("location") else None,
        "start": _fmt(start),
        "end": _fmt(end),
        "isAllDay": is_all_day,
        "calendar": calendar_name,
        "organizer": organizer_email,
        "attendees": attendees,
    }


def _fetch_range(start: datetime, end: datetime) -> list[dict]:
    results: list[dict] = []
    for cal in _get_calendars():
        name = cal.get_display_name() if hasattr(cal, "get_display_name") else str(cal.name)
        try:
            objs = cal.search(start=start, end=end, event=True, expand=True)
        except Exception as e:
            raise RuntimeError(f"'{name}' 캘린더 조회 실패: {e}") from e
        for obj in objs:
            try:
                results.append(_event_to_dict(obj, name))
            except Exception:
                continue  # 파싱 안 되는 항목(빈 타임존 컴포넌트 등)은 건너뜀
    results.sort(key=lambda e: e["start"] or "")
    return results


def _calendar_display_name(cal) -> str:
    return cal.get_display_name() if hasattr(cal, "get_display_name") else str(cal.name)


def _find_calendar(calendar_name: Optional[str] = None):
    """calendar_name이 없으면 목록의 첫 번째(기본) 캘린더를, 있으면 이름이 일치하는 캘린더를 반환."""
    calendars = _get_calendars()
    if calendar_name is None:
        return calendars[0]
    for cal in calendars:
        if _calendar_display_name(cal) == calendar_name:
            return cal
    names = ", ".join(_calendar_display_name(c) for c in calendars)
    raise ValueError(f"'{calendar_name}' 캘린더를 찾을 수 없습니다. 사용 가능한 캘린더: {names}")


def _parse_dt(value: str, all_day: bool):
    if all_day:
        return datetime.strptime(value, "%Y-%m-%d").date()
    # KST 고정 오프셋(timezone)은 TZID 이름이 없어 icalendar가 "UTC+09:00" 같은
    # 비표준 TZID를 만들어내고, 네이버웍스 서버는 이를 400 Bad Request로 거부한다.
    # UTC(Z)로 변환해서 보내면 TZID 없이도 시각이 정확히 전달된다.
    return datetime.strptime(value, "%Y-%m-%d %H:%M").replace(tzinfo=KST).astimezone(timezone.utc)


def _default_alarm_trigger_utc(dtstart) -> datetime:
    """종일 일정 기준 '하루 전 낮 12시(KST)'를 UTC로 변환해 반환한다."""
    event_date = dtstart.astimezone(KST).date() if isinstance(dtstart, datetime) else dtstart
    day_before = event_date - timedelta(days=1)
    local_noon = datetime(day_before.year, day_before.month, day_before.day, 12, 0, tzinfo=KST)
    return local_noon.astimezone(timezone.utc)


def _timed_alarm_trigger_utc(dtstart: datetime) -> datetime:
    """기간(시간 지정) 일정 기준 '시작 10분 전'을 UTC로 반환한다."""
    return dtstart - timedelta(minutes=10)


def _build_event_ical(dtstart, dtend, summary: str, description: Optional[str],
                       location: Optional[str], reminder: bool, all_day: bool,
                       attendees: Optional[list[str]] = None) -> str:
    """VEVENT(+옵션 VALARM, 옵션 ORGANIZER/ATTENDEE)를 직접 만든다.

    caldav의 kwargs 기반 헬퍼(vcal.create_ical)는 TRIGGER에 VALUE=DATE-TIME
    파라미터를 붙이지 않아, 절대시각 알림이 RFC5545 상 스펙 위반이 된다.
    (TRIGGER의 기본 값 타입은 DURATION이라 파라미터 없이는 상대시간으로 해석됨)
    그래서 icalendar로 직접 만들어 파라미터를 명시한다.

    attendees가 주어지면 ORGANIZER(본인)/ATTENDEE 속성을 추가한다. 이는 표준
    RFC5545 스케줄링 필드이지만, 네이버웍스 CalDAV 서버가 이를 받아서 실제
    초대 알림(수락/거절 UI)으로 처리해주는지는 확인되지 않은 실험적 기능이다.
    """
    cal = icalendar.Calendar()
    cal.add("prodid", "-//naver-works-mcp//KO")
    cal.add("version", "2.0")

    event = icalendar.Event()
    event.add("uid", str(uuid.uuid4()))
    event.add("dtstamp", datetime.now(timezone.utc))
    event.add("summary", summary)
    event.add("dtstart", dtstart)
    event.add("dtend", dtend)
    if description:
        event.add("description", description)
    if location:
        event.add("location", location)

    if attendees:
        if NAVER_EMAIL:
            event.add("organizer", f"mailto:{NAVER_EMAIL}")
        for addr in attendees:
            attendee = icalendar.vCalAddress(f"mailto:{addr}")
            attendee.params["cn"] = icalendar.vText(addr)
            attendee.params["role"] = icalendar.vText("REQ-PARTICIPANT")
            attendee.params["partstat"] = icalendar.vText("NEEDS-ACTION")
            attendee.params["rsvp"] = icalendar.vText("TRUE")
            event.add("attendee", attendee, encode=0)

    if reminder:
        if all_day:
            trigger = _default_alarm_trigger_utc(dtstart)
            alarm_desc = f"[하루 전 알림] {summary}"
        else:
            trigger = _timed_alarm_trigger_utc(dtstart)
            alarm_desc = f"[10분 전 알림] {summary}"
        alarm = icalendar.Alarm()
        alarm.add("action", "DISPLAY")
        alarm.add("description", alarm_desc)
        alarm.add("trigger", trigger, parameters={"VALUE": "DATE-TIME"})
        event.add_component(alarm)

    cal.add_component(event)
    return cal.to_ical().decode("utf-8")


def _create_event(cal, summary: str, start: str, end: str, description: Optional[str],
                   location: Optional[str], all_day: bool, reminder: bool = True,
                   attendees: Optional[list[str]] = None) -> dict:
    cal_name = _calendar_display_name(cal)

    dtstart = _parse_dt(start, all_day)
    dtend = _parse_dt(end, all_day)
    if all_day:
        # iCalendar 종일 일정은 종료일이 배타적(exclusive)이라 하루를 더해준다.
        dtend = dtend + timedelta(days=1)
    if dtend <= dtstart:
        raise ValueError("end가 start보다 앞설 수 없습니다.")

    ical = _build_event_ical(dtstart, dtend, summary, description, location, reminder, all_day,
                              attendees=attendees)

    try:
        obj = cal.save_event(ical=ical)
    except Exception as e:
        raise RuntimeError(f"'{cal_name}' 캘린더에 일정 등록 실패: {e}") from e

    return _event_to_dict(obj, cal_name)


@mcp.tool()
def create_my_event(
    summary: str,
    start: str,
    end: str,
    description: Optional[str] = None,
    location: Optional[str] = None,
    all_day: bool = False,
    reminder: bool = True,
    attendees: Optional[list[str]] = None,
) -> dict:
    """
    본인(기본) 캘린더에 새 일정을 등록합니다. list_calendars 결과의 첫 번째 캘린더에 등록됩니다.
    기본적으로 알림이 추가됩니다 (종일 일정은 하루 전 낮 12시, 기간 일정은 시작 10분 전).

    Args:
        summary: 일정 제목
        start: 시작 일시. all_day=False면 "YYYY-MM-DD HH:MM", all_day=True면 "YYYY-MM-DD"
        end: 종료 일시 (형식은 start와 동일, 종일 일정도 종료일 당일을 그대로 입력)
        description: 설명 (선택)
        location: 장소 (선택)
        all_day: 종일 일정 여부
        reminder: 알림 추가 여부 (기본 True, False면 알림 없이 등록. 종일 일정은 하루 전 낮 12시,
            기간 일정은 시작 10분 전에 알림)
        attendees: 초대할 참석자 이메일 목록 (선택, 예: ["kim@example.com"]). 참석자 본인의
            캘린더에도 이 일정이 등록됩니다 (테스트로 확인됨). 단, 초대 알림은 가지 않으므로
            참석자에게는 별도로 메일/메시지로 알려주는 것을 권장합니다.
    """
    cal = _find_calendar(None)
    return _create_event(cal, summary, start, end, description, location, all_day, reminder,
                          attendees=attendees)


@mcp.tool()
def create_team_event(
    calendar_name: str,
    summary: str,
    start: str,
    end: str,
    description: Optional[str] = None,
    location: Optional[str] = None,
    all_day: bool = False,
    reminder: bool = True,
    attendees: Optional[list[str]] = None,
) -> dict:
    """
    팀 공유 캘린더에 새 일정을 등록합니다. calendar_name은 list_calendars 결과의 name과
    정확히 일치해야 합니다. 기본적으로 알림이 추가됩니다 (종일 일정은 하루 전 낮 12시,
    기간 일정은 시작 10분 전).

    Args:
        calendar_name: 등록할 공유 캘린더 이름 (list_calendars 결과의 name)
        summary: 일정 제목
        start: 시작 일시. all_day=False면 "YYYY-MM-DD HH:MM", all_day=True면 "YYYY-MM-DD"
        end: 종료 일시 (형식은 start와 동일, 종일 일정도 종료일 당일을 그대로 입력)
        description: 설명 (선택)
        location: 장소 (선택)
        all_day: 종일 일정 여부
        reminder: 알림 추가 여부 (기본 True, False면 알림 없이 등록. 종일 일정은 하루 전 낮 12시,
            기간 일정은 시작 10분 전에 알림)
        attendees: 초대할 참석자 이메일 목록 (선택, 예: ["kim@example.com"]). 참석자 본인의
            캘린더에도 이 일정이 등록됩니다 (테스트로 확인됨). 단, 초대 알림은 가지 않으므로
            참석자에게는 별도로 메일/메시지로 알려주는 것을 권장합니다.
    """
    cal = _find_calendar(calendar_name)
    return _create_event(cal, summary, start, end, description, location, all_day, reminder,
                          attendees=attendees)


# 캘린더별 이벤트 uid 검색 시 넓게 잡는 기간(일). caldav의 event_by_uid()는 기간을
# 지정하지 않고 검색하는데, 이 서버 환경에서는 그 fallback 로직이 OSError를 내서
# 실패하므로, 명시적으로 넓은 기간을 지정해 우회한다.
EVENT_SEARCH_WINDOW_DAYS = 1825  # 약 5년


def _find_event_by_uid(uid: str, calendar_name: Optional[str] = None):
    """uid로 일정을 찾아 (캘린더, 이벤트) 튜플을 반환한다."""
    calendars = [_find_calendar(calendar_name)] if calendar_name else _get_calendars()
    now = datetime.now(KST)
    start = now - timedelta(days=EVENT_SEARCH_WINDOW_DAYS)
    end = now + timedelta(days=EVENT_SEARCH_WINDOW_DAYS)
    for cal in calendars:
        try:
            objs = cal.search(start=start, end=end, event=True, expand=False)
        except Exception:
            continue
        for obj in objs:
            try:
                if str(obj.icalendar_component.get("uid")) == uid:
                    return cal, obj
            except Exception:
                continue
    raise RuntimeError(
        f"uid={uid} 일정을 찾지 못했습니다. (검색 범위: 오늘 기준 ±{EVENT_SEARCH_WINDOW_DAYS}일)"
    )


def _apply_event_updates(comp, summary, start, end, description, location, all_day) -> None:
    current_start = _comp_get(comp, "dtstart")
    current_all_day = bool(current_start) and not isinstance(current_start, datetime)
    effective_all_day = current_all_day if all_day is None else all_day

    if start is not None or end is not None:
        if start is None or end is None:
            raise ValueError("start와 end는 함께 지정해야 합니다.")
        dtstart = _parse_dt(start, effective_all_day)
        dtend = _parse_dt(end, effective_all_day)
        if effective_all_day:
            dtend = dtend + timedelta(days=1)
        if dtend <= dtstart:
            raise ValueError("end가 start보다 앞설 수 없습니다.")
        comp.pop("dtstart", None)
        comp.add("dtstart", dtstart)
        comp.pop("dtend", None)
        comp.add("dtend", dtend)
    elif all_day is not None and all_day != current_all_day:
        raise ValueError("all_day를 변경하려면 start와 end도 함께 지정해주세요.")

    if summary is not None:
        comp.pop("summary", None)
        comp.add("summary", summary)
    if description is not None:
        comp.pop("description", None)
        comp.add("description", description)
    if location is not None:
        comp.pop("location", None)
        comp.add("location", location)


@mcp.tool()
def update_event(
    uid: str,
    calendar_name: Optional[str] = None,
    summary: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
    description: Optional[str] = None,
    location: Optional[str] = None,
    all_day: Optional[bool] = None,
) -> dict:
    """
    uid로 기존 일정을 수정합니다. 값을 넘기지 않은 필드는 그대로 유지됩니다.

    Args:
        uid: 수정할 일정의 uid (목록/생성 결과에 포함됨)
        calendar_name: 일정이 있는 캘린더 이름 (알고 있으면 검색이 빨라짐, 생략 가능)
        summary: 새 제목 (선택)
        start: 새 시작 일시. start/end는 항상 함께 지정 (형식은 create_my_event와 동일)
        end: 새 종료 일시 (선택, start와 함께 지정)
        description: 새 설명 (선택)
        location: 새 장소 (선택)
        all_day: 종일 일정 여부 변경 (선택, 변경하려면 start/end도 함께 지정)
    """
    cal, obj = _find_event_by_uid(uid, calendar_name)
    comp = obj.icalendar_component
    _apply_event_updates(comp, summary, start, end, description, location, all_day)
    try:
        obj.save()
    except PutError as e:
        # 네이버웍스 서버는 일정 수정 PUT 성공 시 204/201이 아닌 200 OK를 돌려주는데,
        # caldav 라이브러리는 이를 실패로 오판해 PutError를 던진다. 2xx 응답이면
        # 실제로는 반영된 것이므로 그대로 성공 처리한다.
        status = re.match(r"(\d{3})", e.url or "")
        if not (status and 200 <= int(status.group(1)) < 300):
            raise RuntimeError(f"uid={uid} 일정 수정 실패: {e}") from e
    except Exception as e:
        raise RuntimeError(f"uid={uid} 일정 수정 실패: {e}") from e
    return _event_to_dict(obj, _calendar_display_name(cal))


@mcp.tool()
def delete_event(uid: str, calendar_name: Optional[str] = None) -> dict:
    """
    uid로 기존 일정을 삭제합니다.

    Args:
        uid: 삭제할 일정의 uid (목록/생성 결과에 포함됨)
        calendar_name: 일정이 있는 캘린더 이름 (알고 있으면 검색이 빨라짐, 생략 가능)
    """
    cal, obj = _find_event_by_uid(uid, calendar_name)
    summary_val = obj.icalendar_component.get("summary")
    summary = str(summary_val) if summary_val else None
    try:
        obj.delete()
    except Exception as e:
        raise RuntimeError(f"uid={uid} 일정 삭제 실패: {e}") from e
    return {"deleted": True, "uid": uid, "calendar": _calendar_display_name(cal), "summary": summary}


@mcp.tool()
def list_calendars() -> list[dict]:
    """접근 가능한 캘린더 목록을 반환합니다 (이름 확인용)."""
    out = []
    for cal in _get_calendars():
        name = cal.get_display_name() if hasattr(cal, "get_display_name") else str(cal.name)
        out.append({"name": name, "url": str(cal.url)})
    return out


@mcp.tool()
def list_today_events() -> list[dict]:
    """오늘(한국시간 00:00~23:59) 일정 목록을 반환합니다."""
    now = datetime.now(KST)
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=1)
    return _fetch_range(start, end)


@mcp.tool()
def list_upcoming_events(days: int = 7) -> list[dict]:
    """오늘부터 days일 뒤까지의 일정 목록을 반환합니다. (기본 7일)"""
    if days < 1 or days > 90:
        raise ValueError("days는 1~90 사이여야 합니다.")
    now = datetime.now(KST)
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=days)
    return _fetch_range(start, end)


@mcp.tool()
def list_events_by_date_range(from_date: str, to_date: str) -> list[dict]:
    """
    지정한 기간의 일정 목록을 반환합니다.

    Args:
        from_date: 시작일, "YYYY-MM-DD"
        to_date: 종료일, "YYYY-MM-DD" (포함)
    """
    start = datetime.strptime(from_date, "%Y-%m-%d").replace(tzinfo=KST)
    end = datetime.strptime(to_date, "%Y-%m-%d").replace(tzinfo=KST) + timedelta(days=1)
    if end <= start:
        raise ValueError("to_date가 from_date보다 앞설 수 없습니다.")
    return _fetch_range(start, end)


@mcp.tool()
def get_event(uid: str, calendar_name: Optional[str] = None) -> dict:
    """uid로 특정 일정의 상세 정보를 조회합니다. (uid는 목록 조회 결과에 포함됨)"""
    cal, obj = _find_event_by_uid(uid, calendar_name)
    return _event_to_dict(obj, _calendar_display_name(cal))


@mcp.tool()
def search_events(keyword: str, days: int = 30) -> list[dict]:
    """
    앞으로 days일 이내 일정 중 제목/설명에 keyword가 포함된 일정만 찾습니다.
    (대소문자 구분 없음)

    Args:
        keyword: 검색어
        days: 오늘부터 며칠 이내를 검색할지 (기본 30일)
    """
    events = list_upcoming_events(days=days)
    kw = keyword.lower()
    return [
        e
        for e in events
        if kw in (e.get("summary") or "").lower() or kw in (e.get("description") or "").lower()
    ]


if __name__ == "__main__":
    _require_credentials()
    mcp.run()
