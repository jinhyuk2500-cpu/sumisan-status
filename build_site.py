"""수미산 작업 현황 사이트 만들기.
사용법: python3 build_site.py <취합.xlsx> <비밀번호파일> <template.html> <출력 index.html>
엑셀의 '작업지시서'·'사전조사의뢰서' 시트를 읽어 목록을 만들고, 비밀번호로 AES-GCM 암호화해 template에 넣는다.
"""
import sys, json, re, os, base64, datetime
from openpyxl import load_workbook
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

xlsx, pwfile, tpl, out = sys.argv[1:5]
ITER = 600000
pw = open(pwfile, encoding="utf-8").read().strip()

def d(v): return v.strftime('%Y-%m-%d') if isinstance(v, (datetime.date, datetime.datetime)) else ""
def s(v):
    v = "" if v is None else str(v)
    v = re.sub(r'\(?0\d{1,2}-\d{3,4}-\d{4}\)?', '', v)   # 전화번호 제거
    return v.strip()
def who(v):
    m = re.match(r'([가-힣]{2,4}) ?(매니저|책임)', s(v)); return (m.group(1) + " " + m.group(2)) if m else s(v)

# 현장 이름 통일 (같은 현장인데 다르게 적힌 것 합치기). 새 현장은 자동 규칙으로 처리.
ALIAS = [
    (r'북위례|센트럴위례', '힐스테이트 센트럴위례 (북위례)'),
    (r'공평', '공평오피스 (서울 공평 15·16지구)'),
    (r'구리갈매|테라타워', '현대테라타워 구리갈매'),
    (r'분당 ?구미동', '분당 구미동 오피스텔'),
    (r'상봉', '상봉 프레미어스 엠코'),
    (r'노원', '노원 프레미어스 엠코'),
]
def group(site):
    t = s(site)
    for pat, name in ALIAS:
        if re.search(pat, t): return name
    t = re.sub(r'\s*\(.*?\)\s*', ' ', t).strip()          # (1단지) 같은 괄호 제거
    t = re.sub(r'^HS\s*', '힐스테이트 ', t).strip()          # HS → 힐스테이트
    t = re.sub(r'\s+', ' ', t)
    return t or '현장 미상'

wb = load_workbook(xlsx, data_only=True)
items = []
ws = wb['작업지시서']
for r in range(5, ws.max_row + 1):
    c = lambda k: ws.cell(r, k).value
    no = c(1)
    if not isinstance(no, int): continue
    amt = c(14); amt = f"{amt:,}" if isinstance(amt, (int, float)) else s(amt)
    items.append(dict(id=f"w{no:02d}", kind="w", no=no, recv=d(c(2)), site=s(c(4)), title=s(c(6)), docno=s(c(5)),
        start=d(c(11)), due=d(c(12)), amount=amt, sender=who(c(3)),
        detail=" / ".join(x for x in [s(c(7)), ("공사범위: " + s(c(9))) if s(c(9)) else ""] if x),
        cause=s(c(8)), request=s(c(16)), subject=s(c(18)), status=s(c(19)) or "확인 필요", statusDate="", memo=s(c(20))))
ws = wb['사전조사의뢰서']
for r in range(5, ws.max_row + 1):
    c = lambda k: ws.cell(r, k).value
    no = c(1)
    if not isinstance(no, int): continue
    items.append(dict(id=f"p{no:02d}", kind="p", no=no, recv=d(c(2)), site=s(c(4)), title=s(c(7)), docno=s(c(5)),
        start=d(c(10)), due=d(c(11)), amount="", sender=who(c(3)),
        detail=(s(c(6)) + " · " + s(c(7))) if s(c(6)) else s(c(7)), cause=s(c(8)), request=s(c(13)), subject=s(c(15)),
        status=s(c(16)) or "확인 필요", statusDate="", memo=s(c(17))))

for it in items: it["group"] = group(it["site"])
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
data = json.dumps({"updated": now.strftime("%Y-%m-%d %H:%M"), "items": items}, ensure_ascii=False).encode()
salt, iv = os.urandom(16), os.urandom(12)
key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())
ct = AESGCM(key).encrypt(iv, data, None)
e = lambda b: base64.b64encode(b).decode()
payload = json.dumps({"salt": e(salt), "iv": e(iv), "iter": ITER, "ct": e(ct)})
html = open(tpl, encoding="utf-8").read().replace("__PAYLOAD__", payload)
open(out, "w", encoding="utf-8").write(html)
print("items", len(items), "bytes", len(html))
