"""26FW 어반사이드 소진율 대시보드 데이터 빌드.

Google Drive 원본 → 집계 → 비밀번호 암호화 → 26fw-sellthrough/data/data.enc.json
평문 JSON은 디스크(저장소)에 남기지 않는다.

원본 (고정 ID, _자동백업_/_백업_/_bak_ 파일 사용 금지)
  허브   26FW어패럴_재무모델_..._허브.xlsx        → season_kpi, monthly, products(허브 판매분)
  재고   재고관리시트(어반사이드) 구글시트          → 창고 재고 (inventory.available_qty)
  팝업   26FW 팝업 현황.xlsx                       → 팝업 재고·판매 (popup_*)
  주문   채널별 주문 원본 폴더 5개                  → recent3d, color_size

창고 재고(available_qty)와 팝업 재고(popup_available_qty)는 절대 합산하지 않는다.

인증 (먼저 찾는 순서)
  GDRIVE_SA_KEY            서비스 계정 JSON 내용
  GDRIVE_OAUTH_TOKEN       OAuth authorized_user JSON 내용 (refresh_token 포함)
  GDRIVE_OAUTH_TOKEN_FILE  위 JSON 파일 경로 (로컬 개발용)
  GOOGLE_APPLICATION_CREDENTIALS  서비스 계정 키 파일 경로
암호화 비밀번호: DASH_PASSWORD

종료 코드: 0 = 성공(변경 없음 포함), 2 = 검증 실패(쓰지 않음), 1 = 그 밖의 오류
GitHub Actions에서는 변경 여부를 $GITHUB_OUTPUT 의 changed/as_of 로 알려준다.
"""
import base64, collections, datetime as dt, hashlib, io, json, os, re, sys, warnings

import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

warnings.filterwarnings('ignore')
KST = dt.timezone(dt.timedelta(hours=9))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, '26fw-sellthrough', 'data', 'data.enc.json')
ITER = 100_000

SRC = {
    'hub': '1cb2i1XsuMNeJwWRhqlRObpI9DcO_-iXO',
    'inventory': '1ykrQdlyTKAHmf3qgtfAwHLiLgNeFJ5WD3n0wmjxU4I0',
    'popup': '1ont7N0rLV5BvrMz0F9wGS_pB0LlVAvn-',
}
INV_TAB = '일자 별 재고조회 (당시재고)'
JPY_KRW = 9  # 오모테산도 엔화 판매 환산 (¥1=9원, 라인시트 기준)
SEASON_START = dt.date(2026, 7, 1)
MONTHS = ['8월', '9월', '10월', '11월', '12월', '1월', '2월']
MONTH_START = [dt.date(2026, 8, 1), dt.date(2026, 9, 1), dt.date(2026, 10, 1), dt.date(2026, 11, 1),
               dt.date(2026, 12, 1), dt.date(2027, 1, 1), dt.date(2027, 2, 1), dt.date(2027, 3, 1)]

NAMES = [
    '에일러론 울트라 1000 SUL 7D 후드 다운 재킷', '에일러론 울트라 1000 SUL 후드 다운 재킷', '에일러론 울트라 1000 SUL 다운 재킷',
    '에일러론 900 UL 후드 다운 재킷', '에일러론 900 익스페디션 다운 파카', '써모코어 액티브 빌레이 파카', '써모코어 액티브 후드 재킷',
    '제로인 옥타 하이브리드 쉘 베스트', '제로인 옥타 하이브리드 쉘 후드 재킷', '제로인 옥타 인슐레이션 팬츠', '플랩 가루이시 후디',
    '플랩 90 폴라텍 알파 다이렉트 후디', '플랩 190 폴라텍 알파 다이렉트 후디', '미스릴 옥타 에어 베이스 레이어 후디', '제로인 옥타 에어 타이즈',
    '이지모드 프라임플렉스 써모 팬츠', '제로인 옥타 에어 후디', '제로인 옥타 에어 롱슬리브 티셔츠', '메리노 울 롱슬리브 티셔츠', '메리노 울 후디',
    '플랩 90 폴라텍 알파 다이렉트 넥 워머', '플랩 90 폴라텍 알파 다이렉트 바라클라바', '제로인 옥타 하이브리드 캡',
]
ALIAS = {'제로인 옥타 에어 베이스 레이어 후디': '미스릴 옥타 에어 베이스 레이어 후디'}
PATTERNS = sorted([(n, n) for n in NAMES] + list(ALIAS.items()), key=lambda x: -len(x[0]))  # 긴 이름 우선

# 주문 원본 (recent3d·color_size). 무신사·팝업은 포함하지 않는다.
CHANNELS = {
    '자사몰': dict(folder='1Fbo2_utwvj5UcHa6LUr6iEc5rHJ1PD78', prefix='매출_국내_자사몰_', date='결제일시(입금확인일)',
                name='주문상품명(옵션포함)', qty='수량', status='주문상태', key=('주문번호', '자체 품목코드'),
                code='주문상태코드', sku='자체 품목코드'),
    '4910': dict(folder='1N8sAVyqYXVROKf5JJRMwUyw-MfLEU-q-', prefix='매출_국내_4910_', date='결제일', name='상품명', qty='수량',
                 status='주문상태', key=('상품주문번호',), sku='판매자 상품코드', opt='옵션 정보'),
    '29CM': dict(folder='1zUBT9RkFTkyM3Z0m0itcAV0spT7mYWDW', prefix='매출_국내_29CM(주문)_', date='결제완료일시', name='상품명',
                 qty='수량', status='주문상태', key=('주문번호', '옵션코드'), opt='옵션'),
    '일본': dict(folder='1A9sBP0pEWBQZjLtDx96dWkJbR0jjwghk', prefix='매출_해외_일본_자사몰_API_', date='결제일시(입금확인일)',
                name='상품명(KR)', qty='수량', status='주문상태', key=('주문번호', '사용자 품목코드'), code='주문상태코드',
                sku='사용자 품목코드', opt='주문상품명(옵션포함)'),
    '이지체인': dict(folder='1lvY4MVWtBCz_x8GIilZUdVLE6ebvlVxb', prefix='매출_국내_이지체인(오프라인)_', date='구매일자', name='상품명',
                 qty='수량', status='판매/환불', key=('ID', '상품코드'), sku='상품코드', opt='옵션'),
}
EXCLUDE = ('취소', '반품', '환불', '입금전', '수거')
BACKUP = ('_자동백업_', '_백업_', '_bak_', '수동백업', '자동백업')
PERIOD = re.compile(r'(\d{6,8})_(\d{6,8})\.xlsx$')
SIZES = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'OS']
SIZE_RE = re.compile(r'(?<![A-Z])(XXL|2XL|XL|XS|OS|FREE|S|M|L)(?![A-Z])')
NO_COLOR, NO_SIZE = '기타(컬러 미기재)', '(사이즈 미기재)'


class ValidationError(Exception):
    pass


def log(*a):
    print(*a, file=sys.stderr, flush=True)


# ───────────────────────── Drive ─────────────────────────
def drive_session():
    from google.auth.transport.requests import AuthorizedSession
    scopes = ['https://www.googleapis.com/auth/drive.readonly']
    if os.environ.get('GDRIVE_SA_KEY'):
        from google.oauth2 import service_account
        cred = service_account.Credentials.from_service_account_info(json.loads(os.environ['GDRIVE_SA_KEY']), scopes=scopes)
    elif os.environ.get('GDRIVE_OAUTH_TOKEN') or os.environ.get('GDRIVE_OAUTH_TOKEN_FILE'):
        from google.oauth2.credentials import Credentials
        raw = os.environ.get('GDRIVE_OAUTH_TOKEN') or open(os.environ['GDRIVE_OAUTH_TOKEN_FILE'], encoding='utf-8-sig').read()
        raw = raw.strip().lstrip('﻿').strip()
        try:
            info = json.loads(raw)
        except json.JSONDecodeError:
            # 값은 출력하지 않고 형태만 알려준다
            raise SystemExit(f'GDRIVE_OAUTH_TOKEN 이 JSON이 아님 (길이 {len(raw)}, 첫 글자 {"{" if raw[:1] == "{" else "그 외"}) '
                             '— jh_authorized_user.json 파일 내용 전체를 그대로 붙여넣어야 함')
        cred = Credentials.from_authorized_user_info(info)  # 토큰에 저장된 범위 그대로 사용
    elif os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
        from google.oauth2 import service_account
        cred = service_account.Credentials.from_service_account_file(os.environ['GOOGLE_APPLICATION_CREDENTIALS'], scopes=scopes)
    else:
        raise SystemExit('Google 인증 정보 없음: GDRIVE_SA_KEY 또는 GDRIVE_OAUTH_TOKEN(_FILE) 필요')
    return AuthorizedSession(cred)


API = 'https://www.googleapis.com/drive/v3/files'


class Drive:
    def __init__(self):
        self.s = drive_session()

    def meta(self, fid):
        r = self.s.get(f'{API}/{fid}', params={'fields': 'id,name,mimeType,modifiedTime', 'supportsAllDrives': True})
        r.raise_for_status()
        return r.json()

    def get(self, m):
        if m['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            r = self.s.get(f"{API}/{m['id']}/export", params={'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
        else:
            r = self.s.get(f"{API}/{m['id']}", params={'alt': 'media', 'supportsAllDrives': True})
        r.raise_for_status()
        return io.BytesIO(r.content)

    def list_folder(self, folder):
        out, tok = [], None
        while True:
            p = {'q': f"'{folder}' in parents and trashed=false", 'fields': 'nextPageToken,files(id,name,mimeType,modifiedTime)',
                 'pageSize': 200, 'supportsAllDrives': True, 'includeItemsFromAllDrives': True}
            if tok:
                p['pageToken'] = tok
            r = self.s.get(API, params=p)
            r.raise_for_status()
            j = r.json()
            out += j.get('files', [])
            tok = j.get('nextPageToken')
            if not tok:
                return out


def kst(iso):
    return dt.datetime.fromisoformat(iso.replace('Z', '+00:00')).astimezone(KST).strftime('%Y-%m-%d %H:%M')


# ───────────────────────── 공통 ─────────────────────────
def match(name):
    s = str(name or '')
    for pat, canon in PATTERNS:
        if pat in s:
            return canon, pat
    return None, None


def nospace(s):
    return re.sub(r'\s+', '', str(s or ''))


def norm_code(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip().upper()
    return s.lstrip('0') if s.isdigit() else s


def norm_size(s):
    s = str(s or '').strip().upper()
    if s in ('', 'NONE'):
        return ''
    if s in ('F', 'FREE', 'ONE', 'ONESIZE', 'OS'):
        return 'OS'
    return {'2XL': 'XXL'}.get(s, s)


def rows_of(ws):
    it = ws.iter_rows(values_only=True)
    hdr = [str(h).strip() if h is not None else '' for h in next(it)]
    for r in it:
        if r is None or all(v is None for v in r):
            continue
        yield dict(zip(hdr, r))


def num(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


# ───────────────────────── 허브 ─────────────────────────
def read_hub(buf, today):
    ws = openpyxl.load_workbook(buf, data_only=True, read_only=True)['00_요약']
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=120, values_only=True)]
    col_a = [str(r[0]).strip() if r and r[0] is not None else '' for r in grid]

    def row(prefix, start=0):
        for i in range(start, len(col_a)):
            if col_a[i].startswith(prefix):
                return i
        raise ValidationError(f'허브 00_요약에서 "{prefix}" 행을 찾지 못함')

    i_t = row('시즌 목표매출')
    target, refreshed = num(grid[i_t][1]), str(grid[i_t][2] or '')
    cum = num(grid[row('누적 실제매출')][1])
    i_m = row('월', row('② 월별'))
    hdr = [str(v).strip() if v else '' for v in grid[i_m]]
    tgt_row, act_row = grid[i_m + 1], grid[i_m + 2]
    monthly = []
    for m in MONTHS:
        j = hdr.index(m)
        monthly.append({'month': m, 'target': num(tgt_row[j]), 'actual': num(act_row[j])})

    i_p = row('제품명', row('③ 제품별'))
    ph = [str(v).strip() if v else '' for v in grid[i_p]]
    c = {k: ph.index(k) for k in ('제품명', '히어로', '발주수량')}
    c_qty = next(j for j, h in enumerate(ph) if h.startswith('실제판매량'))
    c_rev = next(j for j, h in enumerate(ph) if h.startswith('실제매출'))
    hub_products = []
    for r in grid[i_p + 1:]:
        n = str(r[0] or '').strip()
        if not n or n == '합계':
            break
        canon, _ = match(n)
        hub_products.append({'name': canon or n, 'hero': 'Y' if str(r[c['히어로']] or '').strip().upper() == 'Y' else '',
                             'moq': int(num(r[c['발주수량']])), 'hub_qty': int(num(r[c_qty])), 'hub_rev': num(r[c_rev])})

    # 오늘 기준 목표 누적률 = 전월까지 목표 + 당월 목표 × 경과일/월일수
    acc = 0.0
    for k, m in enumerate(monthly):
        a, b = MONTH_START[k], MONTH_START[k + 1]
        if today >= b:
            acc += m['target']
        elif today >= a:
            acc += m['target'] * today.day / (b - a).days
    kpi = [{'as_of': today.isoformat(), 'season_target': target, 'cum_actual': cum,
            'season_progress': round(cum / target, 4) if target else 0, 'target_cum_rate': round(acc / target, 4) if target else 0}]
    return kpi, monthly, hub_products, refreshed


# ───────────────────────── 창고 재고 ─────────────────────────
def color_of(product_name, pat):
    s = str(product_name or '')
    s = re.sub(r'\[[^\]]*\]', ' ', s).replace('어반사이드', ' ')
    if pat and pat in s:
        s = s.split(pat, 1)[1]
    s = re.split(r'[(（]', s)[0]
    return re.sub(r'\s+', ' ', s).strip()


def read_inventory(buf):
    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    if INV_TAB not in wb.sheetnames:
        raise ValidationError(f'재고관리시트에 "{INV_TAB}" 탭 없음')
    it = wb[INV_TAB].iter_rows(values_only=True)
    hdr = list(next(it))
    ix = {str(h).strip(): i for i, h in enumerate(hdr) if isinstance(h, str)}
    # 첫 날짜 블록(17:00, '텐트 전용상품 여부' 앞) 의 마지막 컬럼. 열 문자 고정 금지.
    first = next(j for j, h in enumerate(hdr) if isinstance(h, dt.datetime))
    last = first
    while last + 1 < len(hdr) and isinstance(hdr[last + 1], dt.datetime):
        last += 1
    col_dt = hdr[last]
    if (col_dt.hour, col_dt.minute) != (17, 0):
        raise ValidationError(f'창고 재고 날짜 컬럼 시각이 17:00이 아님: {col_dt}')
    by_size = collections.Counter()
    sku = {}  # 코드 → (품목, 컬러, 사이즈)
    colors = collections.defaultdict(set)
    for r in it:
        if r is None or str(r[ix['시즌']] or '').strip() != '26FW' or str(r[ix['상품중분류']] or '').strip() != '어패럴':
            continue
        canon, pat = match(r[ix['라인']])
        if not canon:
            canon, pat = match(r[ix['상품명']])
        if not canon:
            log('  ! 품목 매칭 실패(창고):', r[ix['상품명']])
            continue
        _, pat_in_name = match(r[ix['상품명']])
        color = color_of(r[ix['상품명']], pat_in_name or pat) or NO_COLOR
        size = norm_size(r[ix['옵션']])
        opt_code = str(r[ix['공급처옵션']] or '')
        if not size:
            m = re.search(r'-(XS|S|M|L|XL|XXL|OS)$', opt_code)
            size = norm_size(m.group(1)) if m else NO_SIZE
        key = (canon, color, size)
        by_size[key] += int(num(r[last]))
        colors[canon].add(color)
        for code in (r[ix['공급처옵션']], r[ix['상품코드']], r[ix.get('바코드', ix['상품코드'])]):
            c = norm_code(code)
            if c:
                sku[c] = key
    return by_size, sku, colors, col_dt


# ───────────────────────── 팝업 ─────────────────────────
EN_COLOR = {'BLACK': '블랙', 'BLUE': '블루', 'BLUEBLACK': '블루 블랙', 'SILVER': '실버', 'CHARCOAL': '차콜', 'BEIGE': '베이지',
            'BROWN': '브라운', 'CAMEL': '카멜', 'ORANGE': '오렌지', 'LIGHTGRAY': '라이트 그레이', 'LIGHTGREY': '라이트 그레이',
            'DARKGRAY': '다크 그레이', 'DARKGREY': '다크 그레이', 'KHAKIGRAY': '카키 그레이', 'KHAKIGREY': '카키 그레이'}


def norm_color(raw, known):
    """팝업 등 다른 원본의 컬러 표기를 창고 재고 표기(한글)로 맞춘다."""
    s = str(raw or '').strip()
    if not s:
        return NO_COLOR
    s = EN_COLOR.get(nospace(s).upper(), s)
    for k in known:
        if nospace(k) == nospace(s):
            return k
    return s


def read_popup(buf, colors):
    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    stamp = ''
    for r in wb['대시보드'].iter_rows(min_row=1, max_row=5, values_only=True):
        if r and r[0] and '마지막 집계' in str(r[0]):
            stamp = str(r[0]).split('·')[0].replace('마지막 집계:', '').strip()
    stock = collections.Counter()
    by_store = collections.Counter()
    for d in rows_of(wb['재고현황']):
        if not d.get('지점'):
            continue
        canon, _ = match(d.get('상품명'))
        if not canon:
            log('  ! 품목 매칭 실패(팝업 재고):', d.get('상품명'))
            continue
        q = int(num(d.get('잔여재고')))
        stock[(canon, norm_color(d.get('컬러'), colors[canon]), norm_size(d.get('사이즈')) or NO_SIZE)] += q
        by_store[str(d['지점'])] += q
    sold_q, sold_rev, sold_cs = collections.Counter(), collections.Counter(), collections.Counter()
    for d in rows_of(wb['원장']):
        kind = str(d.get('구분') or '')
        if not any(k in kind for k in ('판매', '반품', '환불')):
            continue
        canon, _ = match(d.get('상품명'))
        if not canon:
            continue
        q = -int(num(d.get('수량(±)')))  # 판매는 음수로 기록됨
        amt = abs(num(d.get('판매금액')))
        if '오모테산도' in str(d.get('지점')):
            amt *= JPY_KRW
        sign = 1 if q >= 0 else -1
        sold_q[canon] += q
        sold_rev[canon] += sign * amt
        sold_cs[(canon, norm_color(d.get('컬러'), colors[canon]), norm_size(d.get('사이즈')) or NO_SIZE)] += q
    return stock, by_store, sold_q, sold_rev, sold_cs, stamp


# ───────────────────────── 주문 원본 ─────────────────────────
def ymd(s):
    s = s[-6:]
    return dt.date(2000 + int(s[:2]), int(s[2:4]), int(s[4:]))


def pick_files(files, prefix):
    """기간 파일 선택: 같은 매장·시작일이면 끝날이 가장 늦은(같으면 최신 수정) 파일, 다른 파일에 포함되는 기간은 제외."""
    cand = []
    for f in files:
        n = f['name']
        if not n.startswith(prefix) or any(b in n for b in BACKUP):
            continue
        m = PERIOD.search(n)
        if not m:
            continue
        a, b = ymd(m.group(1)), ymd(m.group(2))
        if b < SEASON_START:
            continue
        store = n[len(prefix):].rsplit('_', 2)[0] if n[len(prefix):].count('_') >= 2 else ''
        cand.append(dict(f, start=a, end=b, store=store))
    cand.sort(key=lambda f: (f['store'], f['start'], f['end'], f['modifiedTime']))
    keep = []
    for f in cand:
        if any(k['store'] == f['store'] and k['start'] <= f['start'] and f['end'] <= k['end'] for k in keep):
            continue
        keep = [k for k in keep if not (k['store'] == f['store'] and f['start'] <= k['start'] and k['end'] <= f['end'])]
        keep.append(f)
    return keep


SEASON_NAME = re.compile(r'\[?(\d{2})(?:FW|SS)\]?')
SEASON_CODE = re.compile(r'(?:^|[-A-Z])(\d{2})[FS](?=-)')


def other_season(name, code):
    """[25FW] 미스릴 등 같은 이름의 지난 시즌 상품 제외."""
    m = SEASON_NAME.search(str(name or ''))
    if m and m.group(1) != '26':
        return True
    m = SEASON_CODE.search(str(code or '').upper())
    return bool(m and m.group(1) != '26')


def parse_cs(ch, d, canon, pat, sku, colors):
    c = CHANNELS[ch]
    code = norm_code(d.get(c['sku'])) if c.get('sku') else ''
    opt = str(d.get(c.get('opt', ''), '') or '')
    size_txt = ''
    m = re.search(r'(?:사이즈|サイズ)\]?\s*=?\s*([A-Za-z0-9]+)', opt) or SIZE_RE.search(opt.split('/')[0].upper())
    if m:
        size_txt = norm_size(m.group(1))
    hit = None
    if code:
        hit = sku.get(code) or (sku.get(f'{code}-{size_txt}') if size_txt else None) or sku.get(f'{code}-OS')
    if hit and hit[0] == canon:
        return hit
    # 텍스트로 추정 (29CM 등)
    known = colors.get(canon, set())
    color = ''
    mc = re.search(r'\[색상\]\s*([^\[/,]+)', opt) or re.search(r'색상\s*=\s*([^,)]+)', opt)
    cands = [mc.group(1)] if mc else []
    cands.append(color_of(d.get(c['name']), pat))
    for cand in cands:
        for k in known:
            if nospace(k) and nospace(k) in nospace(cand):
                color = k
                break
        if color:
            break
    sizes = {s for (p, _, s) in sku.values() if p == canon}
    size = 'OS' if sizes == {'OS'} else (size_txt or NO_SIZE)
    return canon, color or NO_COLOR, size


def read_orders(drv, sku, colors, today):
    daily = collections.defaultdict(collections.Counter)
    by_ch_day = collections.defaultdict(collections.Counter)
    cs = collections.Counter()
    used = []
    for ch, c in CHANNELS.items():
        seen = set()
        for f in pick_files(drv.list_folder(c['folder']), c['prefix']):
            wb = openpyxl.load_workbook(drv.get(f), read_only=True, data_only=True)
            used.append({'channel': ch, 'file': f['name'], 'modified_kst': kst(f['modifiedTime']), 'start': f['start'].isoformat(),
                         'end': f['end'].isoformat()})
            for d in rows_of(wb.worksheets[0]):
                day = d.get(c['date'])
                if day in (None, ''):
                    continue
                day = day.date() if isinstance(day, dt.datetime) else dt.date.fromisoformat(str(day)[:10])
                if c.get('code') and str(d.get(c['code']) or '') == 'N2':  # 카페24 교환 재발송(원주문과 중복)
                    continue
                if any(x in str(d.get(c['status']) or '') for x in EXCLUDE):
                    continue
                key = tuple(str(d.get(k)) for k in c['key'])
                if key in seen:  # 주문번호+품목코드 중복 제거 (겹치는 기간 파일 포함)
                    continue
                seen.add(key)
                canon, pat = match(d.get(c['name']))
                if not canon or other_season(d.get(c['name']), d.get(c.get('sku', ''))):
                    continue
                q = int(num(d.get(c['qty'])))
                daily[day][canon] += q
                by_ch_day[day][ch] += q
                if day >= SEASON_START:
                    cs[parse_cs(ch, d, canon, pat, sku, colors)] += q
    days = [d for d in daily if d <= today]
    if not days:
        raise ValidationError('주문 원본에서 26FW 판매 행을 찾지 못함')
    D = max(days)
    win = [D - dt.timedelta(days=i) for i in (2, 1, 0)]
    tot, chs = collections.Counter(), collections.Counter()
    for d in win:
        tot.update(daily.get(d, {}))
        chs.update(by_ch_day.get(d, {}))
    recent = [{'name': n, 'recent3d_qty': int(tot.get(n, 0))} for n in NAMES]
    rmeta = {'start': win[0].isoformat(), 'end': D.isoformat(), 'total': sum(tot.values()), 'by_channel': dict(chs),
             'files': [u for u in used if u['end'] >= win[0].isoformat()]}
    sig = hashlib.sha256(json.dumps(sorted((u['file'], u['modified_kst']) for u in used), ensure_ascii=False).encode()).hexdigest()[:16]
    return recent, rmeta, cs, used, sig


# ───────────────────────── 암호화 ─────────────────────────
def _key(pw, salt):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())


def encrypt(obj, pw):
    salt, iv = os.urandom(16), os.urandom(12)
    ct = AESGCM(_key(pw, salt)).encrypt(iv, json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    return {'v': 1, 'kdf': 'PBKDF2-SHA256', 'iter': ITER, 'salt': b(salt), 'iv': b(iv), 'ct': b(ct)}


def decrypt(doc, pw):
    d = lambda k: base64.b64decode(doc[k])
    return json.loads(AESGCM(_key(pw, d('salt'))).decrypt(d('iv'), d('ct'), None))


def load_prev(pw):
    if not os.path.exists(OUT):
        return None
    try:
        return decrypt(json.load(open(OUT, encoding='utf-8')), pw)
    except Exception:
        log('  ! 이전 데이터 복호화 실패(비밀번호 변경?) — 직전 대비 검증은 건너뜀')
        return None


# ───────────────────────── 검증 ─────────────────────────
def validate(new, prev, notes):
    errs = []
    if len(new['products']) != 23:
        errs.append(f"품목 수 {len(new['products'])}개 (23개여야 함)")
    if {p['name'] for p in new['products']} != set(NAMES):
        errs.append('품목명이 23개 기준 목록과 다름: ' + ', '.join(sorted({p['name'] for p in new['products']} ^ set(NAMES))))
    inv = {r['name']: r for r in new['inventory']}
    sz_w, sz_p = collections.Counter(), collections.Counter()
    for r in new['inventory_by_size']:
        sz_w[r['product']] += r['available_qty']
        sz_p[r['product']] += r['popup_qty']
    for n, r in inv.items():
        if sz_w[n] != r['available_qty'] or sz_p[n] != r['popup_available_qty']:
            errs.append(f'{n}: 사이즈별 합계({sz_w[n]}/{sz_p[n]})가 품목 재고({r["available_qty"]}/{r["popup_available_qty"]})와 다름')
    wh = sum(r['available_qty'] for r in new['inventory'])
    pop = sum(r['popup_available_qty'] for r in new['inventory'])
    sold = sum(p['sold_qty'] for p in new['products'])
    notes.append(f'창고 재고 {wh:,} · 팝업 재고 {pop:,} · 누적 판매 {sold:,}')
    if prev:
        p_wh = sum(r['available_qty'] for r in prev['inventory'])
        p_pop = sum(r['popup_available_qty'] for r in prev['inventory'])
        p_sold = sum(p['sold_qty'] for p in prev['products'])
        if sold < p_sold:
            errs.append(f'누적 판매수량 감소 {p_sold:,} → {sold:,}')
        if p_sold and abs(sold - p_sold) / p_sold > 0.3:
            errs.append(f'누적 판매수량 ±30% 초과 변동 {p_sold:,} → {sold:,}')
        if p_wh and abs(wh - p_wh) / p_wh > 0.3:
            errs.append(f'창고 재고 ±30% 초과 변동 {p_wh:,} → {wh:,}')
        if pop != p_pop:
            notes.append(f'팝업 재고 변동 {p_pop:,} → {pop:,}')
    if errs:
        raise ValidationError('\n'.join(errs))


# ───────────────────────── main ─────────────────────────
def main():
    pw = os.environ.get('DASH_PASSWORD')
    if not pw:
        raise SystemExit('DASH_PASSWORD 없음')
    now = dt.datetime.now(KST)
    today = now.date()
    prev = load_prev(pw)
    drv = Drive()

    metas = {k: drv.meta(v) for k, v in SRC.items()}
    for k, m in metas.items():
        if any(b in m['name'] for b in BACKUP):
            raise ValidationError(f'{k} 원본이 백업 파일임: {m["name"]}')
        log(f'· {k}: {m["name"]} ({kst(m["modifiedTime"])})')

    kpi, monthly, hub_products, hub_refreshed = read_hub(drv.get(metas['hub']), today)
    wh_size, sku, colors, wh_col = read_inventory(drv.get(metas['inventory']))
    pop_stock, pop_store, pop_q, pop_rev, pop_cs, pop_stamp = read_popup(drv.get(metas['popup']), colors)
    recent, rmeta, ord_cs, used, sig = read_orders(drv, sku, colors, today)

    products = []
    for h in hub_products:
        n = h['name']
        q = h['hub_qty'] + pop_q.get(n, 0)
        products.append({'name': n, 'hero': h['hero'], 'moq': h['moq'], 'sold_qty': q, 'sold_rev': round(h['hub_rev'] + pop_rev.get(n, 0), 1),
                         'sell_through': round(q / h['moq'], 4) if h['moq'] else 0,
                         'popup_sold_qty': pop_q.get(n, 0), 'popup_sold_rev': round(pop_rev.get(n, 0), 1)})

    as_of_date = wh_col.date().isoformat()
    inventory, by_size = [], []
    for n in NAMES:
        inventory.append({'name': n, 'available_qty': sum(v for k, v in wh_size.items() if k[0] == n),
                          'popup_available_qty': sum(v for k, v in pop_stock.items() if k[0] == n), 'as_of_date': as_of_date})
    for key in sorted(set(wh_size) | set(pop_stock)):
        by_size.append({'product': key[0], 'color': key[1], 'size': key[2],
                        'available_qty': wh_size.get(key, 0), 'popup_qty': pop_stock.get(key, 0)})

    cs = ord_cs + pop_cs
    color_size = [{'product': k[0], 'color': k[1], 'size': k[2], 'qty': v} for k, v in sorted(cs.items()) if v]

    meta = {
        'built_at_kst': now.strftime('%Y-%m-%d %H:%M'),
        'sources': [
            {'label': '재무모델 허브 (시즌 KPI·월별·품목 판매)', 'file': metas['hub']['name'], 'modified_kst': kst(metas['hub']['modifiedTime']),
             'note': hub_refreshed.strip('()')},
            {'label': '창고 재고 (재고관리시트)', 'file': metas['inventory']['name'], 'modified_kst': kst(metas['inventory']['modifiedTime']),
             'note': f"{INV_TAB} · {wh_col.strftime('%Y-%m-%d %H:%M')} 컬럼"},
            {'label': '팝업 재고·판매', 'file': metas['popup']['name'], 'modified_kst': kst(metas['popup']['modifiedTime']),
             'note': f"마지막 집계 {pop_stamp} · " + ' / '.join(f'{k} {v}' for k, v in pop_store.items())},
            {'label': '최근 3일 판매 (주문 원본)', 'file': ', '.join(f['file'] for f in rmeta['files']),
             'modified_kst': max((f['modified_kst'] for f in rmeta['files']), default=''),
             'note': f"결제일 {rmeta['start']} ~ {rmeta['end']} · " + ' · '.join(f'{k} {v}' for k, v in rmeta['by_channel'].items())},
            {'label': '컬러·사이즈 판매 (주문 원본 + 팝업)', 'file': f'주문 파일 {len(used)}개 (무신사 제외)', 'modified_kst': '',
             'note': f'{SEASON_START.isoformat()} 이후 결제분'},
        ],
        'warehouse_col': wh_col.strftime('%Y-%m-%d %H:%M'),
        'recent3d': {'start': rmeta['start'], 'end': rmeta['end']},
        'signature': {'hub': metas['hub']['modifiedTime'], 'inventory': metas['inventory']['modifiedTime'],
                      'popup': metas['popup']['modifiedTime'], 'orders': sig},
    }
    new = {'season_kpi': kpi, 'monthly': monthly, 'products': products, 'inventory': inventory,
           'inventory_by_size': by_size, 'color_size': color_size, 'recent3d': recent, 'meta': meta}

    notes = []
    validate(new, prev, notes)
    for n in notes:
        log('· ' + n)

    strip = lambda o: {k: v for k, v in o.items() if k != 'meta'} | {'meta': {k: v for k, v in o['meta'].items() if k != 'built_at_kst'}}
    changed = prev is None or strip(prev) != strip(new)
    if changed:
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, 'w', encoding='utf-8') as fh:
            json.dump(encrypt(new, pw), fh)
            fh.write('\n')
        log('· 저장:', os.path.relpath(OUT, ROOT))
    else:
        log('· 원본·결과 변화 없음 — 기존 데이터 유지')
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a', encoding='utf-8') as fh:
            fh.write(f"changed={'true' if changed else 'false'}\nas_of={meta['built_at_kst']}\n")
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as fh:
            fh.write('### 26FW 소진율 빌드\n' + '\n'.join(f'- {n}' for n in notes) + f"\n- 변경: {'있음' if changed else '없음'}\n")


if __name__ == '__main__':
    try:
        main()
    except ValidationError as e:
        log('✖ 검증 실패 — 데이터를 쓰지 않음\n' + str(e))
        if os.environ.get('GITHUB_STEP_SUMMARY'):
            with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as fh:
                fh.write('### ✖ 26FW 소진율 검증 실패\n```\n' + str(e) + '\n```\n')
        sys.exit(2)
