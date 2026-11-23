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
from html.parser import HTMLParser

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
    'price': '1elpoQQ3CW4aNpcZfFJbu8q3ShUGKhNoZT2p3dGw4MFA',   # [어반사이드_어패럴] Price table — 마켓 기간·가격
}
MARKET_KEYS = ('JSUN', '제이선', '네고무카', '무카', '4910', '고윈')
INV_TAB = '일자 별 재고조회 (당시재고)'
JPY_KRW = 9  # 오모테산도 엔화 판매 환산 (¥1=9원, 라인시트 기준)
SEASON_START = dt.date(2026, 8, 1)  # 허브와 같은 기준 (7월 선런칭분은 허브에도 없음)
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
MUSINSA = dict(folder='1fx0hwqs_6lpVLn5DUu_jYLZQ3SWlZYPg', prefix='매출_국내_무신사_어패럴_')
# CM2 기본값 — 허브 01_입력값에서 읽어 덮어쓴다 (29CM·아난티는 허브에 없어 고정 가정)
FEE_DEFAULT = {'자사몰': 0.035, '무신사': 0.275, '4910': 0.15, '일본': 0.224, '이지체인': 0.30, '29CM': 0.30,
               '팝업 아난티 코브': 0.30, '팝업 더현대 충청': 0.30, '팝업 더현대 오모테산도': 0.60}
SHIP_KR, SHIP_JP, RESERVE = 6330, 20000, 0.03
DOMESTIC = {'자사몰', '무신사', '4910', '29CM', '이지체인', '팝업 아난티 코브', '팝업 더현대 충청'}
ONLINE_KR = {'자사몰', '무신사', '4910', '29CM'}
BACKUP = ('_자동백업_', '_백업_', '_bak_', '수동백업', '자동백업')
PERIOD = re.compile(r'(\d{6,8})_(\d{6,8})\.xlsx?$')
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
    """허브는 스크립트가 수식만 쓰고 계산 결과를 저장하지 않을 때가 있어(00_요약이 빈칸),
    수식이 참조하는 입력값에서 직접 합산한다.
      02_제품마스터  제품명·히어로·발주수량·판매가(VAT포함)·원가(VAT제외)   ← 값
      04_월별계산기  상세 계산표 23품목×7개월 행의 월·실제판매량(AE)·실제매출(AH) ← 값 (제품명 칸은 수식)
      01_입력값      ⑤ 월별 목표매출                                           ← 값
    """
    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    for t in ('02_제품마스터', '04_월별계산기_및_매출집계', '01_입력값'):
        if t not in wb.sheetnames:
            raise ValidationError(f'허브에 "{t}" 탭 없음')

    # 제품마스터 (행 순서 = 04 상세 계산표의 품목 순서)
    master, hdr = [], None
    for r in wb['02_제품마스터'].iter_rows(values_only=True):
        cells = [str(v).strip() if v is not None else '' for v in r]
        if hdr is None:
            if '제품명' in cells and any(c.startswith('발주수량') for c in cells):
                hdr = cells
            continue
        n = cells[hdr.index('제품명')] if len(cells) > hdr.index('제품명') else ''
        if not n:
            if master:
                break
            continue
        canon, _ = match(n)
        col = lambda pre: next(i for i, h in enumerate(hdr) if h.startswith(pre))
        master.append({'name': canon or n, 'hero': 'Y' if cells[col('히어로')].upper() == 'Y' else '',
                       'moq': int(num(r[col('발주수량')])), 'tag': num(r[col('판매가')]), 'cost': num(r[col('원가(VAT제외)')])})
    if len(master) != 23:
        raise ValidationError(f'허브 02_제품마스터 품목 {len(master)}개 (23개여야 함)')

    # 04 상세 계산표: 품목마다 8월~2월 7행
    rows = list(wb['04_월별계산기_및_매출집계'].iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows) if r and str(r[0] or '').strip() == '제품명'
               and '월' in [str(v or '').strip() for v in r] and any(str(v or '').startswith('실제판매량') for v in r)), None)
    if hi is None:
        raise ValidationError('허브 04 상세 계산표 머리글을 찾지 못함')
    h = [str(v or '').strip() for v in rows[hi]]
    c_m = h.index('월')
    c_q = next(i for i, x in enumerate(h) if x.startswith('실제판매량'))
    c_r = next(i for i, x in enumerate(h) if x.startswith('실제매출'))
    detail = [r for r in rows[hi + 1:hi + 1 + 23 * 7]]
    if [str(r[c_m] or '').strip() if c_m < len(r) else '' for r in detail] != MONTHS * 23:
        raise ValidationError('허브 04 상세 계산표가 23품목×7개월 순서가 아님')
    act = collections.Counter()
    for i, p in enumerate(master):
        blk = detail[i * 7:(i + 1) * 7]
        cell = lambda r, j: r[j] if j < len(r) else None
        p['hub_qty'] = int(sum(num(cell(r, c_q)) for r in blk))
        p['hub_rev'] = sum(num(cell(r, c_r)) for r in blk)
        for r in blk:
            act[str(r[c_m]).strip()] += num(cell(r, c_r))
    hub_products = master

    # 월 목표 (01_입력값 '목표매출(원)' 행, 머리글 '월')
    tgt = {}
    rr = list(wb['01_입력값'].iter_rows(values_only=True))
    for i, r in enumerate(rr):
        if r and str(r[0] or '').strip() == '목표매출(원)' and i > 0 and str(rr[i - 1][0] or '').strip() == '월':
            mh = [str(v or '').strip() for v in rr[i - 1]]
            tgt = {m: num(r[mh.index(m)]) for m in MONTHS if m in mh}
            break
    if len(tgt) != 7:
        raise ValidationError('허브 01_입력값 월별 목표매출을 찾지 못함')
    monthly = [{'month': m, 'target': tgt[m], 'actual': act[m]} for m in MONTHS]
    target, cum = sum(tgt.values()), sum(act.values())

    refreshed = ''
    for r in wb['00_요약'].iter_rows(min_row=1, max_row=15, values_only=True) if '00_요약' in wb.sheetnames else []:
        for v in r or ():
            if v and '매출 최신화' in str(v):
                refreshed = str(v)

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

    # 취소·반품 매출 (허브 00_취소매출현황: 자사몰·4910·29CM·일본 자사몰만 집계, 무신사·이지체인은 구조상 미집계)
    cancels = []
    if '00_취소매출현황' in wb.sheetnames:
        cg = [list(r) for r in wb['00_취소매출현황'].iter_rows(min_row=1, max_row=40, values_only=True)]
        hi = next((i for i, r in enumerate(cg) if r and str(r[0] or '').strip() == '채널'), None)
        if hi is not None:
            ch = [str(v).strip() if v else '' for v in cg[hi]]
            for r in cg[hi + 1:]:
                name = str(r[0] or '').strip()
                if not name or name == '합계' or name.startswith('('):
                    break
                for m in MONTHS:
                    if m in ch:
                        cancels.append({'channel': name, 'month': m, 'amount': num(r[ch.index(m)])})
    # 채널수수료·운송비: 01_입력값
    fees, ship = dict(FEE_DEFAULT), SHIP_KR
    if '01_입력값' in wb.sheetnames:
        alias = {'자사몰': ['자사몰'], '무신사': ['무신사'], '4910': ['4910'], '일본몰': ['일본'],
                 '현대백화점': ['이지체인', '팝업 더현대 충청'], '일본 팝업': ['팝업 더현대 오모테산도']}
        for r in wb['01_입력값'].iter_rows(values_only=True):
            a = str(r[0] or '').strip()
            if a == '국내운송비' and num(r[1]):
                ship = num(r[1])
            for k, targets in alias.items():
                if a.startswith(k) and len(r) > 3 and isinstance(r[3], (int, float)) and 0 < r[3] < 1:
                    for t in targets:
                        fees[t] = float(r[3])
    return kpi, monthly, hub_products, refreshed, cancels, fees, ship


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
    pending = collections.Counter()  # '접수' = 결제됐지만 아직 출고 안 된 수량 (판매에도 잡혀 있음)
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
        if '접수' in ix:
            pending[key] += int(num(r[ix['접수']]))
        colors[canon].add(color)
        for code in (r[ix['공급처옵션']], r[ix['상품코드']], r[ix.get('바코드', ix['상품코드'])]):
            c = norm_code(code)
            if c:
                sku[c] = key
    sizes = collections.defaultdict(set)
    for (p, _, z) in by_size:
        sizes[p].add(z)
    one_size = {p for p, zs in sizes.items() if zs == {'OS'}}
    return by_size, pending, sku, colors, one_size, col_dt


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
    s = s.replace('챠콜', '차콜')
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
    sold_q, sold_rev = collections.Counter(), collections.Counter()
    recs = []  # 판매 기록 (주문 원본과 같은 형태로 합친다)
    for d in rows_of(wb['원장']):
        kind = str(d.get('구분') or '')
        if not any(k in kind for k in ('판매', '반품', '환불')):
            continue
        canon, _ = match(d.get('상품명'))
        if not canon:
            continue
        q = -int(num(d.get('수량(±)')))  # 판매는 음수로 기록됨
        amt = abs(num(d.get('판매금액')))
        store = str(d.get('지점') or '').strip()
        if '오모테산도' in store:
            amt *= JPY_KRW
        sign = 1 if q >= 0 else -1
        sold_q[canon] += q
        sold_rev[canon] += sign * amt
        day = d.get('일자')
        day = day.date() if isinstance(day, dt.datetime) else None
        if day:
            recs.append({'ch': f'팝업 {store}', 'day': day, 'qty': q, 'paid': sign * amt,
                         'key': (canon, norm_color(d.get('컬러'), colors[canon]), norm_size(d.get('사이즈')) or NO_SIZE)})
    return stock, by_store, sold_q, sold_rev, recs, stamp


# ───────────────────────── 마켓(외부 제휴) 기간 ─────────────────────────
def read_markets(buf):
    """Price table [국내_온라인] 26FW 행 중 외부 마켓(제이선·네고무카·4910·고윈) 구간.
    → (품목별 [(시작, 종료+1일, 마켓가)], 4910 마켓 구간 목록)"""
    wb = openpyxl.load_workbook(buf, data_only=True, read_only=True)
    ws = next((w for w in wb.worksheets if '국내' in w.title and 'PriceTable' in w.title), wb.worksheets[0])
    by_item, w4910 = collections.defaultdict(set), set()
    for r in ws.iter_rows(values_only=True):
        if not (r and len(r) > 11 and r[1] and '26FW' in str(r[1]) and r[3]):
            continue
        name = str(r[7] or '').strip()
        if not any(k in name for k in MARKET_KEYS):
            continue
        s, e = to_date(r[5]), to_date(r[6])
        try:
            price = int(float(r[11]))
        except (TypeError, ValueError):
            continue
        if not (s and e):
            continue
        if '4910' in name:
            w4910.add((s, e))
            continue
        canon, _ = match(r[3])
        if canon:
            by_item[canon].add((s, e + dt.timedelta(days=1), price))
    return by_item, sorted(w4910)


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


# 주문 상태 분류 (판매로 세지 않는 행 중 재고·비용에 영향을 주는 것)
RET_OPEN = ('반품신청', '반품처리중', '반품접수', '반품 접수', '반품진행', '반품 진행', '반품요청', '반품 요청', '수거중', '수거 중', '수거완료')
RET_DONE = ('반품완료', '반품 완료')


def classify(status):
    """→ 'sale' | 'exchange' | 'return_open' | 'return_done' | 'skip'"""
    s = str(status or '')
    if any(k in s for k in RET_OPEN):
        return 'return_open'
    if any(k in s for k in RET_DONE):
        return 'return_done'
    if any(k in s for k in EXCLUDE):
        return 'skip'
    if '교환' in s:
        return 'exchange'
    return 'sale'


def to_date(v):
    if v in (None, ''):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    s = str(v).strip()
    if re.fullmatch(r'\d{8}', s):
        return dt.date(int(s[:4]), int(s[4:6]), int(s[6:]))
    try:
        return dt.date.fromisoformat(s[:10])
    except ValueError:
        return None


def cafe24_paid(rows, jpy):
    """카페24(자사몰·일본몰): 쿠폰·적립금·할인코드는 주문의 첫 줄에만 적혀 있어 주문 안 상품들에 금액 비율로 나눈다."""
    amt_col = '상품구매금액(JPY)' if jpy else '상품구매금액'
    disc_cols = (('쿠폰 할인금액(JPY)', '사용한 적립금액(JPY)', '할인코드 할인금액') if jpy
                 else ('쿠폰 할인금액(최종)', '사용한 적립금액(최종)', '할인코드 할인금액'))
    item_disc = '상품별 추가할인금액(JPY)' if jpy else '상품별 추가할인금액'
    gross, disc = collections.Counter(), collections.Counter()
    for d in rows:
        o = str(d.get('주문번호'))
        gross[o] += num(d.get(amt_col))
        disc[o] += sum(num(d.get(k)) for k in disc_cols)
    out = []
    for d in rows:
        o = str(d.get('주문번호'))
        g = num(d.get(amt_col))
        share = disc[o] * g / gross[o] if gross[o] else 0
        paid = max(0.0, g - share - num(d.get(item_disc)))
        out.append(paid * (JPY_KRW if jpy else 1))
    return out


class HtmlTable(HTMLParser):
    """무신사 정산 파일(.xls 확장자지만 실제로는 HTML 표)."""
    def __init__(self):
        super().__init__()
        self.rows, self.cur, self.cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.cur = []
        elif tag in ('td', 'th') and self.cur is not None:
            self.cell = ''

    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self.cell is not None:
            self.cur.append(self.cell.strip())
            self.cell = None
        elif tag == 'tr' and self.cur is not None:
            self.rows.append(self.cur)
            self.cur = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell += data


def read_musinsa(drv, colors, one_size, add, used):
    seen = set()
    for f in pick_files(drv.list_folder(MUSINSA['folder']), MUSINSA['prefix']):
        raw = drv.get(f).getvalue()
        txt = None
        for enc in ('utf-8', 'cp949', 'euc-kr'):
            try:
                txt = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if txt is None:
            raise ValidationError(f'무신사 파일 인코딩을 읽지 못함: {f["name"]}')
        t = HtmlTable()
        t.feed(txt)
        if not t.rows:
            continue
        used.append({'channel': '무신사', 'file': f['name'], 'modified_kst': kst(f['modifiedTime']), 'start': f['start'].isoformat(),
                     'end': f['end'].isoformat()})
        hdr = t.rows[0]
        ix = {}
        for i, h in enumerate(hdr):
            ix.setdefault(h, i)          # 같은 이름(과세·소계)은 첫 번째 열
        i_settle = max(i for i, h in enumerate(hdr) if h == '정산금액')

        def g(r, k):
            return r[ix[k]] if k in ix and ix[k] < len(r) else ''

        for r in t.rows[1:]:
            if len(r) < len(hdr) - 2 or not g(r, '상품'):
                continue
            kind_txt, claim = g(r, '구분'), g(r, '클레임상태')
            if not kind_txt.startswith('판매'):
                continue
            key = g(r, '주문일련번호')
            if key in seen:
                continue
            seen.add(key)
            name = g(r, '상품')
            canon, _ = match(name)
            if not canon or other_season(name, g(r, '스타일넘버')):
                continue
            day = to_date(g(r, '일자'))
            if not day:
                continue
            opt = re.split(r'\s*[:^]\s*', g(r, '옵션'))
            color = norm_color(opt[0] if opt else '', colors[canon])
            size = 'OS' if canon in one_size else (norm_size(opt[1]) if len(opt) > 1 else '') or NO_SIZE
            sku_key = (canon, color, size)
            q = int(num(g(r, '수량'))) or 1
            if '환불' in kind_txt or '환불' in claim:
                if '요청' in claim:
                    add('무신사', day, sku_key, q, 'return_open')
                elif '출고전' not in kind_txt:
                    add('무신사', day, sku_key, q, 'return_done')
                continue
            paid = (num(g(r, '판매금액')) - num(g(r, '할인금액')) - num(g(r, '소계')) - num(g(r, '적립금'))
                    - num(g(r, '장바구니할인')) - num(g(r, '그룹할인')) - num(g(r, '추가할인')))
            settle = num(r[i_settle]) if i_settle < len(r) else 0
            kind = 'exchange' if '교환' in kind_txt or '교환' in claim else 'sale'
            # 무신사는 정산금액이 있어 실제 수수료(쿠폰 부담분 포함) = 실결제 − 정산
            add('무신사', day, sku_key, q, kind, paid=paid, fee=max(0.0, paid - settle) / 1.1)


def read_orders(drv, sku, colors, one_size, popup_recs, fees, markets=(None, ())):
    """주문 원본(5개 채널 + 무신사) + 팝업 판매를 행 단위 기록으로 모은다."""
    recs = []           # 판매·교환: dict(ch, day, key, qty, paid, fee)
    events = collections.defaultdict(collections.Counter)   # 'exchange'|'return_open'|'return_done' → (품목,컬러,사이즈) → 수량
    used = []

    mk_item, mk_4910 = markets[0] or {}, markets[1]

    def add(ch, day, key, q, kind, paid=0.0, fee=None, mkt=False):
        if day < SEASON_START:
            return
        if kind in ('sale', 'exchange'):
            recs.append({'ch': ch, 'day': day, 'key': key, 'qty': q, 'paid': paid, 'mkt': mkt,
                         'fee': fee if fee is not None else paid / 1.1 * fees.get(ch, 0.0)})
        if kind != 'sale':
            events[kind][key] += q

    for ch, c in CHANNELS.items():
        seen = set()
        for f in pick_files(drv.list_folder(c['folder']), c['prefix']):
            wb = openpyxl.load_workbook(drv.get(f), read_only=True, data_only=True)
            used.append({'channel': ch, 'file': f['name'], 'modified_kst': kst(f['modifiedTime']), 'start': f['start'].isoformat(),
                         'end': f['end'].isoformat()})
            rows = list(rows_of(wb.worksheets[0]))
            fee_base = None
            if ch in ('자사몰', '일본'):
                paids = cafe24_paid(rows, jpy=(ch == '일본'))
            elif ch == '4910':
                paids = [num(d.get('결제액')) for d in rows]
                # 4910 플랫폼 부담 쿠폰·적립금은 우리 정산에 들어오므로 수수료 기준에 더한다
                fee_base = [num(d.get('결제액')) + num(d.get('쿠폰지원금')) + num(d.get('적립금지원금')) + num(d.get('결제수단지원금'))
                            for d in rows]
            elif ch == '29CM':
                paids = [num(d.get('실 판매액')) for d in rows]
            else:  # 이지체인
                paids = [num(d.get('실판매금액')) for d in rows]
            for i, d in enumerate(rows):
                day = to_date(d.get(c['date']))
                if not day:
                    continue
                if c.get('code') and str(d.get(c['code']) or '') == 'N2':  # 카페24 교환 재발송(원주문과 중복)
                    continue
                key = tuple(str(d.get(k)) for k in c['key'])
                if key in seen:  # 주문번호+품목코드 중복 제거 (겹치는 기간 파일 포함)
                    continue
                seen.add(key)
                canon, pat = match(d.get(c['name']))
                if not canon or other_season(d.get(c['name']), d.get(c.get('sku', ''))):
                    continue
                kind = classify(d.get(c['status']))
                if kind == 'skip':
                    continue
                q = int(num(d.get(c['qty'])))
                fee = fee_base[i] / 1.1 * fees.get(ch, 0.0) if fee_base else None
                # 외부 마켓 판매(평시 속도에서 뺄 '거품'): 자사몰은 마켓 기간에 마켓가로 결제된 줄, 4910은 4910 마켓 기간 전부
                mkt = False
                if ch == '자사몰' and q:
                    unit = num(d.get('상품구매금액')) / q
                    mkt = any(s <= day <= e and abs(unit - pr) <= 1000 for s, e, pr in mk_item.get(canon, ()))
                elif ch == '4910':
                    mkt = any(s <= day <= e for s, e in mk_4910)
                add(ch, day, parse_cs(ch, d, canon, pat, sku, colors), q, kind, paid=paids[i], fee=fee, mkt=mkt)
    read_musinsa(drv, colors, one_size, add, used)
    for r in popup_recs:
        add(r['ch'], r['day'], r['key'], r['qty'], 'sale', paid=r['paid'])
    if not recs:
        raise ValidationError('주문 원본에서 26FW 판매 행을 찾지 못함')
    sig = hashlib.sha256(json.dumps(sorted((u['file'], u['modified_kst']) for u in used), ensure_ascii=False).encode()).hexdigest()[:16]
    return recs, events, used, sig


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
    for r in new['inventory']:
        if r['real_qty'] != r['available_qty'] - r['pending_qty'] + r['returning_qty'] + r['popup_available_qty']:
            errs.append(f"{r['name']}: 실재고 계산 불일치")
    wh = sum(r['available_qty'] for r in new['inventory'])
    pop = sum(r['popup_available_qty'] for r in new['inventory'])
    sold = sum(p['sold_qty'] for p in new['products'])
    notes.append(f'창고 재고 {wh:,} · 팝업 재고 {pop:,} · 누적 판매 {sold:,}')
    if prev:
        p_wh = sum(r['available_qty'] for r in prev['inventory'])
        p_pop = sum(r['popup_available_qty'] for r in prev['inventory'])
        p_sold = sum(p['sold_qty'] for p in prev['products'])
        # 취소가 사후 반영되면 누적이 조금 줄 수 있다 — 20개 또는 1% 넘게 줄 때만 이상으로 본다
        if p_sold - sold > max(20, p_sold * 0.01):
            errs.append(f'누적 판매수량 감소 {p_sold:,} → {sold:,}')
        elif sold < p_sold:
            notes.append(f'누적 판매수량 소폭 감소 {p_sold:,} → {sold:,} (취소 반영으로 보고 허용)')
        if p_sold and abs(sold - p_sold) / p_sold > 0.3:
            errs.append(f'누적 판매수량 ±30% 초과 변동 {p_sold:,} → {sold:,}')
        if p_wh and abs(wh - p_wh) / p_wh > 0.3:
            errs.append(f'창고 재고 ±30% 초과 변동 {p_wh:,} → {wh:,}')
        if pop != p_pop:
            notes.append(f'팝업 재고 변동 {p_pop:,} → {pop:,}')
    if errs:
        raise ValidationError('\n'.join(errs))


# ───────────────────────── main ─────────────────────────
def make_insights(D, win3, r3, p3, r3m, r3ch, inv, prod):
    """최근 3일 판매(팝업 제외)를 직전 3일과 비교해 상단에 띄울 한 줄 요약 몇 개를 만든다."""
    short = lambda n: n.replace('어반사이드 ', '')
    chg = lambda a, b: f'{(a / b - 1) * 100:+.0f}%' if b else '신규'
    tot, ptot, mtot = sum(r3.values()), sum(p3.values()), sum(r3m.values())
    lines = []
    trend = '직전 3일과 비슷' if ptot and abs(tot / ptot - 1) < .05 else f'직전 3일 대비 {chg(tot, ptot)}'
    head = f'최근 3일 {tot:,}개 판매 · {trend}'
    if mtot and mtot / tot >= .3:
        head += f' (외부 마켓 {mtot:,}개, {mtot / tot:.0%} 포함)'
    top = r3.most_common(3)
    if top:
        lines.append('많이 팔린 품목: ' + ', '.join(f'{short(n)} {q}개({chg(q, p3[n])})' for n, q in top))
    ups = sorted(((q - p3[n], n) for n, q in r3.items() if q >= 10 and q >= 2 * max(p3[n], 1)), reverse=True)[:2]
    if ups:
        lines.append('급증: ' + ', '.join(f'{short(n)} {p3[n]}→{r3[n]}개' + (' (마켓 판매 위주)' if r3m[n] / r3[n] >= .5 else '')
                                          for _, n in ups))
    downs = sorted(((p3[n] - r3[n], n) for n in p3 if p3[n] >= 10 and r3[n] <= p3[n] / 2), reverse=True)[:2]
    if downs:
        lines.append('주춤: ' + ', '.join(f'{short(n)} {p3[n]}→{r3[n]}개' for _, n in downs))
    risk = []
    for n, q in r3.items():
        organic = q - r3m[n]
        real = (inv.get(n) or {}).get('real_qty', 0)
        if organic >= 3 and real:
            days = real / (organic / 3)
            if days <= 45:
                risk.append((days, n, real))
    if risk:
        lines.append('이 속도면 곧 품절: ' + ', '.join(f'{short(n)} 약 {d:.0f}일분(실재고 {real}개)' for d, n, real in sorted(risk)[:3]))
    if r3ch:
        ch, q = max(r3ch.items(), key=lambda x: x[1])
        lines.append(f'채널: {ch} {q / tot:.0%}' + ''.join(f' · {c} {v / tot:.0%}' for c, v in sorted(r3ch.items(), key=lambda x: -x[1])[1:3]))
    return {'start': min(win3).isoformat(), 'end': D.isoformat(), 'headline': head, 'lines': lines}


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

    kpi, monthly, hub_products, hub_refreshed, cancels, fees, ship_kr = read_hub(drv.get(metas['hub']), today)
    wh_size, pending, sku, colors, one_size, wh_col = read_inventory(drv.get(metas['inventory']))
    pop_stock, pop_store, pop_q, pop_rev, pop_recs, pop_stamp = read_popup(drv.get(metas['popup']), colors)
    markets = read_markets(drv.get(metas['price']))
    recs, events, used, sig = read_orders(drv, sku, colors, one_size, pop_recs, fees, markets)

    # ── 판매 기록에서 파생 ──
    online = [r for r in recs if not r['ch'].startswith('팝업')]
    D = max(r['day'] for r in online if r['day'] <= today)
    win3 = {D - dt.timedelta(days=i) for i in range(3)}
    win14 = {D - dt.timedelta(days=i) for i in range(14)}
    prev3 = {D - dt.timedelta(days=i) for i in range(3, 6)}
    cs, cs14, r3, r3ch, mkq = (collections.Counter() for _ in range(5))
    p3, r3m = collections.Counter(), collections.Counter()   # 직전 3일 판매, 최근 3일 중 외부 마켓분
    mix = collections.Counter()
    agg = collections.defaultdict(lambda: collections.Counter())   # 품목 → 금액·수량 합계
    weekly = collections.defaultdict(lambda: collections.Counter())  # (품목, 주 시작일) → paid, qty
    for r in recs:
        n, q = r['key'][0], r['qty']
        cs[r['key']] += q
        if r['day'] in win14 and not r['mkt']:   # 품절 임박 판단용: 마켓 거품 제외 평시 속도
            cs14[r['key']] += q
        if r['mkt']:
            mkq[n] += q
        if r['day'] in win3 and not r['ch'].startswith('팝업'):
            r3[n] += q
            r3ch[r['ch']] += q
            if r['mkt']:
                r3m[n] += q
        if r['day'] in prev3 and not r['ch'].startswith('팝업'):
            p3[n] += q
        mix[(n, '팝업' if r['ch'].startswith('팝업') else r['ch'])] += q
        a = agg[n]
        a['qty'] += q
        a['net'] += r['paid'] / 1.1
        a['fee'] += r['fee']
        a['ship'] += q * (ship_kr if r['ch'] in ONLINE_KR else SHIP_JP if r['ch'] == '일본' else 0)
        if r['ch'] in DOMESTIC:   # 객단가·할인율·가격 추이는 원화 정가와 비교 가능한 국내 판매만
            a['dom_qty'] += q
            a['dom_paid'] += r['paid']
            wk = r['day'] - dt.timedelta(days=r['day'].weekday())
            weekly[(n, wk)]['paid'] += r['paid']
            weekly[(n, wk)]['qty'] += q
    ret_cost = collections.Counter()
    for kind in ('return_done', 'exchange'):
        for k, q in events[kind].items():
            ret_cost[k[0]] += q * 2 * ship_kr   # 반품·교환 왕복 운송비

    products, tot = [], collections.Counter()
    for h in hub_products:
        n = h['name']
        q = h['hub_qty'] + pop_q.get(n, 0)
        a = agg[n]
        cogs = a['qty'] * h['cost']
        reserve = a['net'] * RESERVE
        cm2 = a['net'] - a['fee'] - cogs - a['ship'] - reserve - ret_cost[n]
        avg_price = a['dom_paid'] / a['dom_qty'] if a['dom_qty'] else 0
        for k, v in (('net', a['net']), ('cm2', cm2), ('dom_paid', a['dom_paid']), ('dom_tag', a['dom_qty'] * h['tag'])):
            tot[k] += v
        products.append({'name': n, 'hero': h['hero'], 'moq': h['moq'], 'sold_qty': q, 'sold_rev': round(h['hub_rev'] + pop_rev.get(n, 0), 1),
                         'sell_through': round(q / h['moq'], 4) if h['moq'] else 0,
                         'popup_sold_qty': pop_q.get(n, 0), 'popup_sold_rev': round(pop_rev.get(n, 0), 1),
                         'order_qty': int(a['qty']), 'market_qty': int(mkq[n]), 'tag_price': h['tag'], 'avg_price': round(avg_price),
                         'discount_rate': round(1 - avg_price / h['tag'], 4) if h['tag'] and avg_price else None,
                         'cm2': round(cm2), 'cm2_rate': round(cm2 / a['net'], 4) if a['net'] else None})
    price_kpi = {'discount_rate': round(1 - tot['dom_paid'] / tot['dom_tag'], 4) if tot['dom_tag'] else None,
                 'cm2': round(tot['cm2']), 'net_rev': round(tot['net']),
                 'cm2_rate': round(tot['cm2'] / tot['net'], 4) if tot['net'] else None,
                 'assumptions': {'fees': fees, 'ship_kr': ship_kr, 'ship_jp': SHIP_JP, 'reserve': RESERVE}}

    # ── 재고: 창고·팝업은 따로, 실재고 = 창고 − 출고대기 + 회수중 반품 + 팝업 ──
    as_of_date = wh_col.date().isoformat()
    returning = events['return_open']
    keys = sorted(set(wh_size) | set(pop_stock) | set(pending) | set(returning))
    by_size = []
    for k in keys:
        w, p, pe, rt = wh_size.get(k, 0), pop_stock.get(k, 0), pending.get(k, 0), returning.get(k, 0)
        # 출고대기는 창고에 있는 만큼만 뺀다. 넘치는 수량은 재고 없이 받은 주문(미출고)
        by_size.append({'product': k[0], 'color': k[1], 'size': k[2], 'available_qty': w, 'popup_qty': p,
                        'pending_qty': min(pe, max(w, 0)), 'backorder_qty': max(0, pe - max(w, 0)), 'returning_qty': rt,
                        'real_qty': max(0, w - pe) + rt + p})
    inventory = []
    for n in NAMES:
        rows = [r for r in by_size if r['product'] == n]
        inventory.append({'name': n, 'as_of_date': as_of_date,
                          **{f: sum(r[s] for r in rows) for f, s in (('available_qty', 'available_qty'), ('popup_available_qty', 'popup_qty'),
                                                                   ('pending_qty', 'pending_qty'), ('backorder_qty', 'backorder_qty'),
                                                                   ('returning_qty', 'returning_qty'),
                                                                   ('real_qty', 'real_qty'))}})

    insights = make_insights(D, win3, r3, p3, r3m, r3ch, {i['name']: i for i in inventory}, {p['name']: p for p in products})
    color_size = [{'product': k[0], 'color': k[1], 'size': k[2], 'qty': v, 'qty14': cs14.get(k, 0)} for k, v in sorted(cs.items()) if v]
    recent = [{'name': n, 'recent3d_qty': int(r3.get(n, 0))} for n in NAMES]
    price_weekly = [{'name': n, 'week': wk.isoformat(), 'qty': int(v['qty']), 'avg_price': round(v['paid'] / v['qty'])}
                    for (n, wk), v in sorted(weekly.items()) if v['qty'] > 0]
    channel_mix = [{'name': n, 'channel': ch, 'qty': q} for (n, ch), q in sorted(mix.items()) if q]

    # ── 데이터 신선도 ──
    stale = []
    if (today - wh_col.date()).days > 2:
        stale.append(f'창고 재고가 {wh_col:%m/%d} 기준이에요')
    hub_mod = dt.datetime.fromisoformat(metas['hub']['modifiedTime'].replace('Z', '+00:00')).astimezone(KST).date()
    if (today - hub_mod).days > 2:
        stale.append(f'재무모델 허브가 {hub_mod:%m/%d} 이후 갱신되지 않았어요')
    if (today - D).days > 3:
        stale.append(f'주문 원본이 {D:%m/%d} 결제분까지만 올라와 있어요')

    files3 = [u for u in used if u['end'] >= min(win3).isoformat()]
    meta = {
        'built_at_kst': now.strftime('%Y-%m-%d %H:%M'),
        'sources': [
            {'label': '재무모델 허브 (시즌 KPI·월별·품목 판매·정가·원가)', 'file': metas['hub']['name'], 'modified_kst': kst(metas['hub']['modifiedTime']),
             'note': hub_refreshed.strip('()')},
            {'label': '창고 재고 (재고관리시트)', 'file': metas['inventory']['name'], 'modified_kst': kst(metas['inventory']['modifiedTime']),
             'note': f"{INV_TAB} · {wh_col.strftime('%Y-%m-%d %H:%M')} 컬럼 · 출고대기 = '접수' 열"},
            {'label': '팝업 재고·판매', 'file': metas['popup']['name'], 'modified_kst': kst(metas['popup']['modifiedTime']),
             'note': f"마지막 집계 {pop_stamp} · " + ' / '.join(f'{k} {v}' for k, v in pop_store.items())},
            {'label': '최근 3일 판매 (주문 원본, 팝업 제외)', 'file': ', '.join(f['file'] for f in files3),
             'modified_kst': max((f['modified_kst'] for f in files3), default=''),
             'note': f"결제일 {min(win3)} ~ {D} · " + ' · '.join(f'{k} {v}' for k, v in r3ch.items())},
            {'label': '컬러·사이즈·가격·CM2 (주문 원본 + 무신사 + 팝업)', 'file': f'주문 파일 {len(used)}개', 'modified_kst': '',
             'note': f'{SEASON_START.isoformat()} 이후 결제분 · 반품 회수중 {sum(returning.values())}개'},
        ],
        'warehouse_col': wh_col.strftime('%Y-%m-%d %H:%M'),
        'recent3d': {'start': min(win3).isoformat(), 'end': D.isoformat()},
        'stale': stale,
        'signature': {'hub': metas['hub']['modifiedTime'], 'inventory': metas['inventory']['modifiedTime'],
                      'popup': metas['popup']['modifiedTime'], 'orders': sig},
    }
    new = {'season_kpi': kpi, 'monthly': monthly, 'products': products, 'inventory': inventory,
           'inventory_by_size': by_size, 'color_size': color_size, 'recent3d': recent, 'cancels': cancels,
           'price_kpi': price_kpi, 'price_weekly': price_weekly, 'channel_mix': channel_mix, 'insights': insights, 'meta': meta}

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
