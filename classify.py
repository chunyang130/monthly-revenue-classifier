#!/usr/bin/env python3
"""本部 每月轉入/轉出（業績彙整表）分類 + 彙總.

用法:
  python3 classify.py --raw 原始1.xls 原始2.xls ... [--history 已分類1.xls ...] [--sources 名單來源表.xlsx] [--staff 姓名=單位 ...] [--override 學員姓名=來源 ...] --out 輸出.xlsx

- raw: 系統匯出、尚未分類的 4 張表（北區轉入/北區轉出/南區轉入/南區轉出），順序不拘，
  依第1列(轉入/轉出) 與第2列查詢條件中的單位(北區本部/南區本部) 自動辨識。
- history: 以前已分類好的月檔（含「單位」「來源」欄），用來更新人員對照與學員來源記錄。
"""
import argparse, os, re, subprocess, sys, tempfile, collections
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.comments import Comment

# ---- 人員 → 單位 基準表（範例為虛構人員；OWN = 本部自家業務，依分頁填 北區本部/南區本部）----
# 實際使用時，這份對照表會被 --history 歷史分類檔自動覆蓋（以最新月份為準，因為人員會調動）
STAFF = {
 # 本部自家業務（OWN）
 '王小明':'OWN','林小華':'OWN','陳小美':'OWN','張大同':'OWN',
 # 校園合作窗口
 '李校一':'校園A','吳校二':'校園A','周校三':'校園B','鄭校四':'校園C','黃校五':'校園D','許校六':'校園E',
 # 其他分館
 '劉館一':'分館A','楊館二':'分館B','蔡館三':'分館C','謝館四':'分館D','郭館五':'分館E','曾館六':'分館F',
 # 其他
 '何通路':'轉學通路','羅外區':'其他區',
}
# 舊月份的寫法 → 最新寫法
NORM = {'本部':'OWN','北區本部':'OWN','南區本部':'OWN','外部通路':'OWN',
        '南區':'分館B','分館C':'分館C','分館D':'分館D','北區':'分館A'}
OWN_LABEL = {'北區':'北區本部','南區':'南區本部'}
# 彙總時的單位分組
GROUP = {'北區本部':'本部','南區本部':'本部',
         '校園A':'校園','校園B':'校園','校園C':'校園','校園D':'校園','校園E':'校園',
         '分館A':'其他館','分館B':'其他館','分館C':'其他館','分館D':'其他館','分館E':'其他館','分館F':'其他館',
         '轉學通路':'其他','其他區':'其他'}
SHEET_ORDER = ['北區轉入','北區轉出','南區轉入','南區轉出']

YELLOW = PatternFill('solid', fgColor='FFFFFF00')   # 待填
ORANGE = PatternFill('solid', fgColor='FFFCE4D6')   # 系統建議，請確認
HDR = PatternFill('solid', fgColor='FFD9E1F2')


def to_xlsx(path, tmp):
    if path.lower().endswith('.xlsx'):
        return path
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', tmp, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + '.xlsx')


def find_header(rows):
    for i, r in enumerate(rows):
        if r and r[0] in ('轉入區域', '轉出區域', '課程入帳區域'):
            return i
    raise ValueError('找不到表頭列')


def sheet_key(rows):
    title, cond = str(rows[0][0] or ''), str(rows[1][0] or '')
    io = '轉入' if '(轉入)' in title else '轉出'
    unit = cond.split('轉入單位：' if io == '轉入' else '轉出單位：')[-1].split('　')[0]
    br = '北區' if '北區' in unit else '南區' if '南區' in unit else None
    if br is None:
        raise ValueError(f'無法判斷分館: {cond}')
    return br + io, re.search(r'(\d{4}/\d{2}/\d{2})~(\d{4}/\d{2}/\d{2})', cond)


def load_history(paths, tmp):
    staff, src = {}, {}  # name -> (date, unit) ; student -> (date, source)
    for p in paths:
        wb = openpyxl.load_workbook(to_xlsx(p, tmp), data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            try:
                hi = find_header(rows)
            except ValueError:
                continue
            h = [str(c) if c else '' for c in rows[hi]]
            if '單位' not in h:
                continue
            iu, isrc, iname, isid = h.index('單位'), h.index('來源') if '來源' in h else None, h.index('績效人員姓名'), h.index('學員編號')
            idate = next(h.index(c) for c in ('轉入日期', '退費日期', '轉出日期') if c in h)
            for r in rows[hi + 1:]:
                if not r or not r[iname]:
                    continue
                d = str(r[idate] or '')
                u = str(r[iu] or '').strip()
                if u:
                    u = NORM.get(u, u)
                    if r[iname] not in staff or d >= staff[r[iname]][0]:
                        staff[r[iname]] = (d, u)
                s = str(r[isrc] or '').strip() if isrc is not None else ''
                if s and s not in ('?',) and not s.startswith('退換') and u == 'OWN':
                    if r[isid] not in src or d >= src[r[isid]][0]:
                        src[r[isid]] = (d, s)
    return {k: v[1] for k, v in staff.items()}, {k: v[1] for k, v in src.items()}


# ---- 名單來源表（名單來源表（各月諮詢明細））----
GS_SRC = ['聯招','活動','校廣','專案','包套轉出','自來客','親友轉介','轉方案','續報','加報','網單','跨月回班再諮','補費','自營網單','未列來源']
GS_MAP = {'自來客':'自來','網單':'網單','自營網單':'自營網單','續報':'續報','加報':'加報','親友轉介':'親友',
          '校廣':'校廣','聯招':'聯招','活動':'活動','專案':'專案','包套轉出':'包套轉出','轉方案':'轉方案','補費':'補費'}
NOTE_KEYS = [('自營','自營網單'),('網單','網單'),('公網','網單'),('自來','自來'),('字來','自來'),
             ('親友','親友'),('續報','續報'),('加報','加報'),('校廣','校廣')]


def load_source_sheet(path, tmp):
    """回傳 {學員姓名: [(月, 來源, 依據說明, 是否成交), ...]}。支援 .xlsx/.xls 下載檔或「## Sheet name:」文字匯出。
    來源為 '' 代表聯招單位判斷不出，需請使用者填。"""
    import csv, io
    tabs = {}
    if path.lower().endswith(('.xlsx', '.xls')):
        wb = openpyxl.load_workbook(to_xlsx(path, tmp), data_only=True)
        for ws in wb.worksheets:
            tabs[ws.title.strip()] = [['' if c is None else str(c) for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        txt = open(path, encoding='utf-8').read()
        parts = re.split(r'^## Sheet name: (.+)$', txt, flags=re.M)
        for i in range(1, len(parts), 2):
            tabs[parts[i].strip()] = list(csv.reader(io.StringIO(parts[i + 1].strip('\n'))))
    out = collections.defaultdict(list)
    for title, rows in tabs.items():
        mm = re.fullmatch(r'(\d{1,2})月', title)
        if not mm:
            continue
        month = int(mm.group(1))
        hi = next((i for i, r in enumerate(rows[:10]) if len(r) > 2 and r[2].strip() == '學員姓名'), None)
        if hi is None:
            continue
        idx = {re.sub(r'\s', '', c): i for i, c in enumerate(rows[hi])}
        def g(r, col):
            i = idx.get(col)
            return str(r[i]).strip() if i is not None and i < len(r) and r[i] is not None else ''
        for r in rows[hi + 1:]:
            name = g(r, '學員姓名')
            if not name:
                continue
            cols = [c for c in GS_SRC if g(r, c)]
            note = g(r, '學生狀況簡述(如需填詳細請點右鍵插入備註填寫)')
            jzu = g(r, '聯招單位(勿動)')
            label = None
            for c in cols:
                if c in GS_MAP:
                    label = GS_MAP[c]
                    break
            if label is None:  # 跨月回班再諮／未列來源 → 看簡述開頭
                head = note[:6]
                label = next((v for k, v in NOTE_KEYS if k in head), None)
            if label == '網單' and jzu == 'EC電商':
                label = 'EC'
            person = g(r, '聯招人員')
            if label == '聯招':
                label = joint_label(note, jzu, person)  # 判斷不出時回傳 ''（留黃底請使用者填）
            if label is not None:
                why = f"名單來源表 {month}月：{'+'.join(cols) or '簡述'}" + (f"／{jzu}" if jzu else '') + (f"／聯招人員 {person}" if person else '') + f"／顧問 {g(r, '諮詢顧問')}"
                if label == '':
                    why += ' → 聯招單位無法判斷，請填（例：校園A聯招、校園B聯招）'
                out[name].append((month, label, why, g(r, '名單等級(可自行新增下拉選單)').startswith('Y')))
    return out


JZU_MAP = {'留學部': '留學聯招', '升學部': '升學聯招', '轉學部': '轉學聯招'}
CAMPUS = ('校園A', '校園B', '校園C', '校園D', '校園E')


def joint_label(note, jzu, person):
    """聯招 → 「○○聯招」。
    1. 簡述有寫（如「留學聯招」）就照寫
    2. 聯招單位：留學部→留學聯招、升學部→升學聯招、轉學部→轉學聯招
    3. 其他（北區處、南區處、北區本部…）→ 看聯招人員屬哪個校園：校園A聯招／校園B聯招／校園C聯招／校園D聯招／校園E聯招
    4. 判斷不出 → 回傳 ''，留黃底請使用者填
    """
    m = re.search(r'([一-鿿]{2})聯招', note)
    if m:
        return m.group(1) + '聯招'
    if jzu in JZU_MAP:
        return JZU_MAP[jzu]
    u = STAFF.get(person.strip())
    if u in CAMPUS:
        return u + '聯招'
    return ''


def pick_source(recs, month):
    """同月優先，其次最近的前幾個月；同條件下優先已成交(Y)。"""
    if not recs:
        return None
    same = [r for r in recs if r[0] == month]
    pool = same or [r for r in recs if r[0] <= month] or recs
    pool = sorted(pool, key=lambda r: (r[0], r[3]))
    return pool[-1]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--raw', nargs='+', required=True)
    ap.add_argument('--history', nargs='*', default=[])
    ap.add_argument('--staff', nargs='*', default=[], help='補充對照 姓名=單位，例如 王小明=OWN')
    ap.add_argument('--sources', help='名單來源表（各月諮詢明細）（下載的 xlsx 或文字匯出）')
    ap.add_argument('--override', nargs='*', default=[], help='使用者指定來源 學員姓名=來源，例如 王同學=續報')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    tmp = tempfile.mkdtemp()

    ovr = dict(kv.split('=', 1) for kv in a.override)
    gsrc = load_source_sheet(a.sources, tmp) if a.sources else {}
    staff = dict(STAFF)
    hstaff, hsrc = load_history(a.history, tmp)
    staff.update(hstaff)
    for kv in a.staff:
        k, v = kv.split('=')
        staff[k] = NORM.get(v, v)

    sheets, period = {}, None
    for p in a.raw:
        ws = openpyxl.load_workbook(to_xlsx(p, tmp), data_only=True).worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        key, m = sheet_key(rows)
        period = period or (m.groups() if m else None)
        sheets[key] = rows

    # 同月份 本部業務 既有轉出又有轉入的學員 → 建議「退換」
    def students(key):
        rows = sheets.get(key)
        if not rows:
            return set()
        hi = find_header(rows)
        return {r[3] for r in rows[hi + 1:] if r and r[3]}
    swap = {br: students(br + '轉入') & students(br + '轉出') for br in ('北區', '南區')}

    out = openpyxl.Workbook(); out.remove(out.active)
    summary = collections.defaultdict(lambda: collections.defaultdict(float))
    todo = []
    counts = collections.Counter()

    for key in SHEET_ORDER:
        if key not in sheets:
            continue
        rows = sheets[key]
        br = key[:2]
        hi = find_header(rows)
        hdr = [c for c in rows[hi] if c is not None] + ['單位', '來源']
        ncol = len(hdr)
        iname, isid, iamt = hdr.index('績效人員姓名'), hdr.index('學員編號'), 9
        istu = hdr.index('學員姓名') if '學員姓名' in hdr else None
        idt = 8
        ws = out.create_sheet(key)
        for i in range(hi):
            ws.append([rows[i][0]])
        ws.append(hdr)
        data = []
        for r in rows[hi + 1:]:
            r = r[:ncol - 2]
            if not any(v not in (None, '') for v in r):
                continue
            name = r[iname]
            u = staff.get(name)
            label = OWN_LABEL[br] if u == 'OWN' else u
            src, note, fill_u, fill_s = '', None, None, None
            if label is None:
                fill_u = YELLOW
                todo.append((key, name, '單位', '新人員'))
            if u == 'OWN':
                sid = r[isid]
                stu = str(r[istu]).strip() if istu is not None else ''
                if stu in ovr:
                    src, note = ovr[stu], '使用者指定'
                elif sid in swap[br]:
                    src, note = '退換', '同月同時出現在轉入與轉出（退費後重報）→ 建議「退換」'
                elif istu is not None and (hit := pick_source(gsrc.get(stu), int(str(r[idt])[5:7]))):
                    src, note = hit[1], hit[2]   # hit[1] 為 '' 時＝聯招單位判斷不出，維持黃底
                elif sid in hsrc:
                    src, note = hsrc[sid], f'沿用該學員前次來源「{hsrc[sid]}」'
                if note == '使用者指定' and src != '待確認':
                    fill_s = None
                elif src and src != '待確認':
                    fill_s = ORANGE
                else:
                    fill_s = YELLOW
                    todo.append((key, name, '來源', sid))
            data.append((r, label, src, note, fill_u, fill_s, u))
        # 排序：本部 → 校園 → 其他館 → 其他；同組內依單位、來源、業務
        order = {'本部': 0, '校園': 1, '其他館': 2, '其他': 3}
        data.sort(key=lambda t: (order.get(GROUP.get(t[1] or '', ''), 4), t[1] or '', t[2] or '', str(t[0][iname])))
        for r, label, src, note, fu, fs, u in data:
            ws.append(list(r) + [label or '', src])
            rr = ws.max_row
            cu, cs = ws.cell(rr, ncol - 1), ws.cell(rr, ncol)
            if fu: cu.fill = fu; cu.comment = Comment(f'新人員「{r[iname]}」不在對照表，請填單位', 'Claude')
            if fs: cs.fill = fs
            if note: cs.comment = Comment(note, 'Claude')
            amt = r[iamt] or 0
            sk = src or ('—' if u != 'OWN' else '待填')
            summary[(key, GROUP.get(label or '', '未分類'), label or '未分類')][sk] += amt
            counts[(key, label or '未分類', sk)] += 1
        for c in ws[hi + 1]:
            c.fill, c.font = HDR, Font(bold=True)
        ws.auto_filter.ref = f'A{hi+1}:{openpyxl.utils.get_column_letter(ncol)}{ws.max_row}'
        ws.freeze_panes = f'A{hi+2}'
        for i, w in enumerate([10, 16, 14, 14, 9, 16, 28, 11, 11, 9, 11, 12, 12, 11, 10, 12][:ncol], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ---- 彙總 ----
    ws = out.create_sheet('彙總', 0)
    ws.append([f'本部 轉入/轉出 分類彙總　{period[0]}~{period[1]}' if period else '本部 轉入/轉出 分類彙總'])
    ws['A1'].font = Font(bold=True, size=13)
    ws.append(['黃底＝待填；淡橘底＝系統建議（游標移到儲存格看依據），確認後可清除底色'])
    ws.append([])
    ws.append(['分頁', '分組', '單位', '來源', '筆數', '金額'])
    for c in ws[4]:
        c.fill, c.font = HDR, Font(bold=True)
    tot = collections.defaultdict(float)
    for (key, grp, label), d in sorted(summary.items(), key=lambda kv: (SHEET_ORDER.index(kv[0][0]), {'本部':0,'校園':1,'其他館':2,'其他':3}.get(kv[0][1], 4), kv[0][2])):
        for s, amt in sorted(d.items(), key=lambda x: -x[1]):
            ws.append([key, grp, label, s, counts[(key, label, s)], amt])
            tot[key] += amt
    ws.append([])
    ws.append(['合計'])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    for k in SHEET_ORDER:
        if k in tot:
            ws.append([k, '', '', '', '', tot[k]])
    for br in ('北區', '南區'):
        if br + '轉入' in tot or br + '轉出' in tot:
            ws.append([br + '淨額（轉入−轉出）', '', '', '', '', tot.get(br + '轉入', 0) - tot.get(br + '轉出', 0)])
            ws.cell(ws.max_row, 1).font = Font(bold=True)
    for row in ws.iter_rows(min_row=5, min_col=6, max_col=6):
        row[0].number_format = '#,##0'
    for col, w in zip('ABCDEF', [22, 8, 10, 14, 7, 12]):
        ws.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(a.out) or '.', exist_ok=True)
    out.save(a.out)
    nt = collections.Counter((t[0], t[2]) for t in todo)
    print('完成:', a.out)
    print('新人員:', sorted({t[1] for t in todo if t[3] == '新人員'}) or '無')
    print('待填:', dict(nt) or '無')


if __name__ == '__main__':
    main()
