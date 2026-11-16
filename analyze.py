#!/usr/bin/env python3
"""本中心 數據清洗與交叉分析（依資料分析規格）。

用法:
  python3 analyze.py --input 已分類.xlsx [--codes 歷史檔...] --outdir 輸出資料夾 [--title 2026.10] [--judge 學員編號|學員姓名|單位=通路|本部 ...] [--weeks 10/01-10/02,10/03-10/09,...]

輸入：classify.py 產出（或使用者補完）的活頁簿，含 北區轉入/北區轉出/南區轉入/南區轉出 4 個分頁與「單位」「來源」欄。
輸出：來源分析統計表 / 課程類別分析表 / 單位分類與課程統計表 —— 各一個 CSV，一份 Markdown 報告，一份 xlsx；加 --weeks 時另出週營收 CSV。
"""
import argparse, os, re, collections
import openpyxl
import pandas as pd

YS_UNITS = {'本部', '南區本部', '北區本部'}


# ---------- 模組三：課程類別 ----------
def classify_course(course_name):
    c = str(course_name).strip()
    if '家教' in c:
        return '家教課'
    ynso_12m = ['45堂1對5外師Speaking口說課', '45堂1對5 中師Prep課暖身課', '十二個月會話方案', '12個月1對1中師課', '12個月會話方案']
    ynso_6m = ['6個月1對1中師Tutor演練課', '六個月會話方案', '20堂1對5外師Speaking口說課', '20堂1對5 中師Prep課暖身課', '6個月會話方案', '6個月1對1']
    if any(k in c for k in ynso_12m):
        return '會話方案12個月'
    if any(k in c for k in ynso_6m):
        return '會話方案6個月'
    if any(k in c for k in ['1對5', '1對1', '會話方案', 'Prep', 'Speaking', '口說課']):
        if '12' in c or '十二' in c:
            return '會話方案12個月'
        elif '3' in c or '三' in c:
            return '會話方案3個月'
        return '會話方案6個月'
    if '研究所' in c or '研英' in c:
        return '研究所英文'
    if '聽力正音班' in c:
        return '聽力正音班'
    if '初階文法' in c:
        return '初階文法必考17個觀念-雲端'
    if '19議題' in c:
        return '19議題類別'
    u = c.upper()
    if any(k in u for k in ['多益', 'TOEIC']):
        return '多益系列'
    if any(k in u for k in ['托福', 'TOEFL']):
        return '托福系列'
    if any(k in u for k in ['雅思', 'IELTS']):
        return '雅思系列'
    if any(k in u for k in ['GRE', 'GMAT', '填空', '字彙', '閱讀', '邏輯', '計量', '分析寫作']):
        return 'GRE/GMAT系列'
    return '其他'


# ---------- 模組二：通路/本部 ----------
def get_category_and_split(unit, src, amount):
    u = str(unit or '').strip() or '未知'
    s = str(src or '').strip()
    if '留學聯招' in s or '升學聯招' in s:
        return [('本部', amount)]
    if '聯招' in s:
        return [('通路', amount * 0.5), ('本部', amount * 0.5)]
    return [('本部' if u in YS_UNITS else '通路', amount)]


# ---------- 模組四：處別 ----------
def unit_division(region, unit):
    u = str(unit or '').strip()
    if u == '分館D':
        return '分館D'
    if u == '分館E':
        return '分館E'
    if u in ('轉學通路', '轉學部'):
        return '轉學通路'
    if region == '南區' and u in ('校園D', '校園E', '南區', '分館B', '外部通路', '外部通路', '本部', '南區本部'):
        return '南區處'
    if region == '北區' and u in ('校園C', '校園A', '校園B', '分館A', '分館C', '分館C', '本部', '北區本部'):
        return '北區處'
    # 規格書未列的單位（如 分館F、其他區）→ 依轉入/轉出的分頁地區歸處（通路或本部需人工判斷，見 needs_judge）
    return '南區處' if region == '南區' else '北區處'


LISTED_UNITS = {'本部', '南區本部', '北區本部', '校園D', '校園E', '南區', '分館B', '外部通路', '外部通路',
                '校園C', '校園A', '校園B', '分館A', '分館C', '分館C', '分館D', '分館D', '分館E', '分館E', '轉學通路', '轉學部'}


def needs_judge(unit):
    """規格書處別清單沒列到的單位 → 通路／本部要人工判斷。"""
    return str(unit or '').strip() not in LISTED_UNITS


def load(path, code_map, judge):
    wb = openpyxl.load_workbook(path, data_only=True)
    recs = []
    for ws in wb.worksheets:
        m = re.fullmatch(r'(北區|南區)(轉入|轉出)', ws.title.strip())
        if not m:
            continue
        region, io = m.groups()
        rows = list(ws.iter_rows(values_only=True))
        hi = next(i for i, r in enumerate(rows) if r and r[0] in ('轉入區域', '轉出區域', '課程入帳區域'))
        h = [str(c) if c else '' for c in rows[hi]]
        icode = next(i for i, c in enumerate(h) if c.endswith('課程編號'))
        iname = next((i for i, c in enumerate(h) if c.endswith('課程名稱')), None)
        iamt = next(i for i, c in enumerate(h) if c.endswith('金額'))
        isid, iu, isrc = h.index('學員編號'), h.index('單位'), h.index('來源')
        istu = h.index('學員姓名') if '學員姓名' in h else None
        idt = next(i for i, c in enumerate(h) if c in ('轉入日期', '退費日期', '轉出日期'))
        for r in rows[hi + 1:]:
            if not r or not r[isid]:
                continue
            code = r[icode]
            name = r[iname] if iname is not None else None
            if not name:
                name = code_map.get(code, '')
            amt = float(r[iamt] or 0) * (1 if io == '轉入' else -1)
            unit, src = (r[iu] or ''), (r[isrc] or '')
            stu = str(r[istu]).strip() if istu is not None and r[istu] else ''
            if needs_judge(unit):
                # 規格書未列的單位：通路或本部由人工判斷（--judge 學員編號/學員姓名/單位=通路|本部）
                cat0 = judge.get(str(r[isid])) or judge.get(stu) or judge.get(str(unit).strip())
                splits = [(cat0 or '待判斷', amt)]
            else:
                splits = get_category_and_split(unit, src, amt)
            for cat, a in splits:
                recs.append(dict(region=region, io=io, date=str(r[idt] or '')[5:10], sid=r[isid], stu=stu, unit=str(unit).strip(), src=str(src).strip(),
                                 cat=cat, amt=a, course=name, cgroup=classify_course(name),
                                 div=unit_division(region, unit)))
    return pd.DataFrame(recs)


def code_names(paths):
    mp = {}
    for p in paths:
        try:
            wb = openpyxl.load_workbook(p, data_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                for i, v in enumerate(r[:-1]):
                    if isinstance(v, str) and re.fullmatch(r'[A-Z]\d{4}ENG\d{6}', v) and isinstance(r[i + 1], str):
                        mp.setdefault(v, r[i + 1])
    return mp


def net_people(d):
    """淨人數 = 不重複轉入學員數 − 不重複轉出學員數（以學員編號去重）。"""
    return d.loc[d.io == '轉入', 'sid'].nunique() - d.loc[d.io == '轉出', 'sid'].nunique()


def src_label(row):
    """通路的列來源叫「通路」（聯招拆分出來的半筆保留聯招名稱，與本部那半筆對應）；本部的列用分類檔的來源。"""
    if row['cat'] == '通路' and '聯招' not in row['src']:
        return '通路'
    return row['src'] or '未填'


def cats(df):
    return [c for c in ['通路', '本部', '待判斷'] if c in set(df.cat)]


def report_sources(df):
    out = []
    df = df.assign(srcl=df.apply(src_label, axis=1))
    for region in ['南區', '北區']:
        dr = df[df.region == region]
        for cat in cats(dr):
            dc = dr[dr.cat == cat]
            g = dc.groupby('srcl')['amt'].sum().sort_values(ascending=False)
            for s, a in g.items():
                out.append([region, cat, s, a, net_people(dc[dc.srcl == s])])
            out.append([region, cat, f'{cat}小計', dc.amt.sum(), net_people(dc)])
        out.append([region, '', f'{region}總計', dr.amt.sum(), net_people(dr)])
    out.append(['全體', '', '全體總計', df.amt.sum(), net_people(df)])
    return pd.DataFrame(out, columns=['地區', '分類', '來源', '總金額', '人數'])


def report_courses(df):
    out = []
    for cat in cats(df):
        dc = df[df.cat == cat]
        for cg, a in dc.groupby('cgroup')['amt'].sum().sort_values(ascending=False).items():
            out.append([cat, cg, a, net_people(dc[dc.cgroup == cg])])
        out.append([cat, f'{cat}小計', dc.amt.sum(), net_people(dc)])
    out.append(['全體', '全體總計', df.amt.sum(), net_people(df)])
    return pd.DataFrame(out, columns=['分類', '課程類別', '總金額', '淨人數'])


def report_units(df):
    out = []
    order = ['北區處', '南區處', '分館D', '分館E', '轉學通路']
    divs = [d for d in order if d in set(df['div'])] + sorted(set(df['div']) - set(order))
    for dv in divs:
        dd = df[df['div'] == dv]
        for cat in cats(dd):
            dc = dd[dd.cat == cat]
            if dc.empty:
                continue
            for cg, a in dc.groupby('cgroup')['amt'].sum().sort_values(ascending=False).items():
                out.append([dv, cat, cg, a, net_people(dc[dc.cgroup == cg])])
            out.append([dv, cat, f'{cat}小計', dc.amt.sum(), net_people(dc)])
        out.append([dv, '', f'{dv}總計', dd.amt.sum(), net_people(dd)])
    out.append(['全體', '', '全體總計', df.amt.sum(), net_people(df)])
    return pd.DataFrame(out, columns=['單位分類', '分類', '課程類別', '總金額', '淨人數'])


def report_weekly(df, weeks):
    """週營收（0.週營收-本區）：北區＝北區的本部、南區＝南區的本部、通路＝兩區全部通路。
    以轉入日期／退費日期分週；weeks 例：['10/01-10/02', '10/03-10/09', ...]。"""
    out = []
    tot = collections.defaultdict(float)
    for w in weeks:
        a, b = [x.strip() for x in w.split('-')]
        dw = df[(df.date >= a) & (df.date <= b)]
        vals = {'北區': dw[(dw.cat == '本部') & (dw.region == '北區')].amt.sum(),
                '南區': dw[(dw.cat == '本部') & (dw.region == '南區')].amt.sum(),
                '通路': dw[dw.cat == '通路'].amt.sum()}
        for k, v in vals.items():
            out.append([w, k, round(v)])
            tot[k] += v
    for k in ('北區', '南區', '通路'):
        out.append(['小計', k, round(tot[k])])
    out.append(['合計', '', round(sum(tot.values()))])
    covered = sum(((df.date >= w.split('-')[0]) & (df.date <= w.split('-')[1])).sum() for w in weeks)
    return pd.DataFrame(out, columns=['週', '類別', '金額']), covered == len(df)


def to_md(t):
    cols = list(t.columns)
    lines = ['| ' + ' | '.join(cols) + ' |', '|' + '|'.join('---:' if c in ('總金額', '人數', '淨人數') else '---' for c in cols) + '|']
    for _, r in t.iterrows():
        cells = []
        bold = any(str(r[c]).endswith(('小計', '總計')) for c in cols)
        for c in cols:
            v = r[c]
            s = f'{v:,.0f}' if c == '總金額' else (f'{int(v)}' if c in ('人數', '淨人數') else str(v))
            cells.append(f'**{s}**' if bold and s else s)
        lines.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--codes', nargs='*', default=[], help='用來還原課程名稱的歷史檔（課程名稱缺漏時才需要）')
    ap.add_argument('--outdir', required=True)
    ap.add_argument('--title', default='')
    ap.add_argument('--weeks', default='', help='週營收分週，逗號分隔，例：10/01-10/02,10/03-10/09,10/10-10/16,10/17-10/23,10/24-10/31')
    ap.add_argument('--judge', nargs='*', default=[], help='規格書未列單位的人工判斷：學員編號/學員姓名/單位=通路|本部')
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    judge = dict(kv.split('=', 1) for kv in a.judge)
    df = load(a.input, code_names(a.codes), judge)
    t1, t2, t3 = report_sources(df), report_courses(df), report_units(df)
    tot = [t.iloc[-1]['總金額'] for t in (t1, t2, t3)]
    ok = max(tot) - min(tot) < 0.5
    names = ['1_來源分析統計表', '2_課程類別分析表', '3_單位分類與課程統計表']
    pre = f'{a.title}_' if a.title else ''
    for n, t in zip(names, (t1, t2, t3)):
        t.to_csv(os.path.join(a.outdir, f'{pre}{n}.csv'), index=False, encoding='utf-8-sig')
    with pd.ExcelWriter(os.path.join(a.outdir, f'{pre}交叉分析.xlsx')) as xw:
        for n, t in zip(names, (t1, t2, t3)):
            t.to_excel(xw, sheet_name=n[2:], index=False)
    if a.weeks:
        tw, full = report_weekly(df, [w for w in a.weeks.split(',') if w.strip()])
        tw.to_csv(os.path.join(a.outdir, f'{pre}4_週營收.csv'), index=False, encoding='utf-8-sig')
        print(tw.to_string(index=False))
        if not full:
            print('警告：有資料的日期不在任何一週內，請檢查 --weeks')
    others = df[df['cgroup'] == '其他'].groupby('course')['amt'].sum()
    pend = df[df.cat == '待判斷'].drop_duplicates(['sid', 'io', 'unit'])
    judged = df[df.unit.map(needs_judge) & (df.cat != '待判斷')].drop_duplicates(['sid', 'unit'])
    md = [f'# 本中心 交叉分析 {a.title}'.strip(), '',
          f'勾稽：三張表全體總金額 = {tot[0]:,.0f} / {tot[1]:,.0f} / {tot[2]:,.0f} → ' + ('**一致（差額 0）**' if ok else '**不一致，請檢查**'), '',
          '## 第一部分：來源分析統計表', '', to_md(t1), '',
          '## 第二部分：課程類別分析表', '', to_md(t2), '',
          '## 第三部分：單位分類與課程統計表', '', to_md(t3), '']
    if len(pend):
        md += ['**待判斷**：規格書未列的單位，已依轉入／轉出的館別歸處；通路或本部請人工判斷後用 --judge 重跑。']
        md += [f'- {r.region}{r.io}｜{r.unit or "（單位空白）"}｜{r.stu or r.sid}（{r.sid}）' for r in pend.itertuples()]
        md += ['']
    if len(judged):
        md += ['人工判斷：' + '、'.join(f'{r.stu or r.sid}（{r.unit}）→ {r.cat}' for r in judged.itertuples()), '']
    if len(others):
        md += ['歸入「其他」的課程：' + '、'.join(f'{k}（{v:,.0f}）' for k, v in others.items()), '']
    open(os.path.join(a.outdir, f'{pre}交叉分析報告.md'), 'w', encoding='utf-8').write('\n'.join(md))
    print('勾稽', '通過' if ok else '失敗', [f'{x:,.0f}' for x in tot])
    if len(pend):
        lst = [(r.unit or '（單位空白）', r.stu or r.sid, r.sid) for r in pend.itertuples()]
        print(f'待判斷（通路/本部）{len(lst)} 筆:', lst[:20], '…' if len(lst) > 20 else '')
    print('輸出到', a.outdir)


if __name__ == '__main__':
    main()
