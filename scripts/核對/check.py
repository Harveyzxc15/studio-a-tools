# -*- coding: utf-8 -*-
"""交付前核對入口。

  python3 check.py run --facts F.json [--facts G.json] --text 草稿.txt --profile daily_n1 [--out 核對單.md]
      → 印核對單；有 ❌ 退出碼 1（不能交付），只有 ⚠️ 退出碼 0 但要逐條確認
  python3 check.py add F.json "士林 ARpedia 本月" 7 --unit cnt --store 士林 --src "arpedia_daily.query_units"
      → 臨時查的數字補進事實檔（一定要寫來源）；--group 同群組會自動衍生差額／增減%

profile：daily_n1 daily_n2 weekly_n1 weekly_n2 analysis
草稿檔：一段一行。週報評語可直接給 評語.json（{"店": ["段1","段2"]}），會逐店逐段核對。
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import factcheck as fc
import topics


def stores_for(profile):
    return fc.N2_STORES if profile.endswith('n2') else fc.N1_STORES


def run(a):
    facts, checks = [], []
    for p in a.facts:
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
        facts += d['facts']
        checks += [tuple(c) for c in d.get('meta', {}).get('checks', []) if c[0] != '✅']
    with open(a.text, encoding='utf-8') as f:
        raw = f.read()
    parts = []
    if a.text.endswith('.json'):
        for store, paras in json.loads(raw).items():
            for i, p in enumerate(paras, 1):
                parts.append((f'{store} 第{i}段', p))
    else:
        speech_md = a.profile.startswith('weekly') and a.text.endswith('.md')     # 週報講稿：去掉 # 標題行
        body = '\n'.join(l for l in raw.splitlines() if not l.lstrip().startswith('#')) if speech_md else raw
        parts.append(('草稿', body.strip()))
    out, bad, warn = [], 0, 0
    for title, text in parts:
        if a.text.endswith('.json'):
            store = title.split()[0]
            text_ctx = f'{store}：{text}'     # 段落沒寫店名時，數字預設屬於這家店
        else:
            text_ctx = text
        md, b, w, _, _ = fc.report(text_ctx, facts, a.profile, stores_for(a.profile),
                                   extra=checks if not out else None, title=f'核對單｜{title}', rule_text=text,
                                   speech=a.profile.startswith('weekly') and a.text.endswith('.md'))
        out.append(md)
        bad += b
        warn += w
    if a.text.endswith('.json') and a.profile.startswith('weekly'):
        # 跨格檢查：本週各格之間、與上週定稿（同資料夾 評語素材.md 的「上週簡報的評語」）比重複與選題
        mat = os.path.join(os.path.dirname(os.path.abspath(a.text)), '評語素材.md')
        last, src = ({}, '找不到評語素材.md')
        if os.path.exists(mat):
            with open(mat, encoding='utf-8') as f:
                last, src = topics.last_finals_from_material(f.read())
            src = src or '評語素材沒有寫來源（舊版素材）'
        # 上上週：同層前兩個週次資料夾的 評語.json（用來抓同店連三週同主題）
        week_dir = os.path.dirname(os.path.abspath(a.text))
        weeks = sorted(w for w in os.listdir(os.path.dirname(week_dir))
                       if w < os.path.basename(week_dir) and os.path.exists(os.path.join(os.path.dirname(week_dir), w, '評語.json')))
        older = {}
        if len(weeks) >= 2:
            with open(os.path.join(os.path.dirname(week_dir), weeks[-2], '評語.json'), encoding='utf-8') as f:
                older = json.load(f)
        res = topics.cross_checks(json.loads(raw), last, src, older)
        lines = ['## 跨格檢查｜重複與選題（本週各格＋上週定稿＋上上週）', '']
        lines += [f'- {st} {name}：{why}' for st, name, why in res]
        out.append('\n'.join(lines))
        bad += sum(st == '❌' for st, _, _ in res)
        warn += sum(st == '⚠️' for st, _, _ in res)
    if a.profile.startswith('weekly') and os.path.basename(a.text) == '講稿.md':
        # 講稿跨週：跟前一個週次資料夾的 講稿.md 比結構與開場
        week_dir = os.path.dirname(os.path.abspath(a.text))
        prev = sorted(w for w in os.listdir(os.path.dirname(week_dir))
                      if w < os.path.basename(week_dir) and os.path.exists(os.path.join(os.path.dirname(week_dir), w, '講稿.md')))
        last_md = None
        if prev:
            with open(os.path.join(os.path.dirname(week_dir), prev[-1], '講稿.md'), encoding='utf-8') as f:
                last_md = f.read()
        res = topics.speech_checks(raw, last_md)
        if res:
            out.append('\n'.join(['## 講稿多元化｜跟上週比', ''] + [f'- {st} {name}：{why}' for st, name, why in res]))
            bad += sum(st == '❌' for st, _, _ in res)
            warn += sum(st == '⚠️' for st, _, _ in res)
    summary = f'\n# 總結：❌ {bad}｜⚠️ {warn} → ' + ('不能交付' if bad else '可交付（⚠️ 要逐條確認）' if warn else '全部通過')
    text = '\n\n'.join(out) + summary
    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f:
            f.write(text)
    print(text)
    sys.exit(1 if bad else 0)


def add(a):
    with open(a.file, encoding='utf-8') as f:
        d = json.load(f)
    if not a.src:
        sys.exit('❌ 要寫 --src（數字從哪裡查的），沒有來源的數字不能進事實檔')
    v = float(a.value.replace(',', '').rstrip('%')) / (100 if a.value.endswith('%') else 1)
    unit = a.unit or ('rate' if a.value.endswith('%') else None)
    d['facts'].append(fc.fact(v, unit, a.store, a.label, g=a.group, src=a.src))
    fc.save_facts(a.file, d['facts'], d.get('meta'))
    print(f'✅ 已加入：{a.label} = {v}（{a.src}）')


def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest='cmd', required=True)
    r = sub.add_parser('run')
    r.add_argument('--facts', action='append', required=True)
    r.add_argument('--text', required=True)
    r.add_argument('--profile', required=True,
                   choices=['daily_n1', 'daily_n2', 'weekly_n1', 'weekly_n2', 'analysis'])
    r.add_argument('--out')
    ad = sub.add_parser('add')
    ad.add_argument('file')
    ad.add_argument('label')
    ad.add_argument('value')
    ad.add_argument('--unit', choices=['rate', 'amt', 'cnt'])
    ad.add_argument('--store', default='')
    ad.add_argument('--group')
    ad.add_argument('--src')
    a = p.parse_args()
    run(a) if a.cmd == 'run' else add(a)


if __name__ == '__main__':
    main()
