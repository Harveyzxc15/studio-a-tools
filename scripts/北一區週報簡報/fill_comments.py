#!/usr/bin/env python3
"""把評語寫進 build_deck.py 產好的簡報，並匯出 PDF。

    python3 fill_comments.py 27Q1W03            # 讀 輸出/北一區週報簡報/27Q1W03/評語.json
    python3 fill_comments.py 27Q1W03 --region 北二區
    python3 fill_comments.py 27Q1W03 --pdf-only # 只重新匯出 PDF（使用者手改過簡報後）

評語.json 格式：{"士林": ["上週店點狀況", "執行進度與改善狀況"], "微風": [...], ...}（本區每店都要有）
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_deck as b  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('label', help='例：27Q1W03')
    ap.add_argument('--pdf-only', action='store_true')
    ap.add_argument('--region', default='北一區', choices=list(b.REGIONS))
    a = ap.parse_args()
    b.configure(a.region)

    work = b.OUT_BASE / a.label
    state = json.loads((work / 'state.json').read_text(encoding='utf-8'))
    deck = Path(state['deck'])
    pdf = work / (deck.stem + '.pdf')
    ops = []
    if not a.pdf_only:
        comments = json.loads((work / '評語.json').read_text(encoding='utf-8'))
        missing = [n for n in b.NAMES if n not in comments or len(comments[n]) != 2 or not all(comments[n])]
        if missing:
            raise SystemExit(f'❌ 評語.json 缺這幾店（每店要兩段）：{missing}')
        D = b.Deck(json.loads(b.kn('dump', deck)))
        done = set()
        for s, t in D.tables(lambda t, s: '店點狀況' in b.cell(t, 0, 1)):
            for r in range(1, t['rows']):
                k = b.store_key(b.cell(t, r, 0))
                if k:
                    for c in (1, 2):
                        ops.append({'op': 'cell', 's': s['i'], 't': t['i'], 'r': r, 'c': c,
                                    'v': comments[k][c - 1].strip()})
                    done.add(k)
        if done != set(b.NAMES):
            raise SystemExit(f'❌ 簡報評語表只找到 {sorted(done)}，版面可能改過')
    p = work / 'ops_comments.json'
    p.write_text(json.dumps(ops, ensure_ascii=False), encoding='utf-8')
    b.kn('apply', deck, p, pdf)
    b.log(f'✅ 評語已寫入、PDF 已匯出：{pdf}')


if __name__ == '__main__':
    main()
