"""評語貼近度：從 評語修改紀錄*.md 算出每週初稿被改了多少，看初稿是不是越寫越像定稿。

    python3 edit_stats.py          # 更新 評語貼近度.md（build_deck 學完上週評語後也會自動跑）

每段分類（去掉空白、句尾「。」後比較）：
  未修改   一字不改
  只改標點 只差空白／標點（北二 Yuni 常見：數字前後空格、句尾句號）
  微調     相似度 ≥ 80%（換幾個詞、補一句）
  改寫     40%～80%（同題目重寫）
  換題     < 40%（整段換成別的主題）
相似度＝difflib.SequenceMatcher（字元層級）。
"""
import re
from difflib import SequenceMatcher
from pathlib import Path

HERE = Path(__file__).resolve().parent
LOGS = [('北一區', HERE / '評語修改紀錄.md'), ('北二區', HERE / '評語修改紀錄_北二區.md')]
OUT = HERE / '評語貼近度.md'
KINDS = ['未修改', '只改標點', '微調', '改寫', '換題']


def parse(text):
    """→ {週次: [(店／欄, 初稿, 定稿)]}；未修改的段落初稿＝定稿＝None。"""
    weeks = {}
    for block in re.split(r'^## ', text, flags=re.M)[1:]:
        label = block.split('（')[0].split()[0]
        items = []
        for m in re.finditer(r'^- (\S+?)(：未修改.*|\n  - 初稿：(.*)\n  - 定稿：(.*))$', block, flags=re.M):
            items.append((m[1], m[3], m[4]))
        weeks[label] = items
    return weeks


def norm(s):
    return re.sub(r'[\s。．.，,、；;：:（）()]', '', s or '')


def classify(draft, final):
    if draft is None:
        return '未修改', 1.0
    if draft.strip() == final.strip():
        return '未修改', 1.0
    if norm(draft) == norm(final):
        return '只改標點', 1.0
    r = SequenceMatcher(None, norm(draft), norm(final)).ratio()
    return ('微調' if r >= 0.8 else '改寫' if r >= 0.4 else '換題'), r


def kpi(cnt, n):
    """KPI（2026-11-09 定）：換題 ≤1 段，且「未修改＋只改標點」過半。"""
    ok = cnt['換題'] <= 1 and cnt['未修改'] + cnt['只改標點'] > n / 2
    return '✅' if ok else '❌'


def table(region, path):
    if not path.exists():
        return []
    rows = [f'## {region}', '', '| 週次 | 段數 | ' + ' | '.join(KINDS) + ' | 平均相似度 | KPI | 換題的段落 |',
            '|---|---|' + '---|' * len(KINDS) + '---|---|---|']
    for label, items in parse(path.read_text(encoding='utf-8')).items():
        cnt = dict.fromkeys(KINDS, 0)
        sims, swapped = [], []
        for who, d, f in items:
            k, r = classify(d, f)
            cnt[k] += 1
            sims.append(r)
            if k == '換題':
                swapped.append(who.replace('／', ' '))
        if not items:
            continue
        rows.append(f'| {label} | {len(items)} | ' + ' | '.join(str(cnt[k]) for k in KINDS)
                    + f' | {sum(sims) / len(sims):.0%} | {kpi(cnt, len(items))} | {"、".join(swapped) or "-"} |')
    return rows + ['']


def main():
    md = ['# 評語貼近度（初稿 vs 定稿）', '',
          '由 `edit_stats.py` 從評語修改紀錄自動產生。目標：「未修改＋只改標點」逐週變多、「換題」變少。'
          ' **KPI**：每週換題 ≤1 段，且「未修改＋只改標點」超過一半（2026-11-09 起，目標 W07～W10 內連續達標）。',
          '換題代表選題跟使用者想講的不一樣（看修改紀錄裡定稿講了什麼），比用字更值得先改。', '']
    for region, path in LOGS:
        md += table(region, path)
    OUT.write_text('\n'.join(md), encoding='utf-8')
    return OUT


if __name__ == '__main__':
    print(main().read_text(encoding='utf-8'))
