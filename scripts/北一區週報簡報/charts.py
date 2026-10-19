"""維修中心長條圖（取代 Keynote 原生圖表；原生圖表的資料無法用程式改）。

樣式比照原簡報：深色底（透明，吃投影片背景）、白字、去年／今年以虛線分隔、
3PP 圖為「原廠周邊（藍）＋3PP（綠）」堆疊，保外維修為單一藍色長條。
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

FONT = FontProperties(fname='/System/Library/Fonts/STHeiti Medium.ttc')
BLUE, GREEN = '#0A63C9', '#1F8A1F'
GRID = '#8A8A8A'


def _font(size):
    f = FONT.copy()
    f.set_size(size)
    return f


def _nice_max(v):
    """y 軸上限取整（例：137,000 → 150,000；1,545,033 → 2,000,000）。"""
    if v <= 0:
        return 1
    mag = 10 ** (len(str(int(v))) - 1)
    for k in (1, 1.25, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10):
        if k * mag >= v * 1.05:
            return k * mag
    return 10 * mag


def render(path, labels, year_split, series, w_pt, h_pt, stacked):
    """labels: x 軸標籤；year_split: (去年字樣, 今年字樣, 今年第一根的索引)；
    series: [(名稱, 數值list, 顏色)]，stacked=True 時依序往上疊。"""
    n = len(labels)
    xs = list(range(n))
    tops = [sum(s[1][i] for s in series) if stacked else max(s[1][i] for s in series) for i in xs]
    ymax = _nice_max(max(tops))

    fig = plt.figure(figsize=(w_pt / 100, h_pt / 100), dpi=250)
    ax = fig.add_axes([0.09, 0.08, 0.90, 0.80])
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    bottoms = [0] * n
    for name, vals, color in series:
        ax.bar(xs, vals, width=0.72, bottom=bottoms, color=color, zorder=3)
        for x, v, b in zip(xs, vals, bottoms):
            if not v:
                continue
            small = v < ymax * 0.06                 # 太矮的柱子把數字放在柱子上方
            ax.text(x, b + v, f'{v:,.0f}', ha='center', va='bottom' if small else 'top',
                    color='white', fontsize=6.8, fontweight='bold', zorder=4)
        if stacked:
            bottoms = [b + v for b, v in zip(bottoms, vals)]

    ax.set_ylim(0, ymax)
    ax.set_xlim(-0.6, n - 0.4)
    ticks = [ymax * i / 4 for i in range(5)]
    ax.set_yticks(ticks)
    ax.set_yticklabels([f'{t:,.0f}' for t in ticks], color='white', fontsize=8.5, fontweight='bold')
    ax.set_xticks(xs)
    ax.set_xticklabels(labels, color='white', fontproperties=_font(8.5))
    ax.tick_params(length=0)
    ax.grid(axis='y', color=GRID, linewidth=0.6, zorder=0)
    for sp in ax.spines.values():
        sp.set_visible(False)

    ly, cy, split = year_split
    box = dict(facecolor='black', edgecolor='none', pad=2.5)
    ax.text(-0.5, ymax, ly, color='white', fontproperties=_font(9), va='center', bbox=box, zorder=5)
    if split is not None:
        ax.axvline(split - 0.5, color='white', linestyle=(0, (1, 2.5)), linewidth=1.1, zorder=4)
        ax.text(split - 0.5, ymax, cy, color='white', fontproperties=_font(9),
                va='center', ha='center', bbox=box, zorder=5)

    if stacked:
        handles = [plt.Rectangle((0, 0), 1, 1, color=c) for _, _, c in series]
        leg = ax.legend(handles, [s[0] for s in series], loc='upper left', bbox_to_anchor=(0.3, 1.1),
                        ncol=len(series), frameon=False, prop=_font(10), handlelength=0.8)
        for t in leg.get_texts():
            t.set_color('white')
    fig.savefig(path, transparent=True)
    plt.close(fig)
