# -*- coding: utf-8 -*-
"""
遗器分解判断（命令行小工具）——2026-11-20 新增，配合 main.py 的「分解建议」页使用。

它本身不算概率：所有概率/门槛都已经由 main.py 写进 分析结果-豆包.xlsx 的「分解建议」页，
这个脚本只做两件事：
    ① 把那一行查出来（套装 + 部位）
    ② 用你手上的初始有效条数，套进门槛，给出"留 / 赌 / 分解"的结论

用法（在项目根目录或任意目录都行）：
    python code\分解判断-豆包.py 司铎 头 2
    python code\分解判断-豆包.py 司铎 头 2 3      # 第 4 个参数 = +0 时总条数（3 或 4），默认 4
    python code\分解判断-豆包.py --list 司铎       # 只看这个套装的六个部位现状

参数说明：
    套装：可以写全称，也可以写任意片段/简称（如「司铎」「铁骑」「快枪手」）
    部位：头 / 手 / 衣服 / 鞋子 / 球 / 绳子
    初始有效条数：这件 +0 时，4 条副词条里有几条是你的有效词条（0~4）
    +0 总条数：4 = 显示 4 条（5 次强化，满级 9 投）；3 = 显示 3 条、第 4 条已知（4 次强化，满级 8 投）
"""
import sys
import os

import pandas as pd

POS_NAMES = ['头', '手', '衣服', '鞋子', '球', '绳子']
SHEET = '分解建议'


def find_analyze_file():
    """优先用项目根目录下的分析结果表；脚本被挪走时给出明确提示。"""
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    path = os.path.join(root, '分析结果-豆包.xlsx')
    if not os.path.exists(path):
        print("!! 找不到分析结果表：%s" % path)
        print("   先双击 02-一键分析.bat（或 python code\\main.py）生成「分解建议」页再来查")
        sys.exit(1)
    return path


def load_rows(path):
    df = pd.read_excel(path, sheet_name=SHEET)
    rows = []
    for _, r in df.iterrows():
        rows.append({k: r[k] for k in df.columns})
    return rows


def pick_set(rows, query):
    """套装模糊匹配：先精确，再"包含"，多个候选时列出来让你写准一点。"""
    names = sorted({r['套装名称'] for r in rows})
    exact = [n for n in names if n == query]
    if exact:
        return exact[0]
    hit = [n for n in names if query in n]
    if len(hit) == 1:
        return hit[0]
    if not hit:
        print("!! 没找到包含「%s」的套装。可用套装共 %d 个，举几个例子：%s"
              % (query, len(names), '、'.join(names[:8])))
        return None
    print("!! 「%s」匹配到多个套装，请写更准一点：%s" % (query, '、'.join(hit)))
    return None


def to_int(v):
    """门槛列可能是数字，也可能是「不可达」这种字符串。"""
    try:
        return int(v)
    except Exception:
        return None


def read_thresholds():
    """
    阈值不重复写在脚本里 —— 从生成表的「分解阈值」页读（01 从配置表带过来的），
    这样你在配置表改 10/40/严格，这里跟着变。读不到就用默认。
    """
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        '星铁装备刷取-自动生成-豆包.xlsx')
    low, high, strict = 10.0, 40.0, 1
    try:
        df = pd.read_excel(path, sheet_name='分解阈值')
        for _, r in df.iterrows():
            name = str(r.iloc[0])
            try:
                val = float(r.iloc[1])
            except Exception:
                continue
            if '阈值' in name and '低' in name:
                low = val
            elif '阈值' in name and '高' in name:
                high = val
            elif '严格' in name:
                strict = 1 if val >= 0.5 else 0
    except Exception:
        pass
    return low, high, strict


def prob_at_least(n, p, t):
    """P(X >= t)，X ~ Bin(n, p)。n 只会是 4 或 5，直接枚举。"""
    if t <= 0:
        return 1.0
    if t > n:
        return 0.0
    from math import comb
    return float(sum(comb(n, k) * (p ** k) * ((1 - p) ** (n - k)) for k in range(t, n + 1)))


def judge(row, s, base, n):
    """
    返回 (结论标签, 说明文字列表)。标签 ∈ {分解, 谨慎赌, 建议强化, 留着}

    判定顺序（和「分解建议」页两列结论一一对应）：
        ① 保底：歪 2 条后仍能超过现役最低那件？
        ② 升级：按这个 s 实际算"超过现役最高"的概率 p
    """
    notes = []
    if base == 4:
        gate_floor = row['保底门槛 s≥(4条)']
    else:
        gate_floor = row['保底门槛 s≥(3条)']

    T, T_min, cap = int(row['现役最高词条数']), int(row['现役最低词条数']), int(row['初始可中上限'])
    best = s + n                # 全中
    worst = s + max(n - 2, 0)   # 歪 2 条

    notes.append("你的这件：初始 %d 条基础，有效 %d 条 → 最理想 %d 条 / 歪2条 %d 条"
                 % (base, s, best, worst))
    notes.append("现役：最高 %d 条、最低 %d 条（该部位最多只能有 %d 条有效）" % (T, T_min, cap))

    if s > cap:
        notes.append("注意：该部位有效词条种类只有 %d 条，初始不可能有 %d 条有效，请检查输入" % (cap, s))
        return '分解', notes

    low, high, strict = read_thresholds()

    # ① 保底
    gf = to_int(gate_floor)
    if gf is not None and s >= gf:
        notes.append("保底：歪 2 条仍有 %d 条，超过现役最低 %d 条 → 有保留价值" % (worst, T_min))
        floor_ok = True
    else:
        notes.append("保底：歪 2 条只有 %d 条，够不到现役最低 %d 条 → 不构成保留理由" % (worst, T_min))
        floor_ok = False

    # ② 升级：直接把概率算出来（不依赖表里只给 s=2/s=3 的两列）
    target = T + (1 if strict else 0)
    p = prob_at_least(n, s / 4.0, target - s) * 100
    need = target - s
    notes.append("升级：现役最高 %d 条 → 要打到 %d 条才算超过，这件还需强化命中 %d 次（最多 %d 次）；概率 = %.2f%%"
                 % (T, target, need, n, p))
    if p >= high:
        notes.append("概率 ≥ 高阈值 %.0f%%，属于大概率升级" % high)
    elif p >= low:
        notes.append("概率在 低%.0f%% ~ 高%.0f%% 之间，只能算有机会" % (low, high))
    else:
        notes.append("概率低于低阈值 %.0f%%，基本没戏" % low)

    if floor_ok:
        return '留着', notes
    if p >= high:
        return '建议强化', notes
    if p >= low:
        return '谨慎赌', notes
    return '分解', notes


def main():
    args = sys.argv[1:]

    if not args:
        print(__doc__)
        return

    path = find_analyze_file()
    rows = load_rows(path)

    if args[0] == '--list':
        if len(args) < 2:
            print("用法：python code\\分解判断-豆包.py --list 套装片段")
            return
        eq = pick_set(rows, args[1])
        if not eq:
            return
        print("【%s】六个部位现状：" % eq)
        for pos in POS_NAMES:
            r = next((x for x in rows if x['套装名称'] == eq and x['部位名称'] == pos), None)
            if not r:
                print("  %s：这个套装没有该部位（外圈套装只有头/手/衣服/鞋子，内圈只有球/绳子）" % pos)
                continue
            print("  %s：现役最高 %s / 最低 %s 条｜升级门槛(4条) %s、(3条) %s｜保底门槛(4条) %s、(3条) %s"
                  % (pos, r['现役最高词条数'], r['现役最低词条数'],
                     r['升级门槛 s≥(4条)'], r['升级门槛 s≥(3条)'],
                     r['保底门槛 s≥(4条)'], r['保底门槛 s≥(3条)']))
        return

    if len(args) < 3:
        print("用法：python code\\分解判断-豆包.py 套装 部位 初始有效条数 [+0总条数]")
        print("例：  python code\\分解判断-豆包.py 司铎 头 2 4")
        return

    eq_query, pos, s = args[0], args[1], int(args[2])
    base = int(args[3]) if len(args) > 3 else 4
    if pos not in POS_NAMES:
        print("!! 部位只能是：%s" % '、'.join(POS_NAMES))
        return
    if base not in (3, 4):
        print("!! 第 4 个参数只能是 3 或 4（+0 时显示的副词条数）")
        return
    if not 0 <= s <= 4:
        print("!! 初始有效条数只能是 0~4")
        return

    eq = pick_set(rows, eq_query)
    if not eq:
        return
    row = next((x for x in rows if x['套装名称'] == eq and x['部位名称'] == pos), None)
    if row is None:
        print("!! 「%s」没有「%s」部位（外圈套装只有头/手/衣服/鞋子，内圈只有球/绳子）" % (eq, pos))
        return

    label, notes = judge(row, s, base, 5 if base == 4 else 4)
    print("=" * 58)
    print("套装：%s ｜ 部位：%s ｜ 使用者 %s 人（%s）"
          % (eq, pos, row['使用者数'], str(row['能帮到的角色']).rstrip('、')))
    print("-" * 58)
    for t in notes:
        print("· " + t)
    print("-" * 58)
    print("结论：%s" % {'分解': '直接分解（留着也打不出一件更好的）',
                       '谨慎赌': '谨慎：有机会但概率不高，只在你缺这件的时候赌',
                       '建议强化': '建议强化（大概率能超过现役）',
                       '留着': '留着（保底就有价值）'}[label])


if __name__ == '__main__':
    main()
