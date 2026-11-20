# -*- coding: utf-8 -*-
"""
json_关键信息-豆包.py
================================================================================
第二个脚本：只提炼项目需要的关键信息，直接产出可喂给 main.py 的「角色遗器」表。

和第一个脚本（json_全量提取-豆包.py）的分工：
    第一个脚本 = 全量勘察，一条不漏地摊平 JSON 里所有字段，用于了解数据长什么样；
    本脚本     = 定点提炼，只输出「角色遗器」那 17 列，用于替代手工数词条。

复用：计数/翻译/映射/次数还原这些口径全部 import 第一个脚本，保证两个脚本不会各说各话。
      改口径只需改第一个脚本，这里自动同步。

输出：星铁装备刷取-自动生成-豆包.xlsx（7 页签）
    · 角色遗器          配置表登记、有效词条已填、且遗器数据完整的角色 → 可直接喂 main.py
    · 有效词条待补      新补登的角色 + 配置表里有行但「有效词条」还空着的角色；
                        六件上出现过的副词条按件数排好，照着勾 3~4 条即可
    · 数据不完整        JSON 里有、但装备凑不齐的角色（缺件 / 未满级 / 数值异常），逐条给原因和办法
    · 校验              数据完整性检查（名单 / 有效词条 / 权重是否齐）
    · 生成说明          口径说明
    · 遗器相关 / 主词条分布   从「项目配置-豆包.xlsx」原样带过来，整份文件可直接当 main.py 的输入表

> 识别到新角色（JSON 有、配置表没登记）时，本脚本会**自动补登**：
>   config「角色有效词条」加一行、有效词条留空；角色权重.xlsx 加一行、权重 1.0。
> 只追加行尾、不动其它页，写前备份到 bk\备份-豆包\08-自动补登前\，且幂等。
> 这样用户打开配置表就能看到"这两个角色还没填词条"，而不是只在生成表里看到一行提示。

> 「数据不完整」页是为多账号场景加的：换一个号，角色池、装备进度都不一样，
> 没穿满 6 件或数值异常的角色的绝不能塞进「角色遗器」页（否则 main.py 会把缺失部位
> 当成"0 条有效"，概率虚高）。这类角色单独列在那一页，补齐装备重跑即可回到主表。
>
> 2026-11-19 起删掉了原来的「角色遗器-全76」页：角色名单已统一由配置表决定，而配置表覆盖了
> JSON 全部角色，两个页签内容完全重合，留着只会让人分不清"到底该看哪张"，故合并为一页。
> JSON 里有、配置表没登记的角色，现在只体现在「校验」页的提示 + 「有效词条待补」页的候选词条。

用法：
    python json_关键信息-豆包.py
    python json_关键信息-豆包.py 输入.json 输出.xlsx
================================================================================
"""
import importlib.util
import json
import os
import shutil
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
# HERE 是 code\ 目录；配置文件在 config\，数据与产物都在项目根目录
ROOT_DIR = os.path.dirname(HERE)


# ============================================================
# 一、复用第一个脚本的口径（避免两份代码各写一套映射）
# ============================================================
def load_helper():
    """
    把「json_全量提取-豆包.py」当模块加载进来。
    文件名带中文，不能直接 import，所以用 importlib 按路径加载。
    该脚本末尾有 if __name__ == "__main__" 保护，加载时不会跑主流程。
    """
    path = os.path.join(HERE, "json_全量提取-豆包.py")
    spec = importlib.util.spec_from_file_location("quanliang", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


q = load_helper()

# 从第一个脚本借用的东西（名字保持不变，方便对照查阅）
SLOT_CN = q.SLOT_CN          # 头部→头、躯干→衣服、脚部→鞋子、位面球→球、连结绳→绳子
SLOT_ZONE = q.SLOT_ZONE      # 前四外圈、后两内圈
stat_cn = q.stat_cn          # (stat, percent) → 项目词条中文
solve_rolls = q.solve_rolls  # 一件遗器的 4 条副词条 → 各投了几次（精确）
abbr_of = q.abbr_of          # 任意套装写法 → 项目主用简称
resolve_set = q.resolve_set  # 任意套装写法 → 项目全称
resolve_char = q.resolve_char            # 任意角色写法 → 项目统一名
build_char_resolver = q.build_char_resolver
build_set_alias_table = q.build_set_alias_table
load_config_table = q.load_config_table  # 静态配置表（遗器相关等）三档兜底读取
load_config_raw = q.load_config_raw      # 同上，但不把首行当表头（主词条分布要用）
CHAR_COLS = q.CHAR_COLS                  # 角色遗器表的 17 个标准列名
CONFIG_NAME = q.CONFIG_NAME              # 配置表文件名（项目配置-豆包.xlsx）

# 项目「角色遗器」表的列顺序，改这里就等于改输出表头。
# 前 16 列是 main.py 按位置读的，顺序千万别动；
# 第 17 列「权重」= index 16、第 18 列「角色星级」= index 17，两个都在末尾追加的列。
# 星级只在最后输出时给分析表加列用，不参与任何概率计算。
COLUMNS = [
    "角色名称", "遗器套装名称", "内圈套装名称", "有效词条数量", "有效词条",
    "衣服主词条", "鞋子主词条", "球主词条", "绳子主词条",
    "头", "手", "衣服", "鞋子", "球", "绳子",
    "词条数量", "权重", "角色星级",
]

# 六个部位，按 Excel 的 头/手/衣服/鞋子/球/绳子 顺序（顺序决定后面取数的位置）
SLOTS = ["头部", "手部", "躯干", "脚部", "位面球", "连结绳"]
SLOT_KEYS = ["头", "手", "衣服", "鞋子", "球", "绳子"]

# 角色数据：主用 characters.json，旧名 characters_大号.json 兼容一档（沿用第一个脚本的定位逻辑，
# 两个脚本共用同一套口径，避免改名时一处改了另一处没改）
JSON_PATH = q.JSON_PATH
OUT_PATH = os.path.join(ROOT_DIR, "星铁装备刷取-自动生成-豆包.xlsx")

# 权重已经从配置表里拆出来，单独一个表放根目录。
# 分工：有效词条 → config\项目配置-豆包.xlsx 的「角色有效词条」页；权重 → 这里的「角色权重」页。
WEIGHT_NAME = "角色权重.xlsx"
WEIGHT_PATH = os.path.join(ROOT_DIR, WEIGHT_NAME)

# 自动补登前的备份目录：只在第一次自动写入时存一份，方便回滚
BACKUP_DIR = os.path.join(os.path.dirname(ROOT_DIR), "bk", "备份-豆包", "08-自动补登前")


# ============================================================
# 二、从 JSON 里提炼「角色遗器」的一行
# ============================================================
def keep_old_wording(computed_f, computed_abbr, old_text, alias2full):
    """
    避免"换个写法"被当成数据变化。

    项目表里同一套遗器可能有好几种写法同时成立：全称、简称（「拾骨地、哀地」两个都认）、
    以及游戏内正确名和项目表的错字名。所以这里不比对字符串，而是：
        · 把原表那一格拆成几段，每段都用别名表解析成"项目全称"
        · 再看解析出来的这批全称，跟本次算出来的**是不是同一批套装**
    是同一批 → 沿用原表写法（连 2+2 的顺序都不动），让 diff 只反映真正的换装。

    computed_f     本次算出的套装项目全称，list（外圈可能 2 个）
    computed_abbr  对应的项目简称，list
    old_text       原表那一格的内容（可以是「A+B」，A/B 用任何写法都行）
    alias2full     任意写法 → 项目全称
    """
    old_text = "" if old_text is None else str(old_text).strip()
    fallback = "+".join(computed_abbr)
    if not old_text or old_text.lower() == "nan":
        return fallback
    parts = [p.strip() for p in old_text.split("+") if p.strip()]
    fulls = []
    for p in parts:
        f = resolve_set(p, alias2full)     # ← 关键：这里是"任意写法 → 全称"，不是精确匹配
        if f is None:
            return fallback                # 旧写法解析不出来 → 说明真的换了，用新算的
        fulls.append(f)
    if len(fulls) == len(computed_f) and set(fulls) == set(computed_f):
        return old_text                    # 同一批套装，只是写法/顺序不同，沿用原表
    return fallback


def split_outer_inner(pieces):
    """
    把 6 件遗器拆成外圈（头/手/衣服/鞋子）和内圈（球/绳子）。

    外圈可能不止一个套装：标准 4 件套是 1 个，2+2 是 2 个（如知更鸟「司铎+信使」）。
    内圈正常是 1 个，但白露是 1+1 散件（两个不同套装各 1 件），项目 eq2 只认单个套装，
    这种情况取「件数多的那个」，件数相同取球的那套，并在备注里写明。
    """
    outer, inner = [], []
    for pc in pieces:
        bucket = outer if SLOT_ZONE.get(pc["slot"]) == "外圈" else inner
        if pc["set"] not in bucket:
            bucket.append(pc["set"])

    # 内圈 1+1 散件：只能挑一个当 eq2
    inner_pick, inner_note = "", ""
    if len(inner) == 1:
        inner_pick = inner[0]
    elif len(inner) >= 2:
        cnt = {}
        for pc in pieces:
            if pc["set"] in inner:
                cnt[pc["set"]] = cnt.get(pc["set"], 0) + 1
        best = max(inner, key=lambda s: (cnt.get(s, 0), -inner.index(s)))
        inner_pick = best
        inner_note = ("内圈是散件(%s)，项目只支持单套装，暂按件数多的「%s」计"
                      % (" + ".join(inner), best))
    return outer, inner_pick, inner_note


def build_row(c, eff_list, weight, set_resolver, old=None):
    """
    一个 JSON 角色 →「角色遗器」一行的全部字段。

    参数：
        c            JSON 里的角色字典
        eff_list     该角色的有效词条列表（项目定义，来自配置表；新角色为空）
        weight       权重（来自角色权重.xlsx；新角色留空）
        set_resolver (别名表, 全称→简称) 二元组，来自第一个脚本，任何写法都能解析
        old          原表那一行（dict），没有则传 None

    返回 (行字典, 诊断信息字典, 数据问题列表)。
        · 诊断信息只用于「校验」页，不写进 17 列
        · 数据问题列表非空 = 这个角色**数据不足以参与计算**（件数不足 / 遗器未满级 /
          副词条数值不在档位表内），调用方要把它挪进「数据不完整」页，不要塞进「角色遗器」页。
          塞进去会让 main.py 把缺失部位当成"0 条有效"，算出来的概率会虚高。
          多账号场景下这类情况很常见（新号大量角色没装备或只穿了几件）。
    """
    pieces_list = c.get("relics", {}).get("pieces", [])
    pieces = {pc["slot"]: pc for pc in pieces_list}
    outer, inner_pick, inner_note = split_outer_inner(pieces_list)
    old = {} if old is None else old
    alias2full, full2abbr = set_resolver

    row = {"角色名称": "", "权重": weight}
    issues = []          # [(部位, 原因)]，空列表 = 数据完整

    # ---- 件数检查：没穿满 6 件就算不出整套提升概率 ----
    # （另一个账号里 17 个角色一件遗器都没穿、5 个只穿了几件，都走这一条）
    missing_slots = [s for s in SLOTS if s not in pieces]
    if missing_slots:
        issues.append(("、".join(SLOT_CN.get(s, s) for s in missing_slots),
                       "该角色只有 %d 件遗器，缺这 %d 个部位" % (len(pieces_list), len(missing_slots))))

    # ---- 套装 ----
    # 输出口径：一律用游戏内官方全称（JSON 里本来就是官方名，直接把原值写回）。
    #   · 外圈 2+2 用「A+B」连接（main.py 按「+」拆）
    #   · 不再套项目表的简称，也不需要"沿用原表写法"——统一正规名，diff 只反映真实换装
    # 识别方向仍然保留：不管原表写的是简称、旧错字名还是全称，都能解析回同一套（见 alias2full）。
    outer_f = [resolve_set(s, alias2full) or s for s in outer]
    inner_f = [resolve_set(inner_pick, alias2full) or inner_pick] if inner_pick else []
    row["遗器套装名称"] = "+".join(outer) if outer else ""
    row["内圈套装名称"] = inner_pick
    # （outer_f / inner_f 仍算出来，供下面的诊断备注判断"这套在项目表里认不认得"）

    # ---- 有效词条 ----
    eff = list(eff_list) if eff_list else []
    row["有效词条数量"] = len(eff) if eff else ""
    row["有效词条"] = "、".join(eff) if eff else ""
    # 角色星级（4/5）：配置表人工填的，只透传给分析表当筛选列，不参与计算
    # （old 在函数开头已被统一成 dict，没填时 get 返回空串）
    sv = old.get("角色星级", "")
    row["角色星级"] = "" if sv is None or (isinstance(sv, float) and pd.isna(sv)) else str(sv).strip()

    # ---- 四个可变部位的主词条 ----
    for slot, col in (("躯干", "衣服主词条"), ("脚部", "鞋子主词条"),
                      ("位面球", "球主词条"), ("连结绳", "绳子主词条")):
        pc = pieces.get(slot)
        row[col] = stat_cn(pc["main_stat"]["stat"], pc["main_stat"]["percent"]) if pc else ""

    # ---- 六个部位的有效强化次数 ----
    # 口径：该部位【副词条】里属于有效词条的那几条，各自投了几次，加起来
    #      （主词条即使同名也不算，因为项目就是这么定义的）
    diag = {"次数明细": {}, "还原失败": []}
    for slot, col in zip(SLOTS, SLOT_KEYS):
        pc = pieces.get(slot)
        if pc is None:
            row[col] = ""
            continue
        lv = pc.get("level")
        # 未满级遗器：副词条总次数只有 8~9 这个前提不成立，换算不出"投了几次"
        if isinstance(lv, (int, float)) and lv < 15:
            row[col] = ""
            diag["还原失败"].append(SLOT_CN.get(slot, slot))
            issues.append((SLOT_CN.get(slot, slot), "遗器只有 +%g 级（未满级），副词条次数无法还原" % lv))
            continue
        sols, note = solve_rolls(pc["sub_stats"])
        if not sols:
            row[col] = ""
            diag["还原失败"].append(SLOT_CN.get(slot, slot))
            bada = []
            for s in pc["sub_stats"]:
                bada.append("%s%s" % (stat_cn(s["stat"], s["percent"]) or s["stat"],
                                      ("%g%%" % s["value"]) if s.get("percent") else ("%g" % s["value"])))
            issues.append((SLOT_CN.get(slot, slot), "副词条数值（%s）不在档位表内，只算不出强化次数" % "、".join(bada)))
            continue
        # 多解时取"有效次数最大"的那个（与手数习惯一致：能确认命中一次才算）
        if eff:
            def eff_of(kk):
                return sum(k for sub, k in zip(pc["sub_stats"], kk)
                           if stat_cn(sub["stat"], sub["percent"]) in eff)
            rolls = max(sols, key=eff_of)
        else:
            rolls = sols[0]
        k = sum(kk for sub, kk in zip(pc["sub_stats"], rolls)
                if stat_cn(sub["stat"], sub["percent"]) in eff)
        # 没有有效词条定义时留空（而不是 0）：0 会被误读成"一件都没命中"
        row[col] = int(k) if eff else ""
        diag["次数明细"][SLOT_CN.get(slot, slot)] = int(k)

    # ---- 词条数量：只把真正算出来的数字相加 ----
    # 这里曾经写成 sum(row[c] for c in SLOT_KEYS)，一旦有部位是 ""（缺件/算不出）就
    # TypeError: int + str，整个导出直接崩。改成"只累加 int"，空值当 0。
    nums = [row[c] for c in SLOT_KEYS if isinstance(row.get(c), int)]
    row["词条数量"] = sum(nums) if eff else ""

    diag["内圈备注"] = inner_note
    return row, diag, issues


# ============================================================
# 三、把 76 个角色全部转出来，再拆成「原表 55」与「新增 22」
# ============================================================
def build_all(chars, char_df, set_resolver):
    """
    返回 (old_rows, new_rows, bad_rows, diag_map, legacy, no_eff)

    old_rows：配置表名单内、数据完整、**且有效词条不为空**的角色 → 进「角色遗器」页，可直接喂 main.py
    bad_rows：配置表有、JSON 也有，但数据不完整（缺件/未满级/数值在档位表外）
               → 进「数据不完整」页。{角色名称, 部位, 原因}
    new_rows：JSON 有、配置表没有的角色 → 进「有效词条待补」页
    no_eff：  配置表有、JSON 也有、装备也齐，但「有效词条」这一列是空的
               → 同样进「有效词条待补」页。这类角色算出来的每个部位都是 0 条有效，
               塞进主表会让 main.py 把"没定义"当成"一条都没命中"，概率虚高。
    legacy：  配置表有、JSON 没有的角色（如「欢愉主」），本次不参与计算
    顺序：先配置表顺序，再新角色按 JSON 导出顺序。

    名字识别是多写法的：不管是 JSON 官方全名（银狼LV.999）、项目简称（银狼999），
    还是归一化后能对上的写法（缇宝/提宝），都能解析到同一个角色。
    """
    # ---- 角色名解析器：任意写法 → 项目统一名（配置表写法）----
    json_names = [c["name"] for c in chars]
    excel_names = [str(x).strip() for x in char_df["角色名称"].dropna()]
    char_resolver = build_char_resolver(json_names, excel_names)

    # 反向索引：项目统一名 → JSON 角色（万一两个 JSON 名映到同一个，取第一个）
    by_canon = {}
    for c in chars:
        canon = resolve_char(c["name"], char_resolver) or c["name"]
        by_canon.setdefault(canon, c)

    # ---- 配置表名单（逐行）----
    old_rows, diag_map, bad_rows, no_eff = [], {}, [], []
    legacy = []          # 配置表有、JSON 没覆盖的角色
    for _, r in char_df.iterrows():
        ename = str(r["角色名称"]).strip()
        # 表里写的是哪个名字都能认出来：先用解析器统一，再拿统一名查 JSON
        canon = resolve_char(ename, char_resolver) or ename
        c = by_canon.get(canon)
        if c is None:
            # 配置表有、JSON 没覆盖（如「欢愉主」：JSON 只导了当前命途「记忆」那一份）。
            # 注意：正常流程里 drop_uncovered 已经先把这类剔掉了，这里是兜底分支，
            # 不再往主表塞空行（塞了 main.py 会拿到一行空数据）。
            legacy.append(ename)
            continue
        eff_list = [t for t in str(r["有效词条"]).split("、") if t and t != "nan"]
        weight = r["权重"] if pd.notna(r["权重"]) else ""
        # 有效词条是空的：装备再齐也算不出"哪几条算有效"，
        # 这里直接不产出行，改记进 no_eff 列表 → 进「有效词条待补」页。
        # 不这么做的话，六个部位会全算成 0，main.py 会把它当成"一条都没命中"，概率虚高。
        if not eff_list:
            no_eff.append(ename)
            continue
        row, diag, issues = build_row(c, eff_list, weight, set_resolver, old=r)
        row["角色名称"] = ename
        diag_map[ename] = diag
        if issues:
            # 数据不够算：不进主表，原因一条一行写进「数据不完整」页
            for slot, reason in issues:
                bad_rows.append({"角色名称": ename, "部位": slot, "原因": reason,
                                 "件数": "%d/6" % len(c.get("relics", {}).get("pieces", []))})
            continue
        old_rows.append(row)

    # ---- 新增角色（统一名不在配置表名单里的）----
    new_rows = []
    used = {resolve_char(str(n).strip(), char_resolver) or str(n).strip()
            for n in char_df["角色名称"].dropna()}
    for c in chars:
        canon = resolve_char(c["name"], char_resolver) or c["name"]
        if canon in used:
            continue
        row, diag, issues = build_row(c, [], "", set_resolver)
        row["角色名称"] = c["name"]
        new_rows.append(row)
        diag_map[c["name"]] = diag

    return old_rows, new_rows, bad_rows, diag_map, legacy, no_eff


def drop_uncovered(char_df, chars):
    """
    挑掉「配置表里有、但 JSON 里没有数据」的角色。

    典型例子是「欢愉主」：JSON 只导了开拓者当前命途（记忆）那一份，
    所以这个角色的套装、部位、次数全都算不出来。留着会输出一行空数据，
    main.py 读「遗器套装名称」为 NaN 时会在字符串判断那里直接报错。

    返回 (过滤后的表, 被跳掉的角色名列表)
    """
    json_names = [c["name"] for c in chars]
    cfgs = [str(x).strip() for x in char_df["角色名称"]]
    res = build_char_resolver(json_names, cfgs)      # 多写法互通，免得因名字写法差异误判
    covered = {resolve_char(j, res) or j for j in json_names}

    keep_idx, dropped = [], []
    for idx, r in char_df.iterrows():
        name = str(r["角色名称"]).strip()
        canon = resolve_char(name, res) or name
        if canon in covered:
            keep_idx.append(idx)
        else:
            dropped.append(name)
    return char_df.loc[keep_idx].reset_index(drop=True), dropped


def to_frame(rows):
    """按 17 列顺序转成 DataFrame"""
    return pd.DataFrame([{c: r.get(c, "") for c in COLUMNS} for r in rows], columns=COLUMNS)


def merge_config_chars(char_df, chars):
    """
    【已停用 · 2026-11-19】手填表（星铁装备刷取.xlsx）退役后不再调用本函数，留作参考。

    原来它干的事：把手填表和配置表「角色有效词条」合成一张表（并集），
    修掉"新手填漏了就永远不统计"的问题。现在角色名单只来自配置表，没有并集可做。

    规则：
        · 两边都有的角色 → 用手填表那一行（套装、主词条、次数仍以手填为准，保持原行为）
        · 只有配置表有的角色 → 补一行"最简行"：套装/主词条/六部位次数全由 JSON 算，
          只从配置表拿「有效词条」「权重」这两列（这两列算不出来，必须人工判断）

    返回 (合并后的表, 本次补进来的角色名列表)
    """
    cfg = load_char_table_from_config()
    if cfg is None or cfg.empty:
        return char_df, []

    # 多写法归一（官方名 / 项目名 / 别名 都能互认），避免同一个角色被当成两个
    res = build_char_resolver(
        [c["name"] for c in chars],
        [str(x).strip() for x in char_df["角色名称"]] +
        [str(x).strip() for x in cfg["角色名称"]])

    have = set()
    for n in char_df["角色名称"]:
        n = str(n).strip()
        have.add(resolve_char(n, res) or n)

    add_rows, added = [], []
    for _, r in cfg.iterrows():
        name = str(r["角色名称"]).strip()
        if not name or name == "nan":
            continue
        canon = resolve_char(name, res) or name
        if canon in have:
            continue                     # 手填表已经有了，不重复
        add_rows.append(r.to_dict())
        added.append(name)
        have.add(canon)

    if not add_rows:
        return char_df, []
    merged = pd.concat([char_df, pd.DataFrame(add_rows)], ignore_index=True)
    return merged, added


def load_character_weights():
    """
    读根目录「角色权重.xlsx」的「角色权重」页，返回 {角色名: 权重(float)}。

    权重是独立的一件人工判断（这个角色值多少分），和有效词条分开放：
        · 有效词条 → config\\项目配置-豆包.xlsx 的「角色有效词条」页
        · 权重     → 根目录 角色权重.xlsx

    这样只调权重不用碰配置表，也不会因为改权重把有效词条页改乱。
    文件缺失 / 页签缺失 / 单元格不是数字 → 都不崩，只是该角色按"没登记"处理（下游按 1.0）。
    """
    if not os.path.exists(WEIGHT_PATH):
        print("!! 没找到「%s」：本次全部角色按权重 1.0 处理（加权页会退化成未加权）" % WEIGHT_NAME)
        return {}
    try:
        df = pd.read_excel(WEIGHT_PATH, sheet_name="角色权重")
    except Exception as e:
        print("!! 读「%s」失败（%s），权重按 1.0 处理" % (WEIGHT_NAME, e))
        return {}

    out = {}
    for _, r in df.iterrows():
        name = str(r.get("角色名称", "")).strip()
        if not name or name.lower() == "nan":
            continue
        w = r.get("权重", "")
        if w is None or (isinstance(w, float) and pd.isna(w)) or str(w).strip() == "":
            continue                      # 留空 = 不登记，不是错误
        try:
            out[name] = float(w)
        except (TypeError, ValueError):
            print("!! 「%s」的权重不是数字：%r，已忽略（按 1.0 处理）" % (name, w))
    return out


def load_char_table_from_config():
    """
    从配置表「角色有效词条」+ 权重表，造一张"最简角色表"（项目的唯一角色名单来源）。

    填 角色名称 / 角色星级 / 有效词条 / 权重 四列，其余列一律给空字符串
    （不是 NaN——空串走的是"没有旧值"分支，NaN 会让往下比对时出岔子）。
    这样 build_all 会按这份名单逐行生成，套装 / 主词条 / 六部位次数全由 JSON 算。

    2026-11-19 起手填表退役：不再有"手填表优先"那套合并逻辑，
    有效词条来自配置表、权重来自 角色权重.xlsx，两边都没有才算没登记。
    2026-11-20 起配置表多了「角色星级」列（人工填 4/5），只用来给分析表加筛选列，
    不参与计算；这一列不存在时全部给空串，不影响老配置表使用。
    """
    df, src = load_config_table("角色有效词条", ["角色名称", "有效词条", "角色星级"],
                                legacy_xlsx=os.path.join(ROOT_DIR, "__不存在__.xlsx"))
    if df.empty:
        return None

    # 权重表与配置表的名字可能有写法差异（官方全名 / 简称），用解析器统一到配置表的写法
    weights = load_character_weights()
    cfg_names = [str(x).strip() for x in df["角色名称"] if str(x).strip()]
    res = build_char_resolver([str(k).strip() for k in weights], cfg_names)
    wmap = {}
    for k, v in weights.items():
        canon = resolve_char(str(k).strip(), res) or str(k).strip()
        wmap.setdefault(canon, v)

    has_star = "角色星级" in df.columns

    rows = []
    for _, r in df.iterrows():
        name = str(r["角色名称"]).strip()
        if not name or name.lower() == "nan":
            continue
        eff = "" if pd.isna(r["有效词条"]) else str(r["有效词条"]).strip()
        w = wmap.get(name, "")
        # 星级：人工填的 4/5。填成"5星""五星"这种也照原样透传，不在这里猜
        star = ""
        if has_star and pd.notna(r["角色星级"]):
            star = str(r["角色星级"]).strip()
            if star.endswith(".0"):
                star = star[:-2]
        rows.append({"角色名称": name, "角色星级": star, "有效词条": eff, "权重": w})
    out = pd.DataFrame([{c: d.get(c, "") for c in COLUMNS} for d in rows], columns=COLUMNS)
    return out


# ============================================================
# 四、自动补登：识别到新角色就写回配置表 / 权重表
# ============================================================
def find_unregistered(chars, char_df):
    """
    找出「JSON 里有、配置表「角色有效词条」页没登记」的角色名（用 JSON 侧写法）。

    比的是**归一化后的名字**，不是字符串：配置表写「记忆主」、JSON 写「开拓者·记忆」，
    那是同一个角色，不该再补登一次（否则同一个角色会变成两行）。
    """
    json_names = [c["name"] for c in chars]
    cfg_names = [str(x).strip() for x in char_df["角色名称"].dropna()] if char_df is not None else []
    res = build_char_resolver(json_names, cfg_names)
    registered = {resolve_char(n, res) or n for n in cfg_names}

    out = []
    for c in chars:
        nm = str(c.get("name", "")).strip()
        if not nm:
            continue
        if (resolve_char(nm, res) or nm) not in registered:
            out.append(nm)
    return out


def _backup_once(path):
    """把文件复制到 08-自动补登前\，只复制第一次（同名已存在就跳过），避免每次跑都堆备份。"""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        dst = os.path.join(BACKUP_DIR, os.path.basename(path))
        if not os.path.exists(dst):
            shutil.copy2(path, dst)
    except Exception as e:
        print("!! 备份 %s 失败（%s），仍继续写入" % (os.path.basename(path), e))


def _append_rows_safely(path, sheet_name, new_rows):
    """
    往某个 xlsx 的指定页末尾追加若干行，**只动这一页、只加行尾**，其它页签一个单元格都不碰。

    new_rows: [{"列名": 值, ...}, ...]。按**表头名**定位列，不按位置——
    配置表的列顺序会变（比如「角色星级」插在第二列），按位置写必然串列。

    安全措施：
        · 先备份（只第一次）
        · 先存成 .tmp 再 os.replace 替换原文件——中途出错也不会留下半张表
    返回实际追加的行数。
    """
    import openpyxl
    if not new_rows or not os.path.exists(path):
        return 0
    _backup_once(path)

    wb = openpyxl.load_workbook(path)
    if sheet_name not in wb.sheetnames:
        print("!! %s 里没有「%s」页，跳过追加" % (os.path.basename(path), sheet_name))
        return 0
    ws = wb[sheet_name]

    # 表头名 → 列号（第一行）
    header = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v not in (None, ""):
            header[str(v).strip()] = c

    # 找到"最后一行真正有内容"的位置：表格里可能有空行或格式残留，
    # 直接用 max_row 会把新行写到很下面去。
    last = 1
    for r_idx in range(ws.max_row, 0, -1):
        if any(ws.cell(row=r_idx, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            last = r_idx
            break

    written = 0
    for i, d in enumerate(new_rows, start=1):
        for k, v in d.items():
            col = header.get(k)
            if col is None:           # 表里没有这一列就跳过，不硬塞到别的列去
                continue
            ws.cell(row=last + i, column=col, value=v)
        written += 1

    tmp = path + ".tmp"
    wb.save(tmp)
    wb.close()
    os.replace(tmp, path)          # 原子替换：要么是旧表，要么是完整的新表
    return written


def append_new_chars_to_config(new_names):
    """
    把识别到的新角色写进两张表，让人一打开表就看见"这俩还没填词条"：

        config\\项目配置-豆包.xlsx  「角色有效词条」页
            角色名称 = 官方名，有效词条 = 留空，备注 = 自动补登 + 待填有效词条
        <根目录>\\角色权重.xlsx     「角色权重」页
            角色名称 = 官方名，权重 = 1（新角色默认等权，之后按强度自己改）

    为什么词条要留空而不是瞎填一个：填错比空着更糟——空着会被挡在「角色遗器」页外，
    只看得到"待填"提示；填错则会静默算出一堆错误概率。

    只追加、只写行尾，已有角色不会重复写，其它页签一律不动。
    返回 (写进配置表的角色, 写进权重表的角色)。
    """
    if not new_names:
        return [], []
    today = __import__("datetime").date.today().isoformat()
    cfg_path = os.path.join(ROOT_DIR, "config", CONFIG_NAME)

    # ---- ① 配置表「角色有效词条」：角色名称 / 角色星级(空，等人工填) / 有效词条(空) / 备注 ----
    # 按列名写，所以配置表列顺序怎么变都不怕（「角色星级」是 2026-11-20 才插进来的第二列）
    cfg_rows, cfg_added = [], []
    for nm in new_names:
        cfg_rows.append({
            "角色名称": nm,
            "角色星级": "",          # 新角色星级也留空，和词条一起等人补
            "有效词条": "",
            "备注": "自动补登 %s：待填角色星级(4/5) + 有效词条(3~4 条，顿号分隔)" % today,
        })
        cfg_added.append(nm)
    n = _append_rows_safely(cfg_path, "角色有效词条", cfg_rows)
    if not n:
        cfg_added = []

    # ---- ② 角色权重.xlsx「角色权重」：角色名称 / 权重 1 / 备注 ----
    # 先看这角色是不是已经有权重了（有人可能只填了权重、忘了填词条），有就不重复写
    w_added, w_rows = [], []
    if os.path.exists(WEIGHT_PATH):
        try:
            wdf = pd.read_excel(WEIGHT_PATH, sheet_name="角色权重")
            have = {str(x).strip() for x in wdf["角色名称"].dropna()}
        except Exception:
            have = set()
        for nm in new_names:
            if nm in have:
                continue
            w_rows.append({"角色名称": nm, "权重": 1,
                           "备注": "自动补登 %s：新角色默认等权 1，按强度自行调整" % today})
            w_added.append(nm)
        if w_rows:
            n2 = _append_rows_safely(WEIGHT_PATH, "角色权重", w_rows)
            if not n2:
                w_added = []
    return cfg_added, w_added


# ============================================================
# 五、辅助页签
# ============================================================
def build_pending_sheet(chars, new_names, no_eff=None):
    """
    「有效词条待补」：这些角色还没有"哪些副词条算有效"的定义，光看六件数字没法勾。
    这里把每个角色六件上出现过的副词条按出现次数排好，照着勾就行。

    new_names：JSON 有、配置表原本没登记的角色（本次已自动补登进配置表，词条留空）
    no_eff  ：配置表登记了、但「有效词条」这一列还是空的角色
    两类都进这一页，因为用户的动作完全一样：打开配置表，在角色那一行填上有效词条。
    """
    rows = []
    by_json = {c["name"]: c for c in chars}
    status_of = {n: "新角色：已自动补登进配置表，等你填有效词条" for n in (new_names or [])}
    for n in (no_eff or []):
        status_of.setdefault(n, "已登记，但「有效词条」还是空的")
    for name, status in status_of.items():
        c = by_json.get(name)
        if not c:
            continue
        cnt = {}
        for pc in c["relics"]["pieces"]:
            for s in pc["sub_stats"]:
                cn = stat_cn(s["stat"], s["percent"])
                if cn:
                    cnt[cn] = cnt.get(cn, 0) + 1
        # 按出现件数从多到少，出现多的更可能是这个角色的有效词条
        ordered = sorted(cnt.items(), key=lambda x: (-x[1], -(q.RES_WEIGHT.get(x[0]) or 0)))
        rows.append({
            "角色名称": name,
            "状态": status,
            "命途": c.get("path", ""),
            "属性": c.get("element", ""),
            "六件上出现过的副词条(按件数降序)": "、".join(k for k, _ in ordered),
            "各类出现件数": "、".join("%s%d" % (k, v) for k, v in ordered),
            "建议有效词条(仅供参考)": "、".join(k for k, _ in ordered[:4]),
        })
    # 没有待补角色时也要保出表头：否则这一页会变成"0 行 0 列"的空表，
    # 打开来看不出这张表是干什么的（现在配置表已覆盖全部角色，这页大概率是空的）
    cols = ["角色名称", "状态", "命途", "属性", "六件上出现过的副词条(按件数降序)",
            "各类出现件数", "建议有效词条(仅供参考)"]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows, columns=cols)


def build_bad_sheet(bad_rows):
    """
    「数据不完整」：配置表里有、JSON 里也有这个角色，但遗器数据不够算。

    换账号时这一页最有用——新号常见三种情况，都会被列在这里：
        ① 角色一件遗器都没穿 / 只穿了几件 → 缺部位
        ② 遗器没强化满（+0 ~ +14）→ 副词条总次数不是 8~9，算不出"投了几次"
        ③ 副词条数值不在档位表内 → 多半是导出数据异常，需要重导或手动核对

    这些角色**不会**进「角色遗器」页：主表里塞 0 会让 main.py 把缺失部位当成
    "0 条有效词条"，算出来的提升概率虚高。装备补齐后重跑 01 即可自动回到主表。
    """
    cols = ["角色名称", "问题部位", "件数", "原因", "怎么办"]
    if not bad_rows:
        return pd.DataFrame(columns=cols)
    rows = []
    seen = set()
    for b in bad_rows:
        key = (b["角色名称"], b["部位"], b["原因"])
        if key in seen:
            continue
        seen.add(key)
        advice = "补齐这件的遗器后重跑 01-一键导出.bat"
        if "档位表" in b["原因"]:
            advice = "数值异常，建议重新导出 JSON；仍不行就手动核对该件副词条"
        elif "未满级" in b["原因"]:
            advice = "把这件强化到 +15 后重跑 01-一键导出.bat"
        elif "只有" in b["原因"]:
            advice = "给这个角色穿满 6 件遗器后重跑 01-一键导出.bat"
        rows.append({"角色名称": b["角色名称"], "问题部位": b["部位"],
                     "件数": b.get("件数", ""), "原因": b["原因"], "怎么办": advice})
    return pd.DataFrame(rows, columns=cols)


def build_check_sheet(char_df, chars, weights, dropped, char_resolver, bad=None):
    """
    「校验」：查"数据齐不齐"，不再和手填表逐格比对（手填表退役后没有原值可比了）。

    查六件事，每件一行，只列有问题的角色：
        ① 配置表登记了、但 JSON 里没数据 → 套装/部位/次数都算不出来，只能跳过
        ② JSON 有、但配置表没登记有效词条 → 会被算成 0 条有效词条，统计结果偏低
        ③ 有效词条为空或只有 1 条 → 基本算不出东西，多半是没填
        ④ 权重表里没这个角色 → 加权时按 1.0 处理，加权页会失真
        ⑤ JSON 有、但装备凑不齐（缺件/未满级/数值异常）→ 已剔到「数据不完整」页，不参与计算
        ⑥ 配置表没填角色星级（4/5）→ 不影响算数，但分析表没法按星级筛选

    名字比对必须走 char_resolver（配置表里是"银狼999"、JSON 里是"银狼LV.999"，
    直接比字符串会把同一个角色当成两个，报一堆假差异）。

    bad：build_all 返回的 bad_rows（可空）。传了就按角色合并成第 ⑤ 类提示。

    返回一张 3 列的表（角色名称 / 检查项 / 结论）。全部正常时返回一行"无异常"。
    """
    rows = []
    cfg_names = set()
    for _, r in char_df.iterrows():
        name = str(r["角色名称"]).strip()
        if not name or name.lower() == "nan":
            continue
        cfg_names.add(name)

        eff = "" if pd.isna(r.get("有效词条")) else str(r.get("有效词条")).strip()
        if not eff or eff.lower() == "nan":
            rows.append({"角色名称": name, "检查项": "有效词条",
                         "结论": "配置表里没填有效词条，有效次数会算成 0；请在「角色有效词条」页补上"})
        elif len([t for t in eff.split("、") if t.strip()]) <= 1:
            rows.append({"角色名称": name, "检查项": "有效词条",
                         "结论": "只登记了 1 条有效词条（%s），统计会明显偏低；一般要 3~4 条" % eff})

        # 角色星级（4/5）：只用来给分析表加一列方便筛选，不参与任何概率计算。
        # 空着不影响算数，但筛"只看五星"时会漏人，所以顺手提示一句。
        st = r.get("角色星级", "")
        if st is None or (isinstance(st, float) and pd.isna(st)) or str(st).strip() in ("", "nan"):
            rows.append({"角色名称": name, "检查项": "角色星级",
                         "结论": "配置表「角色有效词条」页没填星级（4 或 5）：不影响概率计算，"
                                "但分析表里没法按星级筛选，建议补上"})

        # 权重：load_char_table_from_config 已经把权重表的名字归一化后填进这一列，
        # 空着就说明「角色权重.xlsx」里没有这个角色
        w = r.get("权重")
        if weights is not None and (w is None or (isinstance(w, float) and pd.isna(w)) or str(w).strip() in ("", "nan")):
            rows.append({"角色名称": name, "检查项": "权重",
                         "结论": "「%s」里没有这个角色，加权时按 1.0 处理" % WEIGHT_NAME})

    # ② JSON 有、配置表没登记的角色（不会出现在「角色遗器」页，只在待补/校验页提示）
    for c in chars:
        nm = str(c.get("name", "")).strip()
        if not nm:
            continue
        canon = resolve_char(nm, char_resolver) or nm
        if canon in cfg_names:
            continue          # 同一个角色，只是写法不同
        rows.append({"角色名称": nm, "检查项": "名单",
                     "结论": "JSON 里有、配置表「角色有效词条」页没登记 → 不会进「角色遗器」页；"
                            "去配置表加一行即可纳入统计"})

    # ① 配置表有、JSON 无（已被 drop_uncovered 剔除）
    for nm in dropped:
        rows.append({"角色名称": nm, "检查项": "名单",
                     "结论": "配置表登记了，但 JSON 里没有该角色的数据（只导了当前命途/未导出）→ 本次跳过"})

    # ⑤ 装备凑不齐的角色：bad_rows 是按 (角色, 部位, 原因) 展开的，这里按角色合并成一条，
    # 免得一个角色 6 个部位各来一行把校验页刷满。明细统一去「数据不完整」页看。
    if bad is not None and len(bad):
        # bad 可能是 build_bad_sheet 出来的 DataFrame，也可能是原始的 bad_rows 列表，两种都收
        recs = bad.to_dict("records") if isinstance(bad, pd.DataFrame) else bad
        by_char = {}
        for b in recs:
            by_char.setdefault(str(b["角色名称"]), []).append(str(b["问题部位"]))
        for nm, slots in by_char.items():
            parts = "、".join(slots)
            rows.append({"角色名称": nm, "检查项": "装备",
                         "结论": "装备不完整/无法还原（%s），本次未参与计算；"
                                "明细与处理办法见「数据不完整」页" % parts})

    if not rows:
        return pd.DataFrame([{"角色名称": "-", "检查项": "-",
                              "结论": "无异常：配置表名单、有效词条、权重都对得上 JSON"}])
    return pd.DataFrame(rows, columns=["角色名称", "检查项", "结论"])


def build_readme_sheet(data, n_old, n_new, n_bad=0, n_loss=0, n_noeff=0, n_added=0):
    """「生成说明」：写清口径，方便以后自己回看"""
    rows = [
        ("数据来源", "characters.json（三月七小助手导出）", data.get("exported_at", "")),
        ("角色覆盖", "JSON 角色数 / 进主表 / 本次自动补登 / 待填词条 / 数据不完整",
         "%d / %d / %d / %d / %d" % (len(data["characters"]), n_old, n_added, n_noeff, n_loss)),
        ("角色名单来自", "config「角色有效词条」页", "JSON 里没有该角色数据的会被跳过，见「校验」页"),
        ("新角色怎么办", "识别到 JSON 里有、配置表没登记的角色，脚本会自动补登",
         "配置表里「有效词条」留空、%s 里权重填 1.0，并写备注说明；你再手填词条即可" % WEIGHT_NAME),
        ("页签角色遗器", "配置表登记、有效词条已填、且遗器数据完整的角色",
         "可直接喂 main.py；表末新增「角色星级」列，只供分析表筛选，不参与计算"),
        ("页签有效词条待补", "新补登的角色 + 配置表里有行但词条还为空的角色",
         "同角色六件上出现过的副词条按件数排好，照着勾 3~4 条即可"),
        ("页签数据不完整", "JSON 里有、但装备凑不齐（缺件/未满级/数值异常）的角色，逐条列原因",
         "不参与计算；补齐后重跑 01 即可自动回到「角色遗器」"),
        ("没有别的角色表", "本文件只有这一张角色表（原「角色遗器-全76」页已于 2026-11-19 删除）",
         "未登记 / 词条为空的角色只在「有效词条待补」「校验」两页出现"),
        # 注意：这一格别以 "=" 开头——Excel 会当成公式，读回来就变成空值
        ("有效词条数量", "有效词条列表的个数（一般是 3 或 4）", "取自配置表，不再用手填表"),
        ("六部位数字", "该部位【副词条】里属于有效词条的那几条，各自投了几次之和", "0~9，满值 9"),
        ("词条数量", "六部位数字之和", "var 各部位完成度的分母就按这个算"),
        ("次数怎么来的", "档位表反解：每条副词条数值 → 投了几次", "每件合计必为 8 或 9"),
        ("主词条", "取 JSON 里该部位的主词条", "衣服/鞋子/球/绳子 四列"),
        ("外圈套装", "头手衣服鞋子用到的套装，2+2 用「+」连接", "如「司铎+信使」"),
        ("内圈套装", "球绳子用到的套装；1+1 散件取件数多的那套", "备注写在「校验」页"),
        ("权重", "取自根目录「%s」" % WEIGHT_NAME, "表里没登记的按 1.0 处理"),
        ("角色星级", "配置表「角色有效词条」页的「角色星级」列（人工填 4 或 5）",
         "只带到分析表的角色层级页签做筛选，不参与任何概率计算；空着也不影响算数"),
        ("替换方法", "「角色遗器」整页即项目输入表",
         "遗器相关 / 主词条分布 / 分解阈值 三页来自 config，本文件已自带"),
        ("直接分析", "python main.py 星铁装备刷取-自动生成-豆包.xlsx", "不必再把「角色遗器」手工粘回原输入表"),
    ]
    return pd.DataFrame(rows, columns=["项目", "口径", "备注"])


# ============================================================
# 五、主流程
# ============================================================
def main():
    # 参数：第 1 个 = 角色数据 JSON，第 2 个 = 输出表。手填表已退役，不再接受"输入表"参数。
    json_path = sys.argv[1] if len(sys.argv) > 1 else JSON_PATH
    out_path = sys.argv[2] if len(sys.argv) > 2 else OUT_PATH

    # 走共享的 load_json：它会先把开拓者（is_trailblazer=True）的玩家昵称
    # 改写成「开拓者·命途」，两个脚本口径一致，不会一个认「陆遥」一个不认。
    data = q.load_json(json_path)
    chars = data["characters"]

    # 角色名单的唯一来源 = 配置表「角色有效词条」页（权重另取自 角色权重.xlsx）。
    # 手填表退役后不再有"手填优先"那套合并逻辑，也就不存在"漏填就永远不统计"的问题。
    char_df = load_char_table_from_config()
    dropped = []
    if char_df is None or char_df.empty:
        print("!! 配置表里没有「角色有效词条」页（或整页为空）："
              "这次生成不出任何角色行，请先去配置表补角色；"
              "「遗器相关」「主词条分布」「分解阈值」几页仍会照常带出")
        char_df = pd.DataFrame(columns=COLUMNS)
    else:
        # 配置表里的角色未必都在 JSON 里（如「欢愉主」：JSON 只导了当前命途那一份）。
        # 这类角色算不出套装和部位，硬留着会生成一行空数据（main.py 读到会报错），先挑掉。
        char_df, dropped = drop_uncovered(char_df, chars)
        print("＋ 从配置表读到 %d 个角色定义（权重取自 %s）" % (len(char_df), WEIGHT_NAME))

    # 「遗器相关」已搬到配置表：配置表优先，旧手填表兜底，都没有就给空表
    set_df, set_src = load_config_table("遗器相关", ["遗器全称", "刷取位置", "简称1"])
    if set_src and os.path.basename(set_src) != CONFIG_NAME:
        print("!! 「遗器相关」目前仍从「%s」读，建议搬进「%s」"
              % (os.path.basename(set_src), CONFIG_NAME))
    # 套装别名表：项目全称、各简称、游戏内正确名、项目表错字名 全部登记，任意写法都能解析
    alias2full, full2abbr = build_set_alias_table(set_df)
    set_resolver = (alias2full, full2abbr)

    # ---- 自动补登新角色 ----
    # 识别到「JSON 有、配置表没登记」的角色，就写进配置表（有效词条留空）和权重表（权重 1）。
    # 目的是让用户下次打开配置表时能直接看到"这两个角色还没有有效词条"，而不是只在生成表里看到一行提示。
    new_names = find_unregistered(chars, char_df)
    cfg_added, w_added = append_new_chars_to_config(new_names)
    if cfg_added:
        print("＋ 已在配置表补登 %d 个新角色（有效词条留空）：%s" % (len(cfg_added), "、".join(cfg_added)))
        # 补登后重新读一遍配置表：这些角色会以"已登记、但有效词条为空"的身份
        # 落进「有效词条待补」页（而不是当成未登记角色被忽略），和用户后来手动加行的效果一致。
        char_df = load_char_table_from_config()
        if char_df is not None and not char_df.empty:
            char_df, dropped = drop_uncovered(char_df, chars)
    if w_added:
        print("＋ 已在 %s 补登 %d 个角色，权重默认 1.0：%s"
              % (WEIGHT_NAME, len(w_added), "、".join(w_added)))
    if cfg_added or w_added:
        print("   下一步：打开「%s」的「角色有效词条」页，给这些角色填 3~4 条有效词条"
              "（候选词条见生成表「有效词条待补」页），然后重跑 01" % CONFIG_NAME)

    old_rows, new_rows, bad_rows, diag_map, legacy, no_eff = build_all(chars, char_df, set_resolver)

    # 「角色遗器」页里应该有、实际没有的角色（配置表有 + JSON 有，但数据不完整被剔）
    lost = [r["角色名称"] for r in bad_rows]

    df_old = to_frame(old_rows)
    # 说明：原来的「角色遗器-全76」页已删除（2026-11-19）。
    # 角色名单统一由配置表决定、配置表又覆盖了 JSON 全部角色，两个页签内容完全重合，
    # 留着只会让人分不清该看哪张。new_rows（JSON 有、配置表没登记的角色）现在只用于
    # 「有效词条待补」页和「校验」页的提示，不再单独出一页角色表。
    #
    # 「有效词条待补」页收两类人：① 本次新识别、已自动补登的角色；② 配置表登记了但词条为空的角色。
    # 用户在这两类角色上的动作完全一样：打开配置表填词条。
    pending = build_pending_sheet(chars, [r["角色名称"] for r in new_rows], no_eff)
    bad = build_bad_sheet(bad_rows)

    # 角色解析器：JSON 写法 ↔ 配置表写法。校验页要比"是不是同一个角色"，
    # 必须在它之前建好（比如配置表写"银狼999"、JSON 写"银狼LV.999"，直接比字符串会误报）
    char_resolver = build_char_resolver([c["name"] for c in chars],
                                        [str(x).strip() for x in char_df["角色名称"]])

    # 校验页改成"数据完整性检查"：现在没有手填表原值可比，改为查名单/有效词条/权重是否齐
    check = build_check_sheet(char_df, chars, load_character_weights(), dropped, char_resolver, bad)
    readme = build_readme_sheet(data, len(df_old), len(new_rows), len(bad), len(set(lost)),
                                n_noeff=len(set(no_eff)), n_added=len(cfg_added))

    # ---- 角色名统一换成官方全名 ----
    # 内部所有匹配逻辑仍用"项目统一名"（handled by build_all），这里只在最后一步把显示名换掉，
    # 这样万一名字换错了也不影响计算。JSON 里没有的角色（如「欢愉主」）保持原名。
    # char_resolver 已在上面建好（校验页要用），这里直接复用，不重复构建。
    official_of = {}
    for c in chars:
        canon = resolve_char(c["name"], char_resolver) or c["name"]
        official_of[canon] = c["name"]

    def to_official(name):
        n = str(name).strip()
        return official_of.get(resolve_char(n, char_resolver) or n, n)

    # 配置表有、JSON 无（本账号没这个角色 / 没导出来）的，按官方名提示一遍。
    # 放在这里打印是因为要先把 official_of 建好——配置表里的错字别名也能翻译成官方名。
    if dropped:
        print("· 配置表有、本账号 JSON 里没有，本次跳过（不影响其它角色）：%s"
              % "、".join(to_official(n) for n in dropped))

    for _df in (df_old, pending, bad, check):
        if "角色名称" in _df.columns:
            _df["角色名称"] = list(map(to_official, _df["角色名称"]))

    # 原样复制另外几个页签（来自配置表，header=None 读入，写回时也不写表头，保证一字不差）
    # 这样生成出来的整份文件就能直接当 main.py 的输入表用
    # 「分解阈值」也在其中：10/40 这类口径放配置表，main.py 从生成表读，改口径不用碰代码
    raw_sheets = {}
    for s in ("遗器相关", "主词条分布", "分解阈值"):
        d, src = load_config_raw(s)
        if d is None:
            print("!! 配置表里没有「%s」页，生成表将缺少这一页（不影响 main.py 计算）" % s)
            continue
        raw_sheets[s] = d

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df_old.to_excel(w, sheet_name="角色遗器", index=False)
        pending.to_excel(w, sheet_name="有效词条待补", index=False)
        bad.to_excel(w, sheet_name="数据不完整", index=False)
        check.to_excel(w, sheet_name="校验", index=False)
        readme.to_excel(w, sheet_name="生成说明", index=False)
        for s, d in raw_sheets.items():
            d.to_excel(w, sheet_name=s, index=False, header=False)

    # ---- 控制台摘要 ----
    print("输出: %s" % out_path)
    print("角色遗器 %d 行（JSON 里未登记的 %d 个角色只进待补页）" % (len(df_old), len(new_rows)))
    nums = [r[c] for r in old_rows for c in SLOT_KEYS if isinstance(r.get(c), int)]
    print("六部位次数: %d 个格子，合计 %d 次，单格最大 %d"
          % (len(nums), sum(nums), max(nums) if nums else 0))
    print("校验页提示 %d 条；有效词条待补 %d 个角色；数据不完整 %d 条（%d 个角色）"
          % (len(check), len(pending), len(bad), len(set(lost))))
    if lost:
        print("!! 数据不完整、本次未参与计算：%s" % "、".join(sorted(set(lost))))
        print("   明细见生成表「数据不完整」页；补齐装备后重跑本脚本即可")
    if new_rows:
        print("!! JSON 里有、配置表还没登记（不会进「角色遗器」页）：%s"
              % "、".join(r["角色名称"] for r in new_rows))
        print("   去生成表「有效词条待补」页看推荐词条，复制到 config 配置表 + 角色权重.xlsx 即可纳入")
    if legacy:
        print("   （配置表有、JSON 无数据，本次跳过）:", legacy)


if __name__ == "__main__":
    main()
