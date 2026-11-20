# -*- coding: utf-8 -*-
"""
JSON 全量提取脚本（第一个脚本：只做"记录"，不做筛选）
====================================================
用途：
    把三月七小助手导出的 characters_xxx.json 里的**所有字段**摊平成 Excel 多页签，
    作为后续"第二个脚本（只提取关键信息）"的原始底稿。

输入：
    characters.json             （命令行第 1 个参数，默认取项目目录下的这个；
                                  旧名 characters_大号.json 仍兼容，会自动认）
输出：
    全量记录-豆包.xlsx           （命令行第 2 个参数，默认这个）

页签规划：
    00_说明        提取口径与本次数据体检结论
    01_角色总览    76 个角色 × 角色级全部字段（属性/光锥/星魂/命途/套装结构）
    02_遗器明细    456 件遗器 × 部位/套装/主词条/4 条副词条（核心页）
    03_套装组合    每个角色的套装构成（区分 4+2 / 2+2+2 / 4+1+1）
    04_技能        每个角色每项技能的等级
    05_词条字典    stat key -> 中文 的唯一映射表（12 条副词条 + 主词条专属）
    06_映射对照    JSON 角色名 -> Excel 角色名、套装全称 -> 简称
    07_校验        数据体检与"有效词条"逐件比对结果

设计原则：
    1. 只转换、不判定。除"有效词条比对"这一列外，不替你决定任何取舍，原样落表。
    2. 代码里写死的映射只有「stat key → 中文」「部位 → 项目用词」这两类固定不变的东西；
       名字对照（角色别名、套装别名）全部放在同目录的「别名配置.xlsx」，代码循环读，改配置不用改代码。
    3. 找不到对应项一律留空并进 07_校验，绝不猜。
"""

import io
import os
import sys
import json
import math
import re
import itertools
import collections

import pandas as pd

# ============================================================
# 一、路径与常量
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 本脚本在 code\ 下，配置文件在 config\ 下、数据与产物在项目根目录下，
# 所以这里往上退一层拿根目录，再由此拼各路径（不再依赖运行时的当前工作目录）。
ROOT_DIR = os.path.dirname(BASE_DIR)
CONFIG_DIR = os.path.join(ROOT_DIR, "config")


def _pick_json_path():
    """
    定位角色数据：主用 characters.json；万一还留着旧名 characters_大号.json，
    也兼容读一次并提醒改名，免得旧环境/旧副本一改名字就跑不起来。
    """
    new = os.path.join(ROOT_DIR, "characters.json")
    if os.path.exists(new):
        return new
    old = os.path.join(ROOT_DIR, "characters_大号.json")
    if os.path.exists(old):
        print("!! 建议把「characters_大号.json」改名为「characters.json」（现行名）")
        return old
    return new


JSON_PATH = _pick_json_path()
# 旧手填表（星铁装备刷取.xlsx）已于 2026-11-19 退役，移入 归档-豆包\。
# 这里不再把它当任何输入源；load_config_table / load_config_raw 的兜底参数也改为 None。
# 全量记录只是勘察/留档用的中间产物，不是每次都要看的结果，
# 所以默认写进「归档-豆包」子目录，免得跟 分析结果.xlsx 混在根目录里。
ARCHIVE_DIR = os.path.join(ROOT_DIR, "归档-豆包")
OUT_PATH = os.path.join(ARCHIVE_DIR, "全量记录-豆包.xlsx")

# 命令行可覆盖：python json_全量提取-豆包.py 输入.json 输出.xlsx
if len(sys.argv) > 1:
    JSON_PATH = sys.argv[1]
if len(sys.argv) > 2:
    OUT_PATH = sys.argv[2]

# 部位：JSON 用词 -> 项目用词（Excel「角色遗器」表头用的那套）
SLOT_CN = {
    "头部": "头",
    "手部": "手",
    "躯干": "衣服",
    "脚部": "鞋子",
    "位面球": "球",
    "连结绳": "绳子",
}
# 出金来源：外圈 4 部位（0.25）、内圈 2 部位（0.5）
SLOT_ZONE = {
    "头部": "外圈", "手部": "外圈", "躯干": "外圈", "脚部": "外圈",
    "位面球": "内圈", "连结绳": "内圈",
}

# 项目「角色遗器」表里 6 个数字列 -> JSON 部位名（第 17 列以后算出来的）
EXCEL_NUM_COL = {"头": "头部", "手": "手部", "衣服": "躯干",
                 "鞋子": "脚部", "球": "位面球", "绳子": "连结绳"}

# 「角色遗器」表的 17 个标准列（手填表缺失时用它们造一张空表，保证下游逻辑不用判空）
CHAR_COLS = ["角色名称", "遗器套装名称", "内圈套装名称", "有效词条数量", "有效词条",
             "衣服主词条", "鞋子主词条", "球主词条", "绳子主词条",
             "头", "手", "衣服", "鞋子", "球", "绳子", "词条数量", "权重"]

# 副词条权重（与 math_cal 的 init_equip_res_dict 保持一致，仅作展示列，不参与计算）
RES_WEIGHT = {
    "小生命": 100, "小攻击": 100, "小防御": 100,
    "大生命": 100, "大攻击": 100, "大防御": 100,
    "效果抵抗": 80, "效果命中": 80, "击破": 80,
    "暴击": 60, "暴伤": 60,
    "速度": 40,
}

# ------------------------------------------------------------------
# 副词条"强化次数"还原表
# ------------------------------------------------------------------
# 背景：游戏里一条副词条的最终数值 = 若干次强化的总和，每次强化从低/中/高三档里等概率取值。
#       所以"数值"其实隐含了"这条词条被投了几次"。把次数还原出来，才能算出每个部位的
#       "有效词条次数"（= 有效副词条上累计投了几次），这正是项目 Excel 里 头/手/衣服/... 那 6 个数。
#
# 档位来源（两处独立来源互相印证，数值一致）：
#   · NGA《崩铁|遗器体力价值论》2023-07 的"副词条强化档位"表
#   · gamevika.com 遗器页的 "value of one roll" 表（并注明游戏面板会向下截断显示）
#
# 面板显示规则（这是还原能否成功的关键）：
#   平铺属性（小生命/小攻击/小防御/速度）→ 面板按【整数向下截断】
#   百分比属性（大生命/大攻击/大防御/暴击/暴伤/命中/抵抗/击破）→ 面板按【1 位小数向下截断】
#   注意档位本身带 3 位小数（如大生命 3.456/3.888/4.32），用整数近似会算错，必须用精确值。
ROLL_TIER = {
    ("hp", 0):  (33.87, 38.10, 42.34),        # 小生命
    ("atk", 0): (16.94, 19.05, 21.17),        # 小攻击
    ("def", 0): (16.94, 19.05, 21.17),        # 小防御
    ("hp", 1):  (3.456, 3.888, 4.32),         # 大生命
    ("atk", 1): (3.456, 3.888, 4.32),         # 大攻击
    ("def", 1): (4.32, 4.86, 5.40),           # 大防御（比另外两个基础%高一档）
    ("spd", 0): (2.00, 2.304, 2.60),          # 速度
    ("crit_rate", 1): (2.592, 2.916, 3.24),   # 暴击
    ("crit_dmg", 1):  (5.184, 5.832, 6.48),   # 暴伤
    ("effect_hit_rate", 1): (3.456, 3.888, 4.32),   # 效果命中
    ("effect_res", 1): (3.456, 3.888, 4.32),        # 效果抵抗
    ("break_effect", 1): (5.184, 5.832, 6.48),      # 击破
}

# 每件遗器副词条总强化次数：初始 4 条 = 4 + 5 次强化 = 9；初始 3 条 = 3 + 1(补第 4 条) + 4 = 8
TOTAL_ROLLS = (8, 9)


def _panel_display(x, percent):
    """把游戏内部精确值转成面板显示值（向下截断）"""
    return math.floor(x * 10 + 1e-9) / 10.0 if percent else float(math.floor(x))


def _build_roll_table():
    """
    预计算：某个词条投 k 次(1~5)，在面板上可能显示成哪些值 -> 对应的 k 集合。
    例：大生命投 3 次最低档 = 10.368 -> 面板 10.3
    """
    tab = {}
    for key, tiers in ROLL_TIER.items():
        pct = key[1] == 1
        per_k = {}
        for k in range(1, 6):
            for combo in itertools.combinations_with_replacement(tiers, k):
                shown = _panel_display(sum(combo), pct)
                per_k.setdefault(shown, set()).add(k)
        tab[key] = per_k
    return tab


ROLL_TAB = _build_roll_table()


def solve_rolls(subs):
    """
    给一件遗器的 4 条副词条，反解每条投了几次。

    返回 (解列表, 备注)。解列表里每个元素是长度 4 的元组；正常情况下只有 1 个解。
    约束：4 条次数相加必须落在 TOTAL_ROLLS (8 或 9)，这是"每件 +15 遗器强化次数固定"的硬约束。
    """
    keys = [(s["stat"], 1 if s["percent"] else 0) for s in subs]
    cand = []
    for s, key in zip(subs, keys):
        ks = sorted(ROLL_TAB[key].get(float(s["value"]), []))
        if not ks:
            return [], "数值 %s 不在已知档位组合内" % s["value"]
        cand.append(ks)
    sols = [kk for kk in itertools.product(*cand) if sum(kk) in TOTAL_ROLLS]
    if not sols:
        return [], "次数组合凑不出 8/9 次"
    return sols, ""

# ============================================================
# 一·五、名字对照表：从「别名配置.xlsx」循环读，代码里不写死任何别名
# ============================================================
# 为什么单独放文件：
#   以前别名是代码里的两个字典字面量，加一行就得改代码，而且改了容易忘。
#   现在统一放在同目录的「别名配置.xlsx」，用 pandas 循环读进来。
#   加一个别名 = 在表里加一行，下次运行自动生效，代码一行都不用动。
#
# 配置结构（6 个页签）：
#   角色别名        官方名 | 项目名 | 备注      —— JSON 名 ↔ 项目表写法
#   套装别名        项目全称 | 游戏内正确名 | 备注 —— 打错的名字 ↔ 正确名字
#   套装补充简称    套装 | 补充简称 | 备注        —— 给某套遗器补一个额外叫法（可选）
#   遗器相关        遗器全称 | 刷取位置 | 简称1    —— 套装名的唯一来源（原在 星铁装备刷取.xlsx）
#   主词条分布      各部位主词条出货比例（分块排版，原在 星铁装备刷取.xlsx）
#   说明            纯文档，不参与读取
#
# 这些都属于「静态配置」：平时几乎不变，手工维护；
# 会随游戏进度变化的数据（角色遗器）由脚本从 JSON 自动生成，不放这里。
#
# 容错：文件不存在 / 页签缺失 / 整页空 → 一律当"没有这一项"，
#       靠归一化匹配 + 06_映射对照 的「未匹配」兜底，绝不因为缺配置就直接崩。
CONFIG_NAME = "项目配置-豆包.xlsx"          # 现行文件名
CONFIG_NAME_LEGACY = "别名配置.xlsx"        # 旧名（只有别名那 4 页），存在就兼容读取


def _pick_config_path():
    """
    定位配置文件：新名优先；只有旧名时用旧的并提示改名；都没有就按新名生成空壳。
    配置文件统一放在项目根目录下的 config 子目录里。
    """
    new = os.path.join(CONFIG_DIR, CONFIG_NAME)
    if os.path.exists(new):
        return new
    old = os.path.join(CONFIG_DIR, CONFIG_NAME_LEGACY)
    if os.path.exists(old):
        print("!! 建议把「%s」改名为「%s」（现在里面还少了「遗器相关 / 主词条分布」两页）"
              % (CONFIG_NAME_LEGACY, CONFIG_NAME))
        return old
    return new


ALIAS_XLSX = _pick_config_path()

# 配置文件缺失时，生成一份只有表头的空壳，让你知道该填什么
ALIAS_SHEET_HEADERS = {
    "角色别名": ["官方名", "项目名", "备注"],
    "套装别名": ["项目全称", "游戏内正确名", "备注"],
    "套装补充简称": ["套装", "补充简称", "备注"],
    "遗器相关": ["遗器全称", "刷取位置", "简称1"],
    "说明": ["项目", "说明"],
}   # 注：主词条分布是分块排版，空表头表达不了；它缺失不影响计算（main.py 里是常量）


def create_empty_alias_config(path):
    """配置文件丢了就补一个空壳（只有表头），保证脚本还能跑"""
    empty = {s: pd.DataFrame(columns=cols) for s, cols in ALIAS_SHEET_HEADERS.items()}
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for s, d in empty.items():
            d.to_excel(w, sheet_name=s, index=False)
    print("!! 没找到「%s」，已生成空表头模板；请把配置补进去，否则只能靠归一化匹配"
          % path)


def load_config_table(sheet, cols, legacy_xlsx=None):
    """
    读一张静态配置表（如「遗器相关」），兜底顺序：
        ① 配置表里有   → 用配置表（正常路径）
        ② 配置表没有、但显式传了旧表路径（迁移期兼容用，现在默认不传）→ 用旧的
        ③ 都没有       → 返回带正确列名的空表，不报错
    返回 (DataFrame, 来源路径或 None)
    """
    for path in (ALIAS_XLSX, legacy_xlsx):
        if path and os.path.exists(path):
            try:
                df = pd.read_excel(path, sheet_name=sheet)
                if not df.empty:
                    return df, path
            except Exception:
                continue
    return pd.DataFrame(columns=cols), None


def load_config_raw(sheet, legacy_xlsx=None):
    """
    同 load_config_table，但**不把首行当表头**（header=None），
    用于「主词条分布」这种分块排版、没有规范表头的表。
    读不到返回 (None, None)。
    """
    for path in (ALIAS_XLSX, legacy_xlsx):
        if path and os.path.exists(path):
            try:
                return pd.read_excel(path, sheet_name=sheet, header=None), path
            except Exception:
                continue
    return None, None


def _read_sheet(path, sheet, cols):
    """
    读一个页签，返回 [(第一列, 第二列, 备注), ...]。
    页签不存在 / 全空 / 值缺失 → 返回空列表，不报错。
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return []
    if df.empty:
        return []
    for c in cols[:2]:
        if c not in df.columns:
            return []
    rows = []
    for _, r in df.iterrows():
        a, b = r[cols[0]], r[cols[1]]
        note = r[cols[2]] if len(cols) > 2 and cols[2] in df.columns else ""
        # 两列都空的行（用来做视觉分隔）直接跳过
        if pd.isna(a) and pd.isna(b):
            continue
        # 只有一列有值 → 视为填错，跳过并提示
        if pd.isna(a) or pd.isna(b):
            print("!! 别名配置「%s」页有一行只填了一半，已跳过：%s" % (sheet, (a, b)))
            continue
        rows.append((str(a).strip(), str(b).strip(),
                     "" if pd.isna(note) else str(note).strip()))
    return rows


def load_alias_config(path=None):
    """
    读入全部别名配置，返回 dict：
        char   [(官方名, 项目名, 备注), ...]
        setp   [(项目全称, 游戏内正确名, 备注), ...]
        extra  [(套装, 补充简称, 备注), ...]
    """
    path = path or ALIAS_XLSX
    if not os.path.exists(path):
        create_empty_alias_config(path)
    return {
        "char": _read_sheet(path, "角色别名", ALIAS_SHEET_HEADERS["角色别名"]),
        "setp": _read_sheet(path, "套装别名", ALIAS_SHEET_HEADERS["套装别名"]),
        "extra": _read_sheet(path, "套装补充简称", ALIAS_SHEET_HEADERS["套装补充简称"]),
    }


ALIAS_CFG = load_alias_config()

# 代码里只保留"由配置推导出来的字典"，不再有任何写死的条目
# 角色：官方名 -> 项目名
NAME_ALIAS = {}
for _official, _short, _note in ALIAS_CFG["char"]:
    NAME_ALIAS[_official] = _short
    NAME_ALIAS.setdefault(_short, _short)      # 简称自己也登记一次，双向都能查

# 套装：游戏内正确名 -> 项目全称
SET_ALIAS = {}
for _project, _official, _note in ALIAS_CFG["setp"]:
    SET_ALIAS[_official] = _project
    SET_ALIAS.setdefault(_project, _project)

# 套装补充简称：(套装任意写法, 补充简称)
EXTRA_ABBR = [(_set, _abbr) for _set, _abbr, _note in ALIAS_CFG["extra"]]


# ============================================================
# 二、词条名翻译：把 (stat, percent) 变成项目里的中文词条名
# ============================================================
def stat_cn(stat, percent):
    """
    项目里的词条名体系：大小三围靠 percent 区分，其余按 key 一对一。
    返回 None 表示"这个组合在项目里没有对应中文"（会进 07_校验）。
    """
    flat = {"hp": "小生命", "atk": "小攻击", "def": "小防御", "spd": "速度"}
    pct = {"hp": "大生命", "atk": "大攻击", "def": "大防御"}
    fixed = {
        "crit_rate": "暴击",
        "crit_dmg": "暴伤",
        "effect_hit_rate": "效果命中",
        "effect_res": "效果抵抗",
        "break_effect": "击破",
        # 只可能出现在主词条
        "energy_regen": "充能",
        "outgoing_healing": "治疗量",
        "physical_dmg": "物伤",
        "fire_dmg": "火伤",
        "ice_dmg": "冰伤",
        "lightning_dmg": "雷伤",
        "wind_dmg": "风伤",
        "quantum_dmg": "量子伤",
        "imaginary_dmg": "虚数伤害",
        "elation": "欢愉",
    }
    if stat in ("hp", "atk", "def"):
        return pct[stat] if percent else flat[stat]
    if stat == "spd":
        return "速度"
    return fixed.get(stat)


def name_norm(s):
    """名字归一化，用于"官方全名 vs 简称"的自动匹配"""
    s = str(s)
    for t in ("·", "&账账", "★", " ", "　"):
        s = s.replace(t, "")
    s = re.sub(r"LV\.?", "", s, flags=re.I)   # 银狼LV.999 -> 银狼999
    s = s.replace("医生", "").replace("神", "")  # 真理医生 / 星神★阿哈
    return s.casefold()


def build_char_resolver(json_names, excel_names):
    """
    角色名「解析器」：把任何一个见过的写法，都解析成项目统一用的那个名字。

    项目统一名（canonical）= Excel「角色遗器」表里的写法，因为 main.py 就是按这个名字跑的。
    但同一角色可能有好几种叫法：
        · 游戏内官方全名  银狼LV.999 / 丹恒·饮月 / 真理医生
        · 项目表简称      银狼999 / 龙丹 / 真理
        · 归一化后能对上的  缇宝 vs 提宝、藿藿 vs 霍霍（只差一个字/写法）
        · 开拓者          陆遥（当前命途）↔ 记忆主

    这个解析器把上面所有写法都登记成 key，值统一指向 canonical，
    于是「用哪个名字都能找到同一个角色」，不再依赖单向的 JSON→Excel 转换。

    返回 (alias, norm_map)：
        alias    {精确写法: canonical}   —— 命中就不再做归一化
        norm_map {归一化写法: canonical} —— 兜底，处理标点/空格/LV.999 这类差异
    """
    alias, norm_map = {}, {}

    # ① Excel 名单本身就是 canonical，先把它们全部登记
    for e in excel_names:
        e = str(e).strip()
        if not e:
            continue
        alias.setdefault(e, e)
        norm_map.setdefault(name_norm(e), e)

    # ② 人工别名：JSON 官方名 → Excel 简称，双向都登记
    #    双向的好处：以后数据源换成"官方全名"或"项目简称"都能认
    for official, short in NAME_ALIAS.items():
        target = short if short in alias else alias.get(short, short)
        alias[official] = target
        norm_map.setdefault(name_norm(official), target)
        alias.setdefault(short, target)

    # ③ JSON 里剩下的名字：先看归一化能不能对上，对不上就自己当 canonical
    for j in json_names:
        j = str(j).strip()
        if not j:
            continue
        target = alias.get(j) or norm_map.get(name_norm(j)) or j
        alias[j] = target
        norm_map.setdefault(name_norm(j), target)
    return alias, norm_map


def resolve_char(name, resolver):
    """
    任意写法 → 项目统一名。
    先按精确写法查，再做归一化兜底；都查不到就返回 None（真不认识）。
    """
    alias, norm_map = resolver
    name = str(name).strip()
    if name in alias:
        return alias[name]
    return norm_map.get(name_norm(name))


def build_name_map(json_names, excel_names):
    """
    角色名映射：JSON 名 → Excel 名，对不上留空（保持老接口，内部改用统一解析器）。
    返回 {JSON名: Excel名 或 None}
    """
    resolver = build_char_resolver(json_names, excel_names)
    excel_set = {str(e).strip() for e in excel_names}
    mapping, unmatched = {}, []
    for j in json_names:
        target = resolve_char(j, resolver)
        # 解析到"自己"说明 Excel 里根本没有这个角色
        if target in excel_set:
            mapping[j] = target
        else:
            mapping[j] = None
            unmatched.append(j)
    return mapping, unmatched


# ============================================================
# 三、读源数据
# ============================================================
def build_set_alias_table(set_df):
    """
    套装名「别名表」：把同一套遗器的所有叫法，全部映射到项目表的写法。

    同一套遗器在项目里可能有三种写法同时存在：
        · 项目表全称    「戏梦点星的怜人」(※ 项目表把"伶"打成了"怜")
        · 项目表简称    「诗人」「拾骨地、哀地」(简称1 里可能有多个，用「、」分隔)
        · 游戏内正确名  「戏梦点星的伶人」（JSON 导出的是这个）

    以前的做法是：只把「正确名 → 项目写法」这一个方向写死（SET_ALIAS），
    于是反过来"项目写法 → 正确名"就认不出来。这里改成**双向登记**：

        alias2full[任何写法] = 项目表全称（canonical）
        full2abbr[项目表全称] = 主用简称

    这样无论表里写的是正确名还是打错的名字、是全称还是简称，都能解析到同一套遗器。

    其中"正确名 ↔ 项目写法"的对照来自「别名配置.xlsx」的「套装别名」页（SET_ALIAS），
    额外叫法来自「套装补充简称」页（EXTRA_ABBR），本函数只负责循环挂上去。

    返回 (alias2full, full2abbr)
    """
    alias2full, full2abbr = {}, {}
    for _, r in set_df.iterrows():
        full = str(r["遗器全称"]).strip()
        if not full:
            continue
        abbr_raw = "" if pd.isna(r["简称1"]) else str(r["简称1"]).strip()
        tokens = [t.strip() for t in abbr_raw.split("、") if t.strip()]

        full2abbr[full] = tokens[0] if tokens else full
        alias2full[full] = full                    # 全称自己
        for t in tokens:                           # 表里登记过的所有简称
            alias2full.setdefault(t, full)

    # 项目表写法 ↔ 游戏内正确写法，两个方向都登记
    # 数据来自「别名配置.xlsx」的「套装别名」页，这里只负责循环挂上去
    for official, project in SET_ALIAS.items():
        if project in full2abbr:                   # 项目表用的是打错的那个 → 正确名也指向它
            alias2full.setdefault(official, project)
        elif official in full2abbr:                # 反过来项目表已经用对了 → 打错的名字也认
            alias2full.setdefault(project, official)
        else:                                      # 两边都没有，就别乱认了
            continue

    # 「套装补充简称」页：给某套遗器补一个「遗器相关」表里没登记的叫法
    # 这一列不参与 full2abbr（主用简称仍取表里的），只让识别时能认
    for set_name, extra_abbr in EXTRA_ABBR:
        full = resolve_set(set_name, alias2full)
        if full:
            alias2full.setdefault(extra_abbr, full)
        else:
            print("!! 别名配置「套装补充简称」页里的「%s」认不出是哪套遗器，已跳过" % set_name)
    return alias2full, full2abbr


def resolve_set(name, alias2full):
    """
    任意写法 → 项目表全称。认不出返回 None。
    归一化兜底：去掉「的」「之」这类连接字后再比一次，防止只因一个虚词对不上。
    """
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return None
    name = str(name).strip()
    if name in alias2full:
        return alias2full[name]

    def loose(s):
        return re.sub(r"[的之]", "", str(s)).replace("　", "").replace(" ", "")

    cand = loose(name)
    for k, v in alias2full.items():
        if loose(k) == cand:
            return v
    return None


def abbr_of(name, set_resolver):
    """任意写法 → 项目主用简称；认不出来就给空串"""
    alias2full, full2abbr = set_resolver
    full = resolve_set(name, alias2full)
    return full2abbr.get(full, "") if full else ""


def normalize_trailblazer(data):
    """
    开拓者（主角）的名字在 JSON 里是**玩家自定义昵称**（如「陆遥」「羽球忍者」），
    每个账号都不一样，拿名字做匹配必然出错。好在 JSON 给了 `is_trailblazer=True`，
    所以这里统一改写名字，让下游完全不用关心玩家叫什么：

        is_trailblazer=True 时  name = "开拓者·" + path

    path 是当前命途（开拓者可以切命途，但一个账号一次只导当前那一份），
    所以「开拓者·记忆」表示记忆命途的开拓者，「开拓者·欢愉」表示欢愉命途的。
    改完之后由配置表「角色别名」页把它映射成项目惯用名（记忆主 / 欢愉主），
    没登记的形态会自然落到「有效词条待补」页，不会被静默丢掉。

    path 缺失（理论上不会）时退化为「开拓者」，不会崩。

    原名留在 `trailblazer_nickname` 里，方便追溯"这个号的主角叫什么"。
    返回改过的 data（原地修改，也顺便返回自己，方便链式调用）。
    """
    for c in data.get("characters", []) or []:
        if not c.get("is_trailblazer"):
            continue
        old = str(c.get("name", "")).strip()
        path = str(c.get("path", "") or "").strip()
        c["trailblazer_nickname"] = old          # 留档：玩家自己起的名字
        c["name"] = "开拓者·" + path if path else "开拓者"
    return data


def load_json(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return normalize_trailblazer(json.load(f))


def load_excel():
    """
    读配置：
        角色遗器 —— 【已退役】原来是手填的旧数据表，只用于 06_映射对照 / 07_校验 的比对。
                    2026-11-19 起不再读它（手填表已移入 归档-豆包），那两页改为显示"本次未提供手填值"。
                    角色名单现在统一来自 config 「角色有效词条」页，由 json_关键信息 脚本消费。
        遗器相关 —— 套装名表，已搬到「项目配置」里（load_config_table 三档兜底）

    返回 (char_df, set_map, set_df)，其中 set_map 是**扁平化的别名表**：
        set_map[任何写法] = 主用简称
    这样旧代码里的 set_map.get(x, "") 一个写法都不用改，就自动支持多写法识别了。
    """
    # 手填表退役：给一张只有列名的空表，让下游逻辑照常走（比对结果全为"无旧值"）
    char_df = pd.DataFrame(columns=CHAR_COLS)

    set_df, set_src = load_config_table("遗器相关",
                                        ["遗器全称", "刷取位置", "简称1"])
    if set_src and os.path.basename(set_src) != CONFIG_NAME:
        print("!! 「遗器相关」目前还是从「%s」读的，建议搬进「%s」"
              % (os.path.basename(set_src), CONFIG_NAME))

    alias2full, full2abbr = build_set_alias_table(set_df)
    # 扁平化成 {写法: 简称}，全称与各种别名都进 key
    set_map = {k: full2abbr.get(v, "") for k, v in alias2full.items()}
    return char_df, set_map, set_df


# ============================================================
# 四、各页签的构造
# ============================================================
def build_overview(chars, name_map):
    """01_角色总览：一个角色一行，把 stats / 光锥 / 套装结构全部摊开"""
    rows = []
    for c in chars:
        st = c["stats"]
        base = st.get("base", {})
        adv = st.get("advanced", {})
        dmg = st.get("damage", {})
        lc = c.get("light_cone") or {}
        lc_st = (lc.get("stats") or {})

        # 套装结构：把 sets 列表压成"4+2"这种字符串，并按件数排序
        sets = c["relics"].get("sets", [])
        shape = "+".join(str(s["pieces"]) for s in sorted(sets, key=lambda x: -x["pieces"]))
        set_desc = " / ".join("%s(%d)" % (s["name"], s["pieces"])
                              for s in sorted(sets, key=lambda x: -x["pieces"]))

        row = {
            "角色名(JSON)": c["name"],
            "对应Excel名": name_map.get(c["name"]) or "",
            "在Excel名单内": "是" if name_map.get(c["name"]) else "否",
            "是否开拓者": "是" if c.get("is_trailblazer") else "否",
            "可否切命途": "是" if c.get("path_switchable") else "否",
            "命途": c.get("path", ""),
            "属性": c.get("element", ""),
            "等级": c.get("level"),
            "等级上限": c.get("max_level"),
            "星魂": c.get("eidolon"),
            # ---- 光锥 ----
            "光锥名称": lc.get("name", ""),
            "光锥命途": lc.get("path", ""),
            "光锥等级": lc.get("level"),
            "光锥满级": lc.get("max_level"),
            "光锥叠影": lc.get("superimposition"),
            "光锥生命": lc_st.get("hp"),
            "光锥攻击": lc_st.get("atk"),
            "光锥防御": lc_st.get("def"),
            # ---- 基础三围 + 速度（total/base/extra 三元组）----
            "生命合计": (base.get("hp") or {}).get("total"),
            "生命基础": (base.get("hp") or {}).get("base"),
            "生命加成": (base.get("hp") or {}).get("extra"),
            "攻击合计": (base.get("atk") or {}).get("total"),
            "攻击基础": (base.get("atk") or {}).get("base"),
            "攻击加成": (base.get("atk") or {}).get("extra"),
            "防御合计": (base.get("def") or {}).get("total"),
            "防御基础": (base.get("def") or {}).get("base"),
            "防御加成": (base.get("def") or {}).get("extra"),
            "速度合计": (base.get("spd") or {}).get("total"),
            "速度基础": (base.get("spd") or {}).get("base"),
            "速度加成": (base.get("spd") or {}).get("extra"),
            # ---- 高级属性 ----
            "暴击率%": adv.get("crit_rate"),
            "暴击伤害%": adv.get("crit_dmg"),
            "击破特攻%": adv.get("break_effect"),
            "治疗加成%": adv.get("outgoing_healing"),
            "能量恢复%": adv.get("energy_regen"),
            "效果命中%": adv.get("effect_hit_rate"),
            "效果抵抗%": adv.get("effect_res"),
            "欢愉值%": adv.get("elation"),
            # ---- 属性伤害 / 抗性 ----
            "物伤%": dmg.get("physical_dmg"),
            "火伤%": dmg.get("fire_dmg"),
            "冰伤%": dmg.get("ice_dmg"),
            "雷伤%": dmg.get("lightning_dmg"),
            "风伤%": dmg.get("wind_dmg"),
            "量子伤%": dmg.get("quantum_dmg"),
            "虚数伤%": dmg.get("imaginary_dmg"),
            "物抗%": dmg.get("physical_res"),
            "火抗%": dmg.get("fire_res"),
            "冰抗%": dmg.get("ice_res"),
            "雷抗%": dmg.get("lightning_res"),
            "风抗%": dmg.get("wind_res"),
            "量子抗%": dmg.get("quantum_res"),
            "虚数抗%": dmg.get("imaginary_res"),
            "其他属性(原始)": json.dumps(st.get("other", {}), ensure_ascii=False)
                              if st.get("other") else "",
            # ---- 套装结构 ----
            "套装件数结构": shape,
            "套装构成": set_desc,
            "套装组数": len(sets),
            "导出告警": "; ".join(c.get("warnings", [])) if c.get("warnings") else "",
        }
        rows.append(row)
    return pd.DataFrame(rows)


def build_relic_rows(chars, name_map, set_map, char_df):
    """
    02_遗器明细：一件遗器一行（76 × 6 = 456 行），这是核心页。
    除原始字段外，额外做两件事：
      ① 把每件的主词条与 Excel 里该角色的主词条列对照（校验手填有没有错）
      ② 按该角色的「有效词条」判定这件的副词条有几条有效，并与 Excel 的 6 个数字列对照
    """
    # Excel 里每个角色的信息，供比对用
    excel_info = {}
    for _, r in char_df.iterrows():
        nm = str(r["角色名称"]).strip()
        eff_tokens = set()
        for t in str(r["有效词条"] or "").split("、"):
            if t.strip():
                eff_tokens.add(t.strip())
        # 主词条：Excel 只有 4 个部位有主词条列（头固定小生命、手固定小攻击，不必比）
        MAIN_COL = {
            "躯干": "衣服主词条",   # 衣服
            "脚部": "鞋子主词条",   # 鞋子
            "位面球": "球主词条",   # 球
            "连结绳": "绳子主词条", # 绳子
        }
        mains = {slot: str(r[col]).strip() if pd.notna(r[col]) else ""
                 for slot, col in MAIN_COL.items()}
        nums = {json_slot: (int(r[excel_col]) if pd.notna(r[excel_col]) else None)
                for excel_col, json_slot in EXCEL_NUM_COL.items()}
        excel_info[nm] = {"eff": eff_tokens, "mains": mains, "nums": nums}

    rows = []
    stat_key_seen = set()
    for c in chars:
        jname = c["name"]
        ename = name_map.get(jname)
        info = excel_info.get(ename) if ename else None

        for pc in c["relics"]["pieces"]:
            slot = pc["slot"]
            main = pc["main_stat"]
            main_cn = stat_cn(main["stat"], main["percent"])
            stat_key_seen.add((main["stat"], main["percent"], "主"))

            subs = []
            sub_cn_list = []
            for s in pc["sub_stats"]:
                cn = stat_cn(s["stat"], s["percent"])
                stat_key_seen.add((s["stat"], s["percent"], "副"))
                subs.append({
                    "key": s["stat"], "pct": s["percent"], "cn": cn, "val": s["value"],
                })
                sub_cn_list.append(cn or ("?%s" % s["stat"]))

            # ---- 反解每条副词条投了几次（精确，不是估算）----
            sols, solve_note = solve_rolls(pc["sub_stats"])
            if sols:
                # 极小概率出现多解：取"有效次数最大"的那个，与手数习惯（能确认的才算命中）一致
                if info:
                    def eff_of(kk):
                        return sum(k for sub, k in zip(pc["sub_stats"], kk)
                                   if stat_cn(sub["stat"], sub["percent"]) in info["eff"])
                    rolls = max(sols, key=eff_of)
                else:
                    rolls = sols[0]
                total_rolls = sum(rolls)
            else:
                rolls = [None] * len(subs)
                total_rolls = None

            # 有效统计：种类数（最多 4）与次数（有效副词条上累计投了几次）
            eff_kinds = 0
            eff_rolls = 0
            if info:
                for sub, k in zip(pc["sub_stats"], rolls):
                    if stat_cn(sub["stat"], sub["percent"]) in info["eff"]:
                        eff_kinds += 1
                        eff_rolls += (k or 0)

            set_full = pc["set"]
            set_abbr = set_map.get(set_full, "")

            # 与 Excel 手填的两个比对：主词条（只比有主词条列的 4 个部位）、有效次数
            excel_main = info["mains"].get(slot, "") if info else ""
            main_match = ""
            if info and slot in info["mains"]:
                main_match = "一致" if str(excel_main) == str(main_cn) else "不一致"
            elif info:
                # 头=小生命、手=小攻击 是固定值，Excel 本来就没有这两列，不算差异
                main_match = "不适用(固定)"
            excel_num = info["nums"].get(slot) if info else None
            diff = ""
            if info and excel_num is not None:
                diff = eff_rolls - excel_num

            row = {
                "角色名(JSON)": jname,
                "对应Excel名": ename or "",
                "部位(JSON)": slot,
                "部位(项目用词)": SLOT_CN.get(slot, ""),
                "内外圈": SLOT_ZONE.get(slot, ""),
                "遗器等级": pc["level"],
                "套装全称": set_full,
                "套装简称": set_abbr,
                "主词条key": main["stat"],
                "主词条是否百分比": "是" if main["percent"] else "否",
                "主词条(中文)": main_cn or "",
                "主词条数值": main["value"],
                "主词条同Excel": main_match,
                "Excel记录主词条": excel_main,
            }
            for i in range(4):
                s = subs[i] if i < len(subs) else None
                row["副词条%d(中文)" % (i + 1)] = (s["cn"] or "") if s else ""
                row["副词条%d数值" % (i + 1)] = s["val"] if s else ""
                row["副词条%d次数" % (i + 1)] = rolls[i] if i < len(rolls) else ""
                row["副词条%d权重" % (i + 1)] = RES_WEIGHT.get(s["cn"]) if s else ""
            row.update({
                "副词条拼接(顿号)": "、".join(sub_cn_list),
                "本条强化次数合计": total_rolls if total_rolls is not None else "",
                "胚子类型": ("初始4条" if total_rolls == 9 else
                             ("初始3条" if total_rolls == 8 else "")),
                "次数还原备注": solve_note,
                "有效词条种类数": eff_kinds if info else "",
                "有效次数": eff_rolls if info else "",
                "Excel手填次数": excel_num if excel_num is not None else "",
                "偏差(算-手填)": diff,
            })
            rows.append(row)

    return pd.DataFrame(rows), stat_key_seen, excel_info


def build_set_combo(chars):
    """
    03_套装组合：单独一页，方便一眼看出哪些角色不是标准 4+2。
    外圈/内圈直接按"这件穿在哪个部位"反推，不靠套装名清单猜，这样 1+1 散件也能显示出来。
    """
    rows = []
    for c in chars:
        sets = c["relics"]["sets"]
        packed = sorted(sets, key=lambda x: -x["pieces"])
        # 外圈 = 头/手/躯干/脚部 用到的套装（按出现顺序去重）
        outer, inner = [], []
        for pc in c["relics"]["pieces"]:
            bucket = outer if SLOT_ZONE.get(pc["slot"]) == "外圈" else inner
            if pc["set"] not in bucket:
                bucket.append(pc["set"])
        rows.append({
            "角色名(JSON)": c["name"],
            "件数结构": "+".join(str(s["pieces"]) for s in packed),
            "组数": len(sets),
            "外圈": " / ".join(outer),
            "内圈": " / ".join(inner),
            "第1套": packed[0]["name"] if len(packed) > 0 else "",
            "第1套件数": packed[0]["pieces"] if len(packed) > 0 else "",
            "第2套": packed[1]["name"] if len(packed) > 1 else "",
            "第2套件数": packed[1]["pieces"] if len(packed) > 1 else "",
            "第3套": packed[2]["name"] if len(packed) > 2 else "",
            "第3套件数": packed[2]["pieces"] if len(packed) > 2 else "",
            "是否标准4+2": "是" if (len(sets) == 2 and packed[0]["pieces"] == 4
                                  and packed[1]["pieces"] == 2 and len(outer) == 1) else "否",
        })
    return pd.DataFrame(rows)


def build_skills(chars):
    """04_技能：一个角色的每项技能一行"""
    rows = []
    for c in chars:
        for i, s in enumerate(c["skills"]):
            rows.append({
                "角色名(JSON)": c["name"],
                "序号": i + 1,
                "技能类型": s.get("type", ""),
                "等级": s.get("level"),
                "满级": s.get("max_level"),
                "是否满级": "是" if s.get("level") == s.get("max_level") else "否",
            })
    return pd.DataFrame(rows)


def build_stat_dict(stat_key_seen):
    """
    05_词条字典：本次数据里实际出现过的 (key, percent) 组合，
    附上项目词条名、副词条权重、单次强化三档数值（次数还原靠的就是这张表）。
    """
    rows = []
    for key, pct, usage in sorted(stat_key_seen, key=lambda x: (x[2], x[0], x[1])):
        cn = stat_cn(key, pct)
        tiers = ROLL_TIER.get((key, 1 if pct else 0))
        rows.append({
            "stat key": key,
            "percent": "是" if pct else "否",
            "项目词条名": cn or "",
            "出现位置": ("主词条" if usage == "主" else "副词条"),
            "项目副词条权重": RES_WEIGHT.get(cn) if usage == "副" else "",
            "单次低档": tiers[0] if tiers else "",
            "单次中档": tiers[1] if tiers else "",
            "单次高档": tiers[2] if tiers else "",
            "面板显示规则": ("整数截断" if (tiers and not pct and key != "spd") else
                              ("整数截断" if key == "spd" else ("1位小数截断" if tiers else ""))),
            "缺失映射": "" if cn else "是",
        })
    return pd.DataFrame(rows)


def build_map_sheet(chars, name_map, set_map, unmatched_names, alias2full):
    """06_映射对照：角色名 + 套装名两段，直接贴在一页里"""
    rows = []
    for c in chars:
        j = c["name"]
        e = name_map.get(j)
        rows.append({
            "类别": "角色名",
            "原始(JSON)": j,
            "项目用": e or "",
            "处理方式": ("人工别名" if j in NAME_ALIAS else
                         ("归一化匹配" if e else "未匹配")),
            "说明": "" if e else "Excel 里没有这个角色，需要新增或忽略",
        })
    used_sets = sorted({s["name"] for c in chars for s in c["relics"]["sets"]})
    for s in used_sets:
        # 现在 set_map 里既有项目写法、也有游戏内正确写法，直接查就行
        full = resolve_set(s, alias2full)
        abbr = set_map.get(s, "")
        note = ""
        if s in SET_ALIAS:
            note = "JSON 写法与项目表不同，已双向对齐为「%s」" % full
        elif full and full != s:
            note = "按别名/简称匹配到「%s」" % full
        if not abbr:
            note = (note + "；" if note else "") + "项目「遗器相关」表里没有简称，需要补充"
        rows.append({
            "类别": "套装名",
            "原始(JSON)": s,
            "项目用": abbr,
            "处理方式": ("别名修正" if s in SET_ALIAS else
                         ("直接匹配" if (abbr and full == s) else
                          ("别名/简称匹配" if abbr else "未匹配"))),
            "说明": note,
        })
    return pd.DataFrame(rows)


def build_check_sheet(chars, relic_df, stat_key_seen, unmatched_names, set_map, meta):
    """07_校验：本次数据体检 + 有效词条比对汇总，全部是"事实"，不含推断"""
    rows = []
    total_pieces = sum(len(c["relics"]["pieces"]) for c in chars)

    def add(项目, 结果, 说明):
        rows.append({"检查项": 项目, "结果": 结果, "说明": 说明})

    add("角色数", len(chars), "JSON 内的角色条目数")
    add("遗器总件数", total_pieces, "应等于角色数 × 6")
    add("每角色件数", "%d ~ %d" % (min(len(c["relics"]["pieces"]) for c in chars),
                                max(len(c["relics"]["pieces"]) for c in chars)),
        "是否每个角色都是 6 件")
    lv = collections.Counter(pc["level"] for c in chars for pc in c["relics"]["pieces"])
    add("遗器等级分布", json.dumps(dict(sorted(lv.items())), ensure_ascii=False),
        "全部 +15 说明没有半成品遗器")
    ns = collections.Counter(len(pc["sub_stats"]) for c in chars for pc in c["relics"]["pieces"])
    add("副词条条数分布", json.dumps(dict(ns), ensure_ascii=False),
        "+15 之后必定 4 条，少于 4 说明导出有问题")
    add("导出告警(warnings)", sum(1 for c in chars if c.get("warnings")),
        "有告警的角色数")
    add("failed 列表", len(meta.get("failed", [])), "导出失败的角色数，0 表示无缺失")

    # 重复 / 冲突（按 (stat, percent) 判身份，不能按裸 stat）
    dup = conflict = 0
    for c in chars:
        for pc in c["relics"]["pieces"]:
            ids = [(s["stat"], s["percent"]) for s in pc["sub_stats"]]
            if len(set(ids)) != len(ids):
                dup += 1
            if (pc["main_stat"]["stat"], pc["main_stat"]["percent"]) in ids:
                conflict += 1
    add("同一件重复副词条", dup, "按 (stat, percent) 判身份，大/小攻击算两条")
    add("主副词条真冲突", conflict, "同件里主词条与副词条完全相同")

    # 词条映射缺失
    missing = [k for k in stat_key_seen if not stat_cn(k[0], k[1])]
    add("未能翻译的词条组合", len(missing), str(missing))

    # 套装 / 角色名未匹配
    used_sets = {s["name"] for c in chars for s in c["relics"]["sets"]}
    no_abbr = [s for s in used_sets if not set_map.get(s, "")]
    add("未收录简称的套装", len(no_abbr), str(sorted(no_abbr)))
    # 角色名单现在以 config「角色有效词条」页为准（手填表已退役，不再有第二份名单可比）
    add("未在配置表登记的角色", len(unmatched_names) if unmatched_names else 0,
        "（JSON 里有、config「角色有效词条」页没登记，不会进「角色遗器」页，见生成表「有效词条待补」）"
        if unmatched_names else "全部已登记")

    # 有效词条逐件比对：只有"这一件在旧手填表里确实填了次数"才比。
    # 手填表 2026-11-19 起退役，这一列正常情况下是空的——空值不能当成"不一致"
    # （NaN != 0 在 pandas 里是 True，会把每一件都算成差异，曾经骗过一次）。
    cmp_all = relic_df[pd.notna(relic_df["对应Excel名"]) & (relic_df["对应Excel名"].astype(str) != "")]
    filled = cmp_all[cmp_all["Excel手填次数"].apply(
        lambda v: str(v).strip() not in ("", "nan", "None"))]
    if len(filled):
        n_all = len(filled)
        n_diff = int((filled["偏差(算-手填)"] != 0).sum())
        add("可比对遗器件数", n_all, "手填表里有次数的件数（含主词条比对与有效次数比对）")
        add("有效次数与手填一致的件数", n_all - n_diff,
            "%.1f%% 一致。口径=「有效强化次数(0~8/9)」，由档位表 + 面板截断规则精确还原"
            % (100.0 * (n_all - n_diff) / n_all))
        add("有效次数不一致的件数", n_diff,
            "明细见 02_遗器明细「偏差(算-手填)」列；差异一般来自手抄笔误或换过装备")
        n_main = int((filled["主词条同Excel"] == "不一致").sum())
        n_main_all = int(filled["主词条同Excel"].isin(["一致", "不一致"]).sum())
        add("主词条与手填不一致的件数", n_main,
            "只比有主词条列的 4 个部位（共 %d 件）：其中 %d 件不一致，明细见 02_遗器明细"
            % (n_main_all, n_main))
    else:
        add("有效次数比对", "跳过", "旧手填表已退役，本次没有可比的次数（只做还原，不做比对）")
    # 次数还原的可靠性
    bad_solve = int((relic_df["次数还原备注"] != "").sum())
    add("次数还原失败的件数", bad_solve, "0 表示 456 件全部精确还原成功")
    add("胚子类型分布",
        json.dumps(dict(collections.Counter(relic_df["胚子类型"].tolist())), ensure_ascii=False),
        "初始4条=9次、初始3条=8次，可当作「这件是不是 3 词条起手」的旁证")
    return pd.DataFrame(rows)



# ============================================================
# 五、主流程
# ============================================================
def main():
    data = load_json(JSON_PATH)
    chars = data["characters"]
    # 一次读完：char_df（手填表，可能为空）、set_map（别名扁平表）、set_df（套装名表）
    char_df, set_map, set_df = load_excel()
    # 套装名别名表（全称 + 各种简称 + 游戏内正确写法，全部指向项目表全称）
    alias2full, full2abbr = build_set_alias_table(set_df)

    # 手填表 2026-11-19 起退役，char_df 是空表；"对应Excel名 / 在Excel名单内"这两列
    # 改为跟配置表「角色有效词条」页比（那才是现在的角色名单来源），
    # 否则整列都是「否」，会让人误以为 JSON 里所有角色都没登记。
    cfg_names = []
    try:
        cfg, cfg_src = load_config_table("角色有效词条", ["角色名称", "有效词条"], None)
        if cfg is not None and not cfg.empty:
            cfg_names = [str(x).strip() for x in cfg["角色名称"].dropna()]
    except Exception as e:                       # 配置表缺失/改坏也不该让全量记录跑不出来
        print("!! 读配置表「角色有效词条」页失败（%s），角色名单比对列将全部为空" % e)

    excel_names = cfg_names or [str(x).strip() for x in char_df["角色名称"].dropna()]
    name_map, unmatched_names = build_name_map([c["name"] for c in chars], excel_names)
    # 配置表名单为空（首次搭建、配置表还没填角色）时，"未匹配"没有意义——
    # 不能因为没人可比就把所有角色全标成未匹配。
    if not excel_names:
        unmatched_names = []

    overview = build_overview(chars, name_map)
    relic_df, stat_key_seen, excel_info = build_relic_rows(chars, name_map, set_map, char_df)
    combo = build_set_combo(chars)
    skills = build_skills(chars)
    stat_dict = build_stat_dict(stat_key_seen)
    map_df = build_map_sheet(chars, name_map, set_map, unmatched_names, alias2full)
    check = build_check_sheet(chars, relic_df, stat_key_seen, unmatched_names, set_map, data)

    # 00_说明
    info_rows = [
        ("source", "输入文件", os.path.basename(JSON_PATH)),
        ("exported_at", "导出时间", data.get("exported_at")),
        ("schema_version", "导出格式版本", data.get("schema_version")),
        ("complete", "是否完整导出", data.get("complete")),
        ("characters", "角色数", len(chars)),
        ("pieces", "遗器件数", len(relic_df)),
        ("", "本表口径", "只搬运、不判定：除「有效次数/主词条」两列比对外，全部字段原样落表"),
        ("", "02页说明", "一件遗器一行；主词条与 4 条副词条都给了 key / 中文 / 数值 / 【强化次数】"),
        ("", "次数还原", "依据档位表(低/中/高三档)+面板截断规则反解，并受'每件合计 8 或 9 次'约束"),
        ("", "档位表来源", "NGA《遗器体力价值论》副词条强化档位表 + gamevika 遗器页单次强化值表"),
        ("", "有效判定", "仅 Excel 名单内的角色有「有效词条」定义，其余角色该列留空"),
        ("", "中文映射", "大小三围靠 percent 区分：小生命/大生命 等，见 05_词条字典"),
        ("", "套装简称", "取自「遗器相关」表；3 处项目表错字已在 06_映射对照 自动对齐"),
    ]
    info_df = pd.DataFrame(info_rows, columns=["键", "项目", "内容"])

    # 输出目录可能还不存在（比如刚清理过归档），先建出来
    os.makedirs(os.path.dirname(os.path.abspath(OUT_PATH)), exist_ok=True)

    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as w:
        info_df.to_excel(w, sheet_name="00_说明", index=False)
        overview.to_excel(w, sheet_name="01_角色总览", index=False)
        relic_df.to_excel(w, sheet_name="02_遗器明细", index=False)
        combo.to_excel(w, sheet_name="03_套装组合", index=False)
        skills.to_excel(w, sheet_name="04_技能", index=False)
        stat_dict.to_excel(w, sheet_name="05_词条字典", index=False)
        map_df.to_excel(w, sheet_name="06_映射对照", index=False)
        check.to_excel(w, sheet_name="07_校验", index=False)

    print("输出:", OUT_PATH)
    print("角色总览 %d 行 / 遗器明细 %d 行 / 套装组合 %d 行 / 技能 %d 行"
          % (len(overview), len(relic_df), len(combo), len(skills)))
    print("\n=== 07_校验 摘要 ===")
    for _, r in check.iterrows():
        print("  %-22s %-40s %s" % (r["检查项"], str(r["结果"])[:40], r["说明"]))


if __name__ == "__main__":
    main()
