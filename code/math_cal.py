# -*- coding: utf-8 -*-
"""
================================================================================
math_cal.py —— 遗器副词条概率的数学核心
================================================================================

【这个文件负责什么】
    给定"这个角色认哪些词条算有效"，算出**一件遗器强化到 +15 之后**，
    "有效词条数是 0 / 1 / 2 / ... / 9" 各自的概率。

【为什么这么算】
    星铁的遗器规则：
      1) 初始有 3 条或 4 条副词条（4 词条的概率 20%，即 e_equip34 = {3:0.8, 4:0.2}）；
      2) 副词条从 12 种里"带权不放回"抽取——抽走一条，它的权重就从池子里扣掉；
      3) 主词条同名的那一条不会作为副词条出现（所以在 main.py 里先把它剔出去）；
      4) +15 途中一共强化 5 次，每次在已有的那些副词条里**等概率**挑一条 +1。
         所以如果初始是 3 词条，前 1 次强化是"补第 4 条"，剩下 4 次是真正的强化；
         等价于"先当作 4 词条，再强化 4 次"——作者就是用这个等价变换简化问题的。

【整体流程】
    cal_p(有效词条列表, 无效词条列表)
      ├─ 第一步 cal_eff_when_init4()：算出"初始 4 条副词条里有 0~4 条有效"的概率分布
      │     └─ cal_p_at_need_eff()：枚举 C(有效,u)×C(无效,4-u) 的所有组合，把每组的概率加起来
      │           └─ cal_p_once()：一组具体词条（带各自权重）被抽到的概率
      │                 = Σ(24 种排列) Π(该条权重 / 抽取时池子剩余总权重)
      ├─ 第二步 强化命中：给"初始有 i 条有效"乘上"5 次/4 次强化里命中 j 次"的二项分布
      │     p_arr3（4 次强化，概率 i/4）与 p_arr4（5 次强化，概率 i/4）
      └─ 第三步 混合：0.8 ×（3 词条的结果）+ 0.2 ×（4 词条的结果）

【输出格式】
    长度 10 的 np.array，下标 k 表示"最终有 k 条有效词条"的概率，总和 = 1。

================================================================================
"""

import copy
import itertools
from itertools import combinations
import numpy as np


# ------------------------------------------------------------------------------
# 全局数据
# ------------------------------------------------------------------------------
# 副词条权重字典（与 main.py 里的 init_equip_res_dict() 必须保持一致）
equip_res = dict()
res_skill_list = []
# 初始词条数的概率：80% 出 3 词条，20% 出 4 词条（社区统计结论）
e_equip34 = {3: 0.8, 4: 0.2}

# 计算结果缓存：键是 (四条词条的权重升序排列..., 总权重)，值是这组词条被抽到的概率。
# 因为"不同角色/不同部位"会出现大量重复的权重组合，缓存能省很多重复计算。
cal_cache_dict = dict()


# ------------------------------------------------------------------------------
# 强化命中矩阵
# ------------------------------------------------------------------------------
# p_arr3：初始 3 词条的情形（等价于"4 条固定后强化 4 次"）
#   横坐标 = 初始有效词条数 i（0~4）
#   纵坐标 = 4 次强化里命中有效词条的次数 j（0~4）
#   每格 = C(4,j) * (i/4)^j * (1-i/4)^(4-j)，即二项分布 B(4, i/4)
#   含义：4 条副词条里有 i 条是有效的，每次强化等概率挑一条，命中有效的概率就是 i/4。
p_arr3 = np.array([[1, 0, 0, 0, 0],
                   [81 / 256, 108 / 256, 54 / 256, 12 / 256, 1 / 256],
                   [1 / 16, 4 / 16, 6 / 16, 4 / 16, 1 / 16],
                   [1 / 256, 12 / 256, 54 / 256, 108 / 256, 81 / 256],
                   [0, 0, 0, 0, 1]
                   ])

# p_arr4：初始 4 词条的情形，强化 5 次 → B(5, i/4)
#   横坐标 i = 0~4，纵坐标 j = 0~5
p_arr4 = np.array([
    [1, 0, 0, 0, 0, 0],
    [243 / 1024, 405 / 1024, 270 / 1024, 90 / 1024, 15 / 1024, 1 / 1024],
    [1 / 32, 5 / 32, 10 / 32, 10 / 32, 5 / 32, 1 / 32],
    [1 / 1024, 15 / 1024, 90 / 1024, 270 / 1024, 405 / 1024, 243 / 1024],
    [0, 0, 0, 0, 0, 1]
])

# 4 条词条的 24 种抽取顺序（3! × 4 = 24，这里直接用 itertools 生成 0~3 的全排列）
perm_order = list(itertools.permutations([0, 1, 2, 3], 4))


def init_equip_res_dict():
    """
    初始化副词条权重表（12 种词条）。
    权重设计：速度 40 最稀有 → 双暴 60 → 命中/抵抗/击破 80 → 大小攻防 100 最常见。
    这套权重直接决定了"双暴速度为什么难出"，是全部概率计算的基础。
    """
    res_map = [
        ('小生命', 100),
        ('小攻击', 100),
        ('小防御', 100),
        ('大生命', 100),
        ('大攻击', 100),
        ('大防御', 100),
        ('速度', 40),
        ('暴击', 60),
        ('暴伤', 60),
        ('效果抵抗', 80),
        ('效果命中', 80),
        ('击破', 80)
    ]
    for name, p in res_map:
        equip_res[name] = p
        res_skill_list.append(name)


# ==============================================================================
# 核心一：一组具体的 4 条词条，被抽到的概率
# ==============================================================================
def cal_p_once(skills, w_arr, w_total):
    """
    计算"恰好抽出这 4 条词条（不管顺序）"的总概率。

    参数：
        skills  这 4 条词条的名字（这个函数里其实没用到，只用于调试）
        w_arr   这 4 条词条的权重，**已按升序排好**（排序是为了让相同的组合只算一次，配合缓存）
        w_total 抽取开始前池子里的总权重

    做法：
        因为 4 条词条谁先谁后都可能，枚举 24 种顺序，
        每种顺序的概率 = Π(第 k 步抽到的那条权重 / 当时池子里的剩余总权重)。
        注意每一步都要把已抽走的权重从池子里扣掉（"不放回"）。

    严格来说，这里假设 4 条词条是"一条一条抽满"，而实际游戏里初始 3/4 词条的
    抽取过程与此等价（条数由 e_equip34 单独处理），所以这样算是自洽的。

    输出：这 4 条词条出现（任意顺序）的总概率。
    """
    # 四条词条权重的乘积（分子部分，每种排列都用同一个分子）
    w_mu = w_arr[0] * w_arr[1] * w_arr[2] * w_arr[3]

    p_list = []
    for i in range(len(perm_order)):
        p = w_mu
        res = w_total                      # 当前池子剩余总权重
        for j in range(4):
            # 概率 = 该条权重 / 剩余总权重；然后把它从池子里扣掉
            p = p / res
            res -= w_arr[perm_order[i][j]]
        p_list.append(p)

    # 24 种顺序的概率加起来 = 这 4 条词条出现的总概率
    return sum(p_list)


# ==============================================================================
# 核心二：枚举组合
# ==============================================================================
def cal_p_at_need_combine(use_name_list, res_name_list):
    """
    【给"同一部位多个角色共用"的场景用】（main.py 的 cal_combine_eq_p 调用）

    枚举"4 条副词条里恰好有 u 条有效"的所有具体组合，
    返回 [(u, 这组词条, 缓存键, 该组出现的概率), ...]。

    为什么要保留"具体组合"而不是只保留数量：
        两个角色可能共用同一件遗器，但有效词条集合不同；
        某一组具体词条（比如 暴击+速度+小攻击+小防御）对 A 有效、对 B 无效。
        必须按具体组合去判断"对谁有用"，才能避免同一件遗器被重复计数。

    ⚠ 注意：u 从 1 开始枚举（range(1, ...)），漏掉了 u=0 即"4 条全无效"的组合。
      这对"当前 0 条有效词条"的部位会略微低估概率。
    """
    ans = []

    # 先算出当前"有效 + 无效"候选池的总权重
    weight_total = 0
    for su in use_name_list:
        weight_total += equip_res[su]
    for sr in res_name_list:
        weight_total += equip_res[sr]

    # u 最多取到 min(4, 有效词条种类数)
    for useful_num in range(1, min(5, len(use_name_list) + 1)):
        combine_use_list = list(combinations(use_name_list, useful_num))       # 选哪 u 条有效
        combine_res_list = list(combinations(res_name_list, 4 - useful_num))   # 再配 4-u 条无效

        for cu in combine_use_list:
            for cr in combine_res_list:
                skills = cu + cr

                # 取出这 4 条各自的权重，并升序排（排序 + 缓存键，保证同一组只算一次）
                w_arr = []
                for s in skills:
                    w_arr.append(equip_res[s])
                w_arr.sort()

                # 准备缓存键：权重序列 + 总权重
                ck = copy.deepcopy(w_arr)
                ck.append(weight_total)
                ck = tuple(ck)

                # 查缓存
                if ck in cal_cache_dict:
                    p = cal_cache_dict[ck]
                else:
                    p = cal_p_once(skills, w_arr, weight_total)
                    cal_cache_dict[ck] = p

                ans.append((useful_num, skills, ck, p))
    return ans


def cal_p_at_need_eff(use_name_list, res_name_list, equip_name_weight_dict, useful_num, res_num):
    """
    "初始 4 条副词条里，恰好有 useful_num 条有效（res_num 条无效）"的概率。

    做法：枚举所有 C(有效, useful_num) × C(无效, res_num) 的组合，
          把每组的出现概率加起来。

    参数 equip_name_weight_dict 目前没用（权重走的是全局 equip_res），保留是为了接口兼容。
    """
    combine_use_list = list(combinations(use_name_list, useful_num))
    combine_res_list = list(combinations(res_name_list, res_num))

    # （以下 comb_total_numb 是"组合总数"的展示用计算，未被后续逻辑使用，属遗留代码）
    t_numb = len(use_name_list) + len(res_name_list)

    comb_total_numb = t_numb / 24
    for i in range(1, 4):
        comb_total_numb *= t_numb - i

    # 当前候选池的总权重
    weight_total = 0
    for su in use_name_list:
        weight_total += equip_res[su]
    for sr in res_name_list:
        weight_total += equip_res[sr]

    # 逐个组合算概率并累加
    p_total_list = []
    for cu in combine_use_list:
        for cr in combine_res_list:
            skills = cu + cr
            w_arr = []
            for s in skills:
                w_arr.append(equip_res[s])
            w_arr.sort()

            # 缓存键（同上）
            ck = copy.deepcopy(w_arr)
            ck.append(weight_total)
            ck = tuple(ck)

            if ck in cal_cache_dict:
                p = cal_cache_dict[ck]
            else:
                p = cal_p_once(skills, w_arr, weight_total)
                cal_cache_dict[ck] = p

            p_total_list.append((skills, ck, p))

    p_final_total = 0
    for sk, w, p in p_total_list:
        p_final_total += p

    return p_final_total


# ==============================================================================
# 核心三：初始有效词条数量分布
# ==============================================================================
def cal_eff_when_init4(use_name_list, res_name_list, equip_name_weight_dict):
    """
    初始 4 条副词条里有效条数（0~4）的概率分布。

    输出：长度 5 的 np.array，下标 k = "初始有 k 条有效"的概率，总和 = 1。
    """
    p_list = []
    for i in range(5):
        # 恰好 i 条有效、4-i 条无效
        ret = cal_p_at_need_eff(use_name_list, res_name_list, equip_name_weight_dict, i, 4 - i)
        p_list.append(ret)
    while len(p_list) < 5:
        p_list.append(0)
    return np.array(p_list)


# ==============================================================================
# 核心四：给定初始有效条数，练满后的分布
# ==============================================================================
def cal_p_when_init_val_num(val_num):
    """
    【已知初始有 val_num 条有效词条】时，练满（+15）后 0~9 条有效词条的分布。

    用途：multi-角色共用一个部位时，先固定"初始几条有效"，再分别套 3 词条 / 4 词条两条路径。

    计算：
        p_init_dist = 只有一个点（val_num 处为 1）的分布
        3 词条路径：分布 × p_arr3 → 沿对角线累加，得到 0~8 的分布
        4 词条路径：分布 × p_arr4 → 沿对角线累加，得到 0~9 的分布
        最后 0.8 × 三词条 + 0.2 × 四词条
    """
    # 初始分布：只有一个确定值
    p_init3_dist = np.array([0, 0, 0, 0, 0])
    p_init3_dist[val_num] = 1
    p_init4_dist = copy.deepcopy(p_init3_dist)

    # region 计算三词条概率分布（可强化 4 次）
    p_arr3_copy = copy.deepcopy(p_arr3)
    p_dict3 = np.zeros(10)
    for init_hit_num in range(5):
        p_arr3_copy[init_hit_num] *= p_init3_dist[init_hit_num]
    for i in range(5):
        for j in range(5):
            p_dict3[i + j] += p_arr3_copy[i][j]     # 初始 i 条 + 强化命中 j 条 = 最终 i+j 条
    # endregion

    # region 计算四词条概率分布（可强化 5 次）
    p_arr4_copy = copy.deepcopy(p_arr4)
    p_dict4 = np.zeros(10)
    for init_hit_num in range(5):
        p_arr4_copy[init_hit_num] *= p_init4_dist[init_hit_num]
    for init_hit_num in range(5):
        for j in range(6):
            p_dict4[init_hit_num + j] += p_arr4_copy[init_hit_num][j]   # 同理，最多 4+5=9
    # endregion

    # 按 80% / 20% 混合
    p_total = e_equip34[3] * p_dict3 + e_equip34[4] * p_dict4

    return p_total


# ==============================================================================
# 对外主函数
# ==============================================================================
'''
输入为：有用的词条列表，其他可选副词条列表、每个词条对应的权重
输出为：【p0,p1,p2,p3,p4,p5,p6,p7,p8,p9】，为最终获取得到的装备中有0-9个有效词条的概率
'''
def cal_p(use_name_list, res_name_list, equip_name_weight_dict):
    """
    ★ 整个项目的数学入口 ★

    参数：
        use_name_list           该角色认的有效副词条，如 ['暴击','暴伤','大生命']
        res_name_list           剩下的"无效"候选副词条
                                （调用方已经剔除了主词条同名的那一条）
        equip_name_weight_dict  权重表（目前用全局的 equip_res，此参数保留兼容）

    返回：长度 10 的 np.array，下标 k = 练满后恰好有 k 条有效词条的概率。

    步骤：
        1) 算初始 4 条里有效条数的分布 p_init_dist
        2) 3 词条路径：p_init_dist × p_arr3（4 次强化），按 i+j 累加
        3) 4 词条路径：p_init_dist × p_arr4（5 次强化），按 i+j 累加
        4) 0.8 × 路径2 + 0.2 × 路径3
    """
    # ① 初始有效条数分布（3 词条和 4 词条在"初始 4 条里有效几条"这件事上分布相同）
    p_init3_dist = cal_eff_when_init4(use_name_list, res_name_list, equip_name_weight_dict)
    p_init4_dist = copy.deepcopy(p_init3_dist)

    # region ② 三词条：强化 4 次
    p_arr3_copy = copy.deepcopy(p_arr3)
    p_dict3 = np.zeros(10)
    for init_hit_num in range(5):
        p_arr3_copy[init_hit_num] *= p_init3_dist[init_hit_num]
    for i in range(5):
        for j in range(5):
            p_dict3[i + j] += p_arr3_copy[i][j]
    # endregion

    # region ③ 四词条：强化 5 次
    p_arr4_copy = copy.deepcopy(p_arr4)
    p_dict4 = np.zeros(10)
    for init_hit_num in range(5):
        p_arr4_copy[init_hit_num] *= p_init4_dist[init_hit_num]
    for init_hit_num in range(5):
        for j in range(6):
            p_dict4[init_hit_num + j] += p_arr4_copy[init_hit_num][j]
    # endregion

    # ④ 混合
    p_total = e_equip34[3] * p_dict3 + e_equip34[4] * p_dict4

    return p_total


# ==============================================================================
# 核心五：定向（指定主词条 / 再锁定 f 条副词条）时的分布
# ==============================================================================
# 定向的成本阶梯（本项目里约定的口径）：
#   1 cost  → 只指定主词条，副词条全随机        → fixed f = 0
#   2 cost  → 指定主词条 + 锁定 1 条有效副词条  → fixed f = 1
#   5 cost  → 指定主词条 + 锁定 2 条有效副词条  → fixed f = 2
#
# 被锁定的 f 条是"必定命中"的有效词条，所以：
#   候选池 = (有效 ∪ 无效) − 已锁的 f 条
#   还要再抽 4 − f 条（带权、不放回），其中恰好 u 条有效
#   初始有效词条数 = f + u
#
# 一个容易踩的细节：强化时"命中有效"的概率是 (f+u)/4，**分母始终是 4**。
# 因为初始 3 词条的那件遗器，第一次强化会补成第 4 条副词条，
# 之后每次强化都是在完整的 4 条里等概率挑一条 —— 所以分母是 4 而不是 3。
# （这也正是原 cal_p 把 3 词条等价成"4 条 + 强化 4 次"的原因。）

# 缓存：pool_u / pool_r 的权重组合决定了一切，同构的部位会大量复用
_fixed_dist_cache = dict()


def cal_p_set(w_arr, w_total):
    """
    任意条数 d（1~4）的一组词条"被抽到"的总概率（不分先后顺序）。

    d = 4 时就是 cal_p_once 做的事；这里做成通用版，
    是因为定向锁 1~2 条之后只剩下 2~3 条要抽，cal_p_once 里硬编码的 4 条会越界。
    """
    d = len(w_arr)
    if d == 0:
        return 1.0

    w_mu = 1
    for w in w_arr:
        w_mu *= w

    p_total = 0.0
    for perm in itertools.permutations(range(d)):
        p = w_mu
        res = w_total                     # 当前池子剩余总权重
        for idx in perm:
            p = p / res                   # 该条权重 / 剩余总权重
            res -= w_arr[idx]             # 不放回：抽走后从池子里扣掉
        p_total += p
    return p_total


def cal_p_with_fixed(use_name_list, res_name_list, fixed_list):
    """
    定向后的"练满后 0~9 条有效词条"分布。

    参数：
        use_name_list  该角色该部位认的有效副词条（调用方已剔除主词条同名的那条）
        res_name_list  无效候选池（调用方已剔除主词条同名的那条）
        fixed_list     被"锁定"的有效副词条名（0~2 条）

    返回：长度 10 的 np.array，下标 k = 练满后恰好 k 条有效词条的概率。

    fixed_list 为空时，本函数与 cal_p() 完全等价（已做过对齐验证，见 README）。
    """
    f = len(fixed_list)
    pool_u = [s for s in use_name_list if s not in fixed_list]      # 剩下的有效候选
    pool_r = [s for s in res_name_list if s not in fixed_list]      # 剩下的无效候选
    k = 4 - f                                                       # 还要抽几条副词条

    # 缓存键：f + 两个池子各自的权重多重集（同构部位直接复用结果）
    key = (f,
           tuple(sorted(equip_res[s] for s in pool_u)),
           tuple(sorted(equip_res[s] for s in pool_r)))
    if key in _fixed_dist_cache:
        return _fixed_dist_cache[key].copy()

    w_total = sum(equip_res[s] for s in pool_u) + sum(equip_res[s] for s in pool_r)

    # ① 剩下的 k 条里恰好 u 条有效的概率分布
    p_u = np.zeros(k + 1)
    if w_total > 0:
        for u in range(k + 1):
            if u > len(pool_u) or (k - u) > len(pool_r):
                continue
            for cu in combinations(pool_u, u):
                for cr in combinations(pool_r, k - u):
                    w_arr = sorted(equip_res[s] for s in cu + cr)
                    p_u[u] += cal_p_set(w_arr, w_total)

    # ② 初始有效条数 = f + u，再分别走 3 词条（强化 4 次）/ 4 词条（强化 5 次）
    p_dict3 = np.zeros(10)
    p_dict4 = np.zeros(10)
    for u in range(k + 1):
        init = f + u                    # ≤ 4，p_arr3 / p_arr4 的行下标都在范围内
        for j in range(5):
            p_dict3[init + j] += p_u[u] * p_arr3[init][j]
        for j in range(6):
            p_dict4[init + j] += p_u[u] * p_arr4[init][j]

    # ③ 80% 三词条 + 20% 四词条
    p_total = e_equip34[3] * p_dict3 + e_equip34[4] * p_dict4

    _fixed_dist_cache[key] = p_total.copy()
    return p_total


# ==============================================================================
# 单元自测：直接 python math_cal.py 可以看到几组典型词条的分布
# ==============================================================================
def main():
    q = [
        (['速度'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '大攻击', '暴击', '暴伤']),
        (['速度', '大攻击'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '暴击', '暴伤']),
        (['速度', '大攻击', '暴击'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '暴伤']),
        (['速度', '大攻击', '暴击', '暴伤'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
        # 下面几组是作者调试时留的，可自行打开对比
        # (['速度', '大攻击', '暴击', '暴伤'], ['小生命', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
        # (['速度', '大攻击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
        # (['速度', '暴击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
    ]

    for quse, qres in q:
        p_total = cal_p(quse, qres, equip_res)
        print('-----------------------------------')
        print(quse, qres)
        print('最终满级不同有效词条数量的概率分布为 ', p_total)
        break    # 只打印第一组（想看全部就删掉这行）

    return


# 模块被 import 时就把权重表初始化好（main.py 依赖这一点）
init_equip_res_dict()

if __name__ == "__main__":
    main()
