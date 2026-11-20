# -*- coding: utf-8 -*-
"""
================================================================================
main.py —— 星铁遗器刷取分析主程序
================================================================================

【这个程序干什么】
    读入「星铁装备刷取-自动生成-豆包.xlsx」（由 01-一键导出.bat 从 characters.json 生成）
    里每个角色当前的遗器情况，算出"继续刷某个遗器本，某件遗器变得更强的概率"，并换算成
    "刷多少次 / 多少天，才有 90%（或 95%）把握出一件更好的"。

【核心公式（单件遗器、单次刷取）】
    P = P_副词条达标  ×  P_主词条  ×  P_部位  ×  1.05

      1) P_副词条达标：由 math_cal.py 计算，输出了 0~9 条有效词条的完整概率分布，
         本文件只取"比当前词条数更多"的那一段求和。
      2) P_主词条：衣服/鞋子/球/绳子各有自己的主词条概率表（见 init_equip_main_dict）。
         头/手的主词条是固定的（小生命/小攻击），概率为 1。
      3) P_部位：外圈 4 件套里某个部位 = 0.25；内圈 2 件套里某个部位 = 0.5。
         （这里算的是"已经掉了该套装之后，落在哪个部位"的条件概率）
      4) 1.05 = 每刷一次平均掉 2.1 件金遗器（均衡6） × 0.5（一个本里两个套装各占一半）
         → 含义是"每刷一次，期望掉落 1.05 件【目标套装】的遗器"。
         也就是说：**套装 50% 已经折进这个 1.05 里了，不需要再单独乘 0.5**。
         所以这里算出的概率单位是"每次 40 体力刷取"，期望次数 = 刷本次数。

【数据流】
    星铁装备刷取-自动生成-豆包.xlsx（01 脚本从 characters.json 生成）
      ├─ 角色遗器  → read_excel_character_info() → character_equip 对象列表
      │               （套装简称在这里被 update_info() 换成全名）
      └─ 遗器相关  → read_excel_equip_info()     → 简称→全名、全名→刷取地点 两张映射表
            ↓
    update_character_equip_anylize() 逐角色、逐部位算概率分布 → 存进对象的 eq4_anylize / eq2_anylize
            ↓
    四类结论（全部写进 分析结果-豆包.xlsx）
      ├─ 角色套装优化概率       ：某角色换外圈/内圈，任一部位变好的概率
      ├─ 角色套装部位优化概率   ：某角色某部位变好的概率（最细）
      ├─ 套装优化准确概率       ：按（套装,部位,主词条）分组去重后的概率（v2.0 主口径）
      └─ 地点优化概率           ：以"一个遗器本"为单位，把本内两个套装概率相加

【术语】
    有效词条：对某个角色有用的副词条（如"暴击、暴伤、大生命"），由人工在 Excel 里配置。
    有效词条数：一件遗器强化满（+15）之后，4 条副词条里有几条落在"有效词条"里。
    主词条：衣服/鞋子/球/绳子必须指定的那一条（头=小生命、手=小攻击是固定的）。
================================================================================
"""

import pandas as pd
import numpy as np
import copy
import os
import sys
import math                     # 组合数 math.comb 用（分解建议算二项分布）
import math_cal

# ------------------------------------------------------------------------------
# 全局缓存 / 映射表（模块级变量，由各个 init / read 函数填充）
# ------------------------------------------------------------------------------
# 装备简称 → 装备全称，例如 '诗人' → '哀歌覆国的诗人'
equip_name_dict = dict()
# 装备全称 → 刷取地点（遗器本名称），例如 '哀歌覆国的诗人' → 'xxx之径'
equip_from_dict = dict()
# 装备全称 → 需求该装备的角色列表 [(角色名, 该角色套装提升概率), ...]，v1 的"旧口径"用到
equip_match_character_dict = dict()
# 各部位的主词条出货概率，键为 (部位, 主词条名)，值为概率
equip_main = dict()
# 副词条权重表，键为词条名，值为权重（如 速度=40、暴击=60、大攻击=100）
equip_res = dict()
# 12 条副词条的名字列表（顺序固定，后面做"从候选池里剔除"的操作会用到）
res_skill_list = []


# ==============================================================================
# 角色遗器信息载体
# ==============================================================================
class character_equip():
    """
    一个角色一行数据。

    对应 Excel「角色遗器」表的一行：
        角色名称 | 遗器套装名称 | 内圈套装名称 | 有效词条数量 | 有效词条
        | 衣服主词条 | 鞋子主词条 | 球主词条 | 绳子主词条
        | 头 | 手 | 衣服 | 鞋子 | 球 | 绳子 | 词条数量(合计，可校验用)

    其中：
        eq4         外圈套装名列表（支持 2+2，所以是 list，可能是 2 个元素）
        eq2         内圈套装名（字符串，1 个元素）
        val_name_list  该角色的"有效词条"列表，如 ['暴击','暴伤','大生命']
        val_num_xxx    各部位**当前**的有效词条数量（人工数出来的）
        eq4_anylize / eq2_anylize  各部位"刷到并练满后 0~9 条有效词条"的概率分布
    """

    def __init__(self, name, eq4, eq2, val, val_list, main_yifu, main_xiezi, main_qiu, main_shengzi
                 , val_num_tou, val_num_shou, val_num_yifu, val_num_xiezi, val_num_qiu, val_num_shengzi
                 , weight=1.0, star=""):
        self.name = name          # 角色名
        # 角色星级（4 / 5），来自配置表「角色有效词条」页的人工填写列。
        # 只用于输出表里的筛选（"只看五星"），**不参与任何概率计算**。
        # 没填就是空字符串，不会默认猜成五星。
        self.star = "" if star is None or (isinstance(star, float) and pd.isna(star)) else str(star).strip()
        self.eq4 = eq4            # 外圈套装（list，2+2 时有两个元素）
        self.eq2 = eq2            # 内圈套装（str）
        self.val_num = val        # Excel 里的"有效词条数量"（有效词条种类数，实际未参与计算）
        self.val_name_list = val_list  # 有效词条种类列表

        # 四个会出主词条的部位，其"要求的主词条"
        self.main_yifu = main_yifu        # 衣服主词条
        self.main_xiezi = main_xiezi      # 鞋子主词条
        self.main_qiu = main_qiu          # 位面球主词条
        self.main_shengzi = main_shengzi  # 连结绳主词条

        # 各部位"当前"有效词条数量（人工从米游社界面上数的）
        self.val_num_tou = val_num_tou
        self.val_num_shou = val_num_shou
        self.val_num_yifu = val_num_yifu
        self.val_num_xiezi = val_num_xiezi
        self.val_num_qiu = val_num_qiu
        self.val_num_shengzi = val_num_shengzi
        # 打包成 list，顺序固定为：头、手、衣服、鞋子、球、绳子
        self.val_num_list = [val_num_tou, val_num_shou, val_num_yifu, val_num_xiezi, val_num_qiu, val_num_shengzi]

        # ------------------------------------------------------------------
        # 角色权重（来自 Excel「角色遗器」表第 17 列「权重」，没填就是 1.0）
        #   weight_raw : 用户在 Excel 里填的原值（只看它的数值本身）
        #   weight     : 归一化之后的值（mean = 1），下游一律用这个
        #
        # 为什么要归一化：见 normalize_character_weight() 的说明。
        # 重要：权重【不参与】任何概率分布的计算，只在"最后把多个角色的结果
        #      加起来"时作为系数出现。概率链（equip_anylize / cal_p /
        #      p_better / cal_combine_eq_p）完全不看这个字段。
        # ------------------------------------------------------------------
        self.weight_raw = float(weight)
        self.weight = float(weight)

        # 计算完成后填充：
        # eq4_anylize 是 4 个 np.array（头/手/衣服/鞋子的概率分布）
        # eq2_anylize 是 2 个 np.array（球/绳子的概率分布）
        self.eq4_anylize = []
        self.eq2_anylize = []

    @property
    def total_val(self):
        """六个部位当前有效词条数的合计（只用于展示/校验）"""
        return self.val_num_tou + self.val_num_shou + self.val_num_yifu + self.val_num_xiezi + self.val_num_qiu + self.val_num_shengzi

    def update_info(self):
        """
        把 Excel 里手填的【套装简称】统一换成【套装全称】，并顺便查出该套装在哪个本刷。

        为什么要换：后面所有的分组、统计都按全称做 key，
        否则"诗人"和"哀歌覆国的诗人"会被当成两套装备。
        """
        self.pos = []                    # 每套装备对应的刷取地点
        for i in range(len(self.eq4)):   # 外圈（可能 2+2，所以循环）
            name = self.eq4[i]
            ori_name = equip_name_dict[name]      # 简称 → 全称
            pos = equip_from_dict[ori_name]       # 全称 → 刷取地点
            self.eq4[i] = ori_name
            self.pos.append(pos)
        name = self.eq2
        ori_name = equip_name_dict[name]          # 内圈同理
        pos = equip_from_dict[ori_name]
        self.eq2 = ori_name
        self.pos.append(pos)
        return

    def p_get_better6(self):
        """
        算出六件遗器各自的"能刷到更好的"概率。

        做法：概率分布是"练满后有 k 条有效词条"的概率（k=0..9），
              当前该部位已有 val 条，那么"更好"= k >= val+1，
              也就是把分布里下标 val+1 到末尾全部加起来。

        结果存进 self.p_better，长度 6，顺序：头、手、衣服、鞋子、球、绳子。
        """
        ret = []
        for i in range(4):                                   # 外圈四件：头/手/衣服/鞋子
            eq_anylize_list = self.eq4_anylize[i]
            eq_val = self.val_num_list[i]                    # 该部位当前有效词条数
            p_better = sum(eq_anylize_list[eq_val + 1:])     # 比当前更多的概率之和
            ret.append(p_better)

        for i in range(2):                                   # 内圈两件：球/绳子
            eq_anylize_list = self.eq2_anylize[i]
            eq_val = self.val_num_list[4 + i]
            p_better = sum(eq_anylize_list[eq_val + 1:])
            ret.append(p_better)

        self.p_better = ret
        return

    def load_eff_num(self, eff_num):
        """
        存下"各部位理论上最多能有几条有效词条"（用于输出表里的"理论上限"）。
        eff_num 顺序：头、手、衣服、鞋子、球、绳子。
        """
        self.eff_num_list = eff_num


# ==============================================================================
# 初始化：主词条概率表 / 副词条权重表
# ==============================================================================
def init_equip_main_dict():
    """
    初始化"部位 + 主词条 → 出货概率"的字典。

    概率来源：社区统计（25000 件样本），与「主词条分布」那张表一致：
        衣服：生命/攻击/防御 各 20%，暴击/暴伤/治疗/命中 各 10%
        鞋子：生命/攻击/防御 各 29.17%，速度 12.5%
        绳子：生命/攻击/防御 各 26.33%，击破 15%，充能 6%
        球  ：生命/攻击/防御 各 12.33%，7 种属性伤各 9%（合计 63%）
    头部/手部的主词条固定（小生命/小攻击），概率为 1。

    ⚠ 注意：这里是硬编码，Excel 里那张「主词条分布」表目前没被读取。
      如果以后游戏调整概率，要改的是这个函数（或者改成读表）。
    """
    yifu_list = [['大生命', 0.2], ['大攻击', 0.2], ['大防御', 0.2],
                 ['暴击', 0.1], ['暴伤', 0.1], ['治疗量', 0.1], ['效果命中', 0.1]]
    xiezi_list = [['大生命', 0.2917], ['大攻击', 0.2917], ['大防御', 0.2917],
                  ['速度', 0.125]]
    shengzi_list = [['大生命', 0.2633], ['大攻击', 0.2633], ['大防御', 0.2633],
                    ['击破', 0.15], ['充能', 0.06]]
    qiu_list = [['大生命', 0.1233], ['大攻击', 0.1233], ['大防御', 0.1233]
        , ['冰伤', 0.09], ['火伤', 0.09], ['物伤', 0.09], ['虚数伤害', 0.09], ['量子伤', 0.09], ['雷伤', 0.09],
                ['风伤', 0.09]]

    # 头/手的主词条是固定的，概率 1
    equip_main[('头', '小生命')] = 1
    equip_main[('手', '小攻击')] = 1

    for name, p in xiezi_list:
        equip_main[('鞋子', name)] = p
    for name, p in yifu_list:
        equip_main[('衣服', name)] = p
    for name, p in qiu_list:
        equip_main[('球', name)] = p
    for name, p in shengzi_list:
        equip_main[('绳子', name)] = p


def init_equip_res_dict():
    """
    初始化副词条权重表（并登记 12 条副词条的名字）。

    权重含义：游戏在生成副词条时是"带权不放回抽取"，
    权重越低的词条越难出。速度 40 最稀有，双暴 60 其次，命中/抵抗/击破 80，
    其他大小攻防 100 最常见。

    ⚠ math_cal.py 里也有一份一模一样的权重表（两份需保持一致）。
    """
    res_map = [
        ('小生命', 100),
        ('小攻击', 100),
        ('小防御', 100),
        ('大生命', 100),
        ('大攻击', 100),
        ('大防御', 100),
        ('速度', 40),      # 最稀有
        ('暴击', 60),      # 稀有
        ('暴伤', 60),      # 稀有
        ('效果抵抗', 80),
        ('效果命中', 80),
        ('击破', 80)
    ]
    for name, p in res_map:
        equip_res[name] = p
        res_skill_list.append(name)


# ==============================================================================
# 读 Excel
# ==============================================================================
def read_excel_equip_info(excel_file_path):
    """
    读「遗器相关」表，建立两张映射：
        equip_from_dict[全称] = 刷取地点
        equip_name_dict[简称] = 全称     （一套装备可以有多个简称，用中文顿号「、」分隔）
        equip_match_character_dict[全称] = []  （预置空列表，后面 v1 口径往里塞角色）

    表结构：遗器全称 | 刷取位置 | 简称1
    """
    # 读取Excel
    df = pd.read_excel(excel_file_path, sheet_name='遗器相关')
    # 读取行名称
    attr_list = df.columns
    # 转化为numpy便于理解，dataframe没怎么用过
    character_list = df.to_numpy()

    for ori_name, pos, sp_name in character_list:
        equip_from_dict[ori_name] = pos
        equip_match_character_dict[ori_name] = []
        # 全称自己也登记成键：这样表里写简称（诗人）还是官方全称（哀歌覆国的诗人）都能查。
        # 用 setdefault，避免某套的简称恰好等于另一套的全称时互相顶掉。
        equip_name_dict.setdefault(ori_name, ori_name)
        if '、' in sp_name:            # 多个简称
            sp = sp_name.split('、')
            for i in sp:
                equip_name_dict[i] = ori_name
        else:                          # 单个简称
            equip_name_dict[sp_name] = ori_name

    return


def exchang_character_info(c_info):
    """
    把「角色遗器」表的一行（numpy 的一行）转成 character_equip 对象。

    参数 c_info 的列顺序（与 Excel 一致）：
        0 角色名称 | 1 遗器套装名称 | 2 内圈套装名称 | 3 有效词条数量 | 4 有效词条
        5 衣服主词条 | 6 鞋子主词条 | 7 球主词条 | 8 绳子主词条
        9 头 | 10 手 | 11 衣服 | 12 鞋子 | 13 球 | 14 绳子 | 15 合计(可校验)
    """
    c_name = c_info[0]
    eq4 = c_info[1]
    if '+' in eq4:                     # 外圈支持 2+2 组合，用「+」分隔
        c_eq4 = eq4.split('+')
    else:
        c_eq4 = [eq4]
    c_eq2 = c_info[2]
    c_val = int(c_info[3])
    c_val_list = c_info[4].split('、')  # 有效词条用「、」分隔

    main_yifu = c_info[5]
    main_xiezi = c_info[6]
    main_qiu = c_info[7]
    main_shengzi = c_info[8]

    val_num_tou = c_info[9]
    val_num_shou = c_info[10]
    val_num_yifu = c_info[11]
    val_num_xiezi = c_info[12]
    val_num_qiu = c_info[13]
    val_num_shengzi = c_info[14]
    total_val_num = val_num_tou + val_num_shou + val_num_yifu + val_num_xiezi + val_num_qiu + val_num_shengzi

    # 权重列（第 17 列，index=16）。
    # 兼容处理：老版本的表没有这一列，或者某一格没填，都按"平权 1.0"处理，
    # 这样旧表直接丢进来也能跑，不会因为缺列报错。
    c_weight = 1.0
    if len(c_info) > 16 and not pd.isna(c_info[16]):
        c_weight = float(c_info[16])

    # 星级列（第 18 列，index=17，由 01 从配置表「角色星级」写过来）。
    # 同样是"有就用、没有就空"：老表没有这一列，也不会报错。
    c_star = ""
    if len(c_info) > 17:
        v = c_info[17]
        c_star = "" if pd.isna(v) else str(v).strip()

    ch_eq = character_equip(c_name, c_eq4, c_eq2, c_val, c_val_list, main_yifu, main_xiezi, main_qiu, main_shengzi
                 , val_num_tou, val_num_shou, val_num_yifu, val_num_xiezi, val_num_qiu, val_num_shengzi
                 , c_weight, c_star)
    ch_eq.update_info()      # 简称 → 全称
    return ch_eq


def read_excel_character_info(excel_file_path):
    """
    读「角色遗器」表，逐行转成 character_equip 对象，返回对象列表。
    这是整个程序的"输入入口"——所有后续计算都基于这里读到的 55 个角色。
    """
    ret_list = []

    df = pd.read_excel(excel_file_path, sheet_name='角色遗器')
    attr_list = df.columns
    character_list = df.to_numpy()
    name_list = character_list[:, 0]

    for i in range(len(character_list)):
        c_info = character_list[i]
        c_data = exchang_character_info(c_info)
        ret_list.append(c_data)

    return ret_list


# ==============================================================================
# 单件遗器的概率分析
# ==============================================================================
# 记录"已计算过哪些（有效词条数, 无效词条数）组合"，运行结束时打印，便于评估计算量
repp_list = set()


def equip_anylize_cal(use_list, res_list):
    """
    真正去算"练满后 0~9 条有效词条"的概率分布。
    直接委托给 math_cal.cal_p()（那边是整个算法的核心）。
    """
    repp_list.add((len(use_list), len(res_list)))
    ret = math_cal.cal_p(use_name_list=use_list, res_name_list=res_list, equip_name_weight_dict=equip_res)

    return ret


def equip_anylize(main_skill, useful_skills, now_skill_num, pos_name):
    """
    单个部位："每次刷本，掉出一件【该部位+该主词条+练满后有效词条数超过当前】"的概率。

    输入：
        main_skill      该部位要求的主词条（头=小生命、手=小攻击，其余来自 Excel）
        useful_skills   该角色的有效词条种类列表
        now_skill_num   该部位当前的有效词条数
        pos_name        部位名（'头'/'手'/'衣服'/'鞋子'/'球'/'绳子'）
    返回：
        np.array，长度 10，表示"练满后 0~9 条有效词条"的概率（已乘上主词条/部位/出金系数）
    """
    # ① 主词条与副词条互斥：主词条同名的那一条不会出现在副词条里，
    #    所以先从"有效词条"里删掉它，再从"全部副词条"里删掉它。
    re_useful = copy.deepcopy(useful_skills)
    if main_skill in re_useful:
        re_useful.remove(main_skill)

    tmp_skill_list = copy.deepcopy(res_skill_list)
    if main_skill in tmp_skill_list:
        tmp_skill_list.remove(main_skill)
    for sk in re_useful:
        if sk in tmp_skill_list:
            tmp_skill_list.remove(sk)      # 剩下的就是"无效候选池"

    # ② 只针对副词条，算 0~9 条有效词条的概率分布
    p_list = equip_anylize_cal(re_useful, tmp_skill_list)

    # ③ 乘上该主词条的出货概率（头/手为 1）
    p_list_add_main = p_list * equip_main[(pos_name, main_skill)]

    # ④ 乘上"掉的是哪个部位"的条件概率：外圈 4 件套取 0.25，内圈 2 件套取 0.5
    if pos_name in ['头', '手', '衣服', '鞋子']:
        p_list_add_main *= 0.25
    else:
        p_list_add_main *= 0.5

    # ⑤ 乘上出金系数 1.05 = 2.1 件/次 × 0.5（两个套装各半）
    #    到这里概率单位就是"每次 40 体力刷取"
    return p_list_add_main * 1.05


def cal_slot_directed_rate(useful_skills, main_skill, now_skill_num, fixed_num):
    """
    定向投放后，该部位"练满时有效词条数超过当前"的概率。

    fixed_num 对应 cost 阶梯：
        0 → 1 cost（只指定主词条）
        1 → 2 cost（主词条 + 锁 1 条副词条）
        2 → 5 cost（主词条 + 锁 2 条副词条）

    和 equip_anylize() 的关键区别：
        equip_anylize 算的是"每次 40 体力刷本、随机掉一件"的概率，
        所以乘了主词条概率 × 部位概率 × 出金系数 1.05。
        定向是"已经指定了这份，只看它强化完能不能超过现在"，
        所以这些系数一个都不该乘 —— 这里是纯粹的副词条概率。

    锁哪几条：在该角色该部位的有效副词条里，挑**权重最低**的。
        权重低 = 天然最难出（速度 40 < 双暴 60 < 命中/抵抗/击破 80 < 大/小攻防 100），
        锁定它们收益最大：既白拿一条稀有词条，又让剩下要抽的池子里全是高权重的
        容易词条，于是"抽到有效"的概率也更高。两个好处同向，所以无需再比较其他组合。
        （已用穷举验证：对同一配置，锁权重最低的组合，提升概率严格最高。）

    返回 (概率, 实际锁定的词条列表)；有效副词条不够锁时返回 (None, 已锁到的)。
    """
    # ① 主词条与副词条互斥（和 equip_anylize ①② 完全同一套逻辑）
    re_useful = copy.deepcopy(useful_skills)
    if main_skill in re_useful:
        re_useful.remove(main_skill)

    tmp_skill_list = copy.deepcopy(res_skill_list)
    if main_skill in tmp_skill_list:
        tmp_skill_list.remove(main_skill)
    for sk in re_useful:
        if sk in tmp_skill_list:
            tmp_skill_list.remove(sk)

    # ② 选要锁的副词条：按权重从低到高
    lock_candidates = sorted(re_useful, key=lambda s: (equip_res[s], s))
    fixed_list = lock_candidates[:fixed_num]
    if len(fixed_list) < fixed_num:
        # 该部位的有效副词条种类不够锁（比如主词条占掉了一种）
        return None, fixed_list

    # ③ 算分布，再把"超过当前"的概率加起来
    dist = math_cal.cal_p_with_fixed(re_useful, tmp_skill_list, fixed_list)
    return float(np.sum(dist[now_skill_num + 1:])), fixed_list


def update_character_equip_anylize(character_info_list):
    """
    对每个角色的 6 个部位各算一次概率分布，并算好"比当前更好"的概率。

    结果写回对象：
        ch.eq4_anylize = (头, 手, 衣服, 鞋子) 的分布
        ch.eq2_anylize = (球, 绳子) 的分布
        ch.p_better    = 六件各自"更好"的概率
    """
    skill_list = set()
    for ch in character_info_list:

        # 收集一下全账号用过哪些词条（仅用于最后打印日志）
        for i in ch.val_name_list:
            skill_list.add(i)
        skill_list.add(ch.main_yifu)
        skill_list.add(ch.main_xiezi)
        skill_list.add(ch.main_qiu)
        skill_list.add(ch.main_shengzi)

        # 外圈四件
        ret_tou = equip_anylize('小生命', ch.val_name_list, ch.val_num_tou, '头')
        ret_shou = equip_anylize('小攻击', ch.val_name_list, ch.val_num_shou, '手')
        ret_yifiu = equip_anylize(ch.main_yifu, ch.val_name_list, ch.val_num_yifu, '衣服')
        ret_xiezi = equip_anylize(ch.main_xiezi, ch.val_name_list, ch.val_num_xiezi, '鞋子')
        ch.eq4_anylize = (ret_tou, ret_shou, ret_yifiu, ret_xiezi)

        # 内圈两件
        ret_qiu = equip_anylize(ch.main_qiu, ch.val_name_list, ch.val_num_qiu, '球')
        ret_shengzi = equip_anylize(ch.main_shengzi, ch.val_name_list, ch.val_num_shengzi, '绳子')
        ch.eq2_anylize = (ret_qiu, ret_shengzi)

        # 六件"更好"的概率
        ch.p_get_better6()

    lsl = list(skill_list)
    lsl.sort()
    print('词条一览 ', lsl)
    print('需计算概率 ', repp_list)
    return


def get_last_zero(arr):
    """
    （旧版遗留工具函数，当前流程未使用）
    本意是找"最后一个概率为 0 的词条数"，但实现有问题：循环里第一轮就 return，实际恒返回 9。
    现已改用 get_eff_num_except_main() + 5 来算理论上限。
    """
    if arr[-1] != 0:
        return 9
    i = 9
    for i in range(len(arr) - 1, 0, -1):
        if i != 0:
            return i
    return i


# ==============================================================================
# 输出一：以角色为核心的排序
# ==============================================================================
def show_character_equipment_list(character_info_list):
    """
    以"角色"为视角输出两张表：

        sort_eq_list      角色-套装 维度：把该角色 4 个（或 2 个）部位的概率相加，
                          表示"刷这套装，这个角色任一部位变好的概率"
        sort_eq_once_list 角色-套装-部位 维度：单部位的概率 + 当前词条数 + 理论上限

    两者都按概率升序排（概率最小 = 最"没希望" = 刷了也难变好）。
    """
    eq_list = []
    eq_once_list = []
    for ch in character_info_list:
        eq4_better = sum(ch.p_better[:4])   # 外圈四件概率之和
        eq2_better = sum(ch.p_better[4:])   # 内圈两件概率之和

        for i in ch.eq4:                    # 外圈（2+2 时会有两套，各算一次）
            eq_list.append(('四件套-' + i, ch.name, eq4_better))
            # 下面四个是逐部位明细；理论上限 = 该角色有效词条种类数(去掉与主词条同名的) + 5 次强化
            eq_once_list.append(('四件套-' + i + '-头部', ch.name, ch.p_better[0], ch.val_num_tou, ch.eff_num_list[0] + 5))
            eq_once_list.append(('四件套-' + i + '-手部', ch.name, ch.p_better[1], ch.val_num_shou, ch.eff_num_list[1] + 5))
            eq_once_list.append(('四件套-' + i + '-衣服', ch.name, ch.p_better[2], ch.val_num_yifu, ch.eff_num_list[2] + 5))
            eq_once_list.append(('四件套-' + i + '-鞋子', ch.name, ch.p_better[3], ch.val_num_xiezi, ch.eff_num_list[3] + 5))
        eq_list.append(('两件套-' + ch.eq2, ch.name, eq2_better))

        eq_once_list.append(('两件套-' + ch.eq2 + '-球', ch.name, ch.p_better[4], ch.val_num_qiu, ch.eff_num_list[4] + 5))
        eq_once_list.append(('两件套-' + ch.eq2 + '-绳子', ch.name, ch.p_better[5], ch.val_num_shengzi, ch.eff_num_list[5] + 5))

    sort_eq_list = sorted(eq_list, key=lambda x: x[2])
    sort_eq_once_list = sorted(eq_once_list, key=lambda x: x[2])

    return sort_eq_list, sort_eq_once_list


# ==============================================================================
# 输出二：以套装 / 刷取地点为核心的排序（v1 旧口径，保留在输出里做对照）
# ==============================================================================
def load_character_to_equipment(character_info_list):
    """
    【旧口径】把"角色-套装"的关联翻过来变成"套装-角色"，
    同一套装下所有角色的提升概率**直接相加**。

    为什么是旧口径：如果两个角色有效词条完全一样，掉一件好遗器只能给其中一个人，
    直接相加会把概率算成两倍。所以 v2.0 改用 load_character_to_equipment_by_key()。
    这里保留下来，输出到「(旧)套装优化概率」「(旧)地点优化概率」两个页签做对照。
    """
    # 角色 → 套装 的概率登记
    for ch in character_info_list:
        eq4_better = sum(ch.p_better[:4])
        eq2_better = sum(ch.p_better[4:])
        equip_match_character_dict[ch.eq2].append((ch.name, eq2_better))
        for i in ch.eq4:
            equip_match_character_dict[i].append((ch.name, eq4_better))

    # 套装维度：把同一套装下所有角色的概率加起来
    equip_match_list = []
    for key in equip_match_character_dict:
        if len(equip_match_character_dict[key]) > 0:
            p_total = 0
            name_total = ''
            for name, p in equip_match_character_dict[key]:
                p_total += p
                name_total += name + '、'
            equip_match_list.append((key, p_total, name_total))
    sort_equip_match_list = sorted(equip_match_list, key=lambda x: x[1])

    # 地点维度：把同一地点下的各套装再相加
    pose_dict = dict()
    for pose, p, name in equip_match_list:
        pf = equip_from_dict[pose]
        if pf in pose_dict:
            pose_dict[pf].append((pose, p, name))
        else:
            pose_dict[pf] = [(pose, p, name)]

    pose_list = []
    for key in pose_dict:
        p_total = 0
        name_total = ''
        eq_name_total = ''
        for eq, p, name in pose_dict[key]:
            p_total += p
            name_total += name
            eq_name_total += eq + '+'
        pose_list.append((key, p_total, eq_name_total, name_total))

    sort_pose_list = sorted(pose_list, key=lambda x: x[1])

    return sort_equip_match_list, sort_pose_list


# ==============================================================================
# 输出三·补：带角色权重的口径（权重与概率完全分开）
# ==============================================================================
def normalize_character_weight(character_info_list):
    """
    把 Excel 里填的原始权重归一化到"平均 = 1"（等价于"所有权重之和 = 角色数"）。

    为什么要归一化：
        不归一化时，"加权得分"会随"你整体填得偏大还是偏小"而整体漂移。
        比如 55 个角色本来全填 1，你把所有人都改成 2，那么加权得分会整体翻倍，
        但这个翻倍不是"权重起作用了"，只是"你整体乘了 2"，等于白做。
        归一化之后，只有【相对大小】有意义：
            权重 2.0 的角色，贡献是权重 1.0 角色的 2 倍（而不是"总共翻倍"）。
        另一个好处是"加权得分 / 未加权得分"这个比值变得可读——
        它直接等于"这批角色相对全体平均的权重水平"。

    实现：w_i ← w_i × N / Σw
        N       角色总数
        Σw      所有原始权重之和
        全填 1.0 时 Σw = N，缩放系数 = 1，权重原地不动
        （所以"全填 1"时，加权结果必然等于未加权结果，这是一个天然的校验点）

    特殊处理：如果 Σw <= 0（用户把权重全填成 0 之类），全部退回 1.0，避免除零。
    """
    n = len(character_info_list)
    w_sum = sum(ch.weight_raw for ch in character_info_list)

    if w_sum <= 0:
        # 兜底：权重总数非正，视为"所有角色平权"
        for ch in character_info_list:
            ch.weight_raw = 1.0
        w_sum = float(n)

    scale = n / w_sum
    for ch in character_info_list:
        ch.weight = ch.weight_raw * scale

    return


def load_character_to_equipment_weighted(character_info_list):
    """
    【加权口径】把"旧口径（角色概率直接相加）"的结果乘以角色权重再重算一遍。

    与 load_character_to_equipment() 的关系：
        完全是同一个加法的结构，唯一区别是每一项 p 变成 p × w。
        概率 p 本身一个字节都没改 —— 权重只作为【求和时的系数】出现，
        这就是"权重与概率完全分开"。

    为什么要另写一个函数，而不是在原函数里改：
        ① 原有的「(旧)套装优化概率」「(旧)地点优化概率」两个页签必须保持
           数值不变（它们是没有权重的基准，要留着对照）；
        ② equip_match_character_dict 是【模块级全局变量】，
           原函数是往里 append 的，如果复用会重复累积，所以这里全用局部 dict。

    输出（三个列表，写成三张新页签）：
        w_eq_list   [套装名, 角色串, 未加权得分, 加权得分, 加权倍数, 涉及角色数]
        w_pos_list  [刷取地点, 套装串, 未加权得分, 加权得分, 加权倍数, 角色串]
        w_char_list [角色名, 原始权重, 归一化权重, 当前有效词条合计,
                     理论最高合计, 当前完成度%, 一般最高合计, 一般完成度%]
                    —— 这一页不参与加权，纯粹是"自检查 + 进度参照"用的明细表

    关于"得分"：
        单次出货概率本身就是小数点后好几位（比如 0.0035），再乘权重更难看。
        所以这一层统一 ×100 展示，并且明确叫"得分"而不是"概率"——
        它已经不是概率了（加权和可能超过 1），只是用来【排序/比大小】的分数。
    """
    # ① 角色 → 套装：登记 (角色名, 该角色对该套装的提升概率, 该角色权重)
    #    注意这里和旧口径完全同构：外圈用 p_better[:4] 之和，内圈用 p_better[4:] 之和
    eq_map = dict()
    for ch in character_info_list:
        eq4_better = float(sum(ch.p_better[:4]))     # 外圈四件的概率之和
        eq2_better = float(sum(ch.p_better[4:]))     # 内圈两件的概率之和
        eq_map.setdefault(ch.eq2, []).append((ch.name, eq2_better, ch.weight))
        for eq_name in ch.eq4:                       # 2+2 时会在两个套装下各记一次（与旧口径一致）
            eq_map.setdefault(eq_name, []).append((ch.name, eq4_better, ch.weight))

    # ② 套装层级：未加权 / 加权 两套求和，顺便记录涉及哪些角色
    #    这里同时算两份和，是为了能在同一行里并排展示、直接相除：
    #       p_plain  = 旧口径（每个角色算一份），它必须严格等于「(旧)套装优化概率」页的值
    #       p_weight = 加权口径（每个角色按自己的归一化权重算份额）
    w_eq_list = []
    for eq_name in eq_map:
        items = eq_map[eq_name]
        p_plain = 0.0        # 未加权：等价于"每个角色权重都是 1"
        p_weight = 0.0       # 加权：每个角色按自己的权重贡献
        name_set = []
        for name, p, w in items:
            p_plain += p
            p_weight += p * w
            if name not in name_set:       # 手动去重：2+2 角色在同一套装里不会重复，
                name_set.append(name)      # 但保留顺序比用 set() 更好读（set 的顺序是随机的）

        # 加权倍数 = 加权得分 / 未加权得分
        #   等于"这批角色相对全体平均的权重水平"：1.0 = 平均水平，
        #   1.606 = 全是权重 2.0 的主力，0.402 = 全是权重 0.5 的仓管。
        #   p_plain 可能为 0（理论上不会，但防御一下，避免除零）
        ratio = (p_weight / p_plain) if p_plain > 0 else 0.0

        w_eq_list.append([eq_name,
                          '、'.join(name_set) + '、',   # 末尾补顿号，和旧页签的写法保持一致
                          round(p_plain * 100, 4),     # ×100 是为了好读；留 4 位小数够排序用了
                          round(p_weight * 100, 4),
                          round(ratio, 4),
                          len(items)])

    # ③ 地点层级：一个本里的几个套装再加起来
    #    pos_map 的值是一个 4 元结构：['本内套装名串', '角色名列表', 未加权和, 加权和]
    #    用 list 而不是 dict 只是为了少写几行，下标含义见上面这行注释。
    #    注意：这里不能再乘一次权重 —— 上一步的 s_weight 已经带权重了，
    #    直接相加即可，否则会变成"权重乘两次"。
    pos_map = dict()
    for eq_name, name_str, s_plain, s_weight, ratio, num in w_eq_list:
        pos = equip_from_dict[eq_name]                  # 套装全称 → 刷取地点
        if pos not in pos_map:
            pos_map[pos] = ['', [], 0.0, 0.0]
        pos_map[pos][0] += eq_name + '、'
        # 角色名去重：同一个本里的两个套装可能被同一个角色同时用到（比如 2+2 跨本），
        # 直接字符串拼接会出现"张三、张三、"，所以逐个查重后再拼
        for n in name_str.split('、'):
            if n and n not in pos_map[pos][1]:
                pos_map[pos][1].append(n)
        pos_map[pos][2] += s_plain
        pos_map[pos][3] += s_weight

    w_pos_list = []
    for pos in pos_map:
        eq_str, names, s_plain, s_weight = pos_map[pos]
        ratio = (s_weight / s_plain) if s_plain > 0 else 0.0   # 地点层级的加权倍数，口径同 ②
        w_pos_list.append([pos,
                           eq_str,
                           round(s_plain, 4),
                           round(s_weight, 4),
                           round(ratio, 4),
                           '、'.join(names) + '、'])

    # ④ 角色权重明细（自检查页：能一眼看出 Excel 里的值有没有被读进来、归一化后是多少）
    w_char_list = []
    # ⑤ 角色 × 部位明细（每个角色 6 行，共 330 行）：看"哪个角色的哪个部位最空"
    w_slot_list = []

    # 部位顺序固定为 头/手/衣服/鞋子/球/绳子，和 val_num_list、eff6 完全对齐
    pos_names = ['头', '手', '衣服', '鞋子', '球', '绳子']

    for ch in character_info_list:
        # ---- 先把"六个部位"的上限取出来 ----
        # 坑：cal_character_equip_update_max_num() 是在 `for i in ch.eq4` 循环里 append 的，
        # 所以 2+2 外圈的角色（eq4 有两个元素）eff_num_list 会有 10 个元素，
        # 外圈 4 个部位被记了两遍。直接 sum() 会把上限算高（比如 63 而不是 45）。
        # 这里只取"真正的六个部位"：前 4 个（外圈头/手/衣服/鞋子）+ 最后 2 个（球/绳子）。
        if hasattr(ch, 'eff_num_list') and len(ch.eff_num_list) >= 6:
            eff6 = list(ch.eff_num_list[:4]) + list(ch.eff_num_list[-2:])
        else:
            eff6 = [0, 0, 0, 0, 0, 0]

        # 理论最高 = 六件遗器各自"副词条里最多能有几条有效"之和 + 30（六件 × 强化 5 次）
        max_eff = sum(eff6) + 30 if eff6 else 0
        done = (ch.total_val / max_eff) if max_eff > 0 else 0.0

        # ---- 「一般最高」：一个"良好遗器"的参照线，而不是满配理论上限 ----
        # 口径：每件遗器按「初始 4 条副词条里默认命中 2 条有效」+「5 次强化里允许不中 1 次，
        #      即命中 4 次」来算 → 每件 2 + 4 = 6 条。
        #
        # 但"初始那 2 条"要受该部位有效词条种类数的限制。分两种情况：
        #
        #   ① 该部位有效副词条上限 >= 2 → 初始 2 条 + 强化命中 4 次（允许不中 1 次）= 6 条
        #
        #   ② 该部位有效副词条上限 = 1 → 初始最多只能有 1 条有效，
        #      这时候不该还要求"6 条里拿 5 条"（那已经是 83% 完美度的神级件了），
        #      所以强化部分放宽到"允许不中 2 次"，即命中 3 次 → 1 + 3 = 4 条
        #      （相当于 6 条上限里拿到 4 条，和 ① 的 9 条上限里拿 6 条一样，都是约 67% 完美度）
        #
        # ② 的典型例子（有效词条只有 2 种，其中一个还被主词条占了）：
        #      大丽花 有效=击破/速度，鞋子主词条=速度 → 副词条只剩击破有效，上限 1
        #      大丽花 绳子主词条=击破       → 副词条只剩速度有效，上限 1
        #      盾丹 有效=大攻击/速度，衣服/鞋子/球/绳子主词条都是大攻击 → 上限 1（4 个部位）
        #      真珠 有效=大防御/速度，鞋子主词条=速度、球主词条=大防御 → 各上限 1
        #
        # 六件合计 = Σ 每件的"一般条数"。这个值一定不超过"理论最高"
        # （理论最高用的是 min(4, eff_i) + 5，两种情况都比它大）。
        common_max = 0
        for e in eff6:
            if e >= 2:
                common_max += 6      # 初始 2 + 强化命中 4
            elif e == 1:
                common_max += 4      # 初始 1 + 强化命中 3
            # e == 0 时该部位不可能有有效副词条，计 0
        common_done = (ch.total_val / common_max) if common_max > 0 else 0.0

        w_char_list.append([ch.name,
                            round(ch.weight_raw, 4),
                            round(ch.weight, 4),
                            ch.total_val,
                            max_eff,
                            round(done * 100, 2),
                            common_max,
                            round(common_done * 100, 2)])

        # ---- ⑤ 逐部位拆开，一行一个部位 ----
        # 这一页的目的：角色层级的完成度是六件平均，看不出"是哪一件在拖后腿"。
        # 拆到部位之后，一眼就能定位"权重 2.0 的角色，鞋子只有 3 条"这种最该砸体力的地方。
        #
        # 后面做 cost 定向优化时也直接吃这张表：
        #   标的 = (套装, 部位, 主词条)，基线 c = 「当前有效词条数」，收益还要看
        #   「理论上限 - 当前」还剩多少空间，这张表的最后两列就是给那一步准备的。
        for i in range(6):
            pos = pos_names[i]
            eff_slot = eff6[i]                 # 该部位"有效副词条"最多几条
            now = ch.val_num_list[i]           # 该部位当前有效词条数

            # 主词条：头/手是固定的，其余来自 Excel
            if pos == '头':
                main_skill = '小生命'
            elif pos == '手':
                main_skill = '小攻击'
            elif pos == '衣服':
                main_skill = ch.main_yifu
            elif pos == '鞋子':
                main_skill = ch.main_xiezi
            elif pos == '球':
                main_skill = ch.main_qiu
            else:
                main_skill = ch.main_shengzi

            # 套装：外圈四件列出该角色穿的所有外圈套装（2+2 用「+」连起来），内圈两件是 eq2
            if i < 4:
                eq_str = '+'.join(ch.eq4)
            else:
                eq_str = ch.eq2

            max_slot = eff_slot + 5 if eff_slot > 0 else 0        # 理论上限（强化 5 次全命中）
            if eff_slot >= 2:
                common_slot = 6                                   # 一般上限：初始 2 + 命中 4
            elif eff_slot == 1:
                common_slot = 4                                   # 一般上限：初始 1 + 命中 3
            else:
                common_slot = 0

            max_rate = (now / max_slot) * 100 if max_slot > 0 else 0.0
            common_rate = (now / common_slot) * 100 if common_slot > 0 else 0.0

            # 剩余可提升条数：正数 = 还有空间，0 = 已经顶到上限
            left_max = max(max_slot - now, 0)
            left_common = max(common_slot - now, 0)

            # ---- 定向三档的提升概率（1 cost / 2 cost / 5 cost）----
            # 口径：已是定向件的条件下，强化完"有效词条数 > 当前"的概率，
            # 不含主词条 / 部位 / 出金系数（那些是"刷本掉落"的概率，定向不需要）
            p1, _ = cal_slot_directed_rate(ch.val_name_list, main_skill, now, 0)
            p2, lock2 = cal_slot_directed_rate(ch.val_name_list, main_skill, now, 1)
            p5, lock5 = cal_slot_directed_rate(ch.val_name_list, main_skill, now, 2)

            w_slot_list.append([ch.name,
                                round(ch.weight, 4),
                                pos,
                                eq_str,
                                main_skill,
                                now,
                                max_slot,
                                round(max_rate, 2),
                                common_slot,
                                round(common_rate, 2),
                                left_max,
                                left_common,
                                round(p1 * 100, 2) if p1 is not None else '',
                                round(p2 * 100, 2) if p2 is not None else '',
                                '、'.join(lock2) if lock2 else '',
                                round(p5 * 100, 2) if p5 is not None else '',
                                '、'.join(lock5) if lock5 else ''])

    # 排序：得分升序（和旧页签保持同样的阅读习惯：越靠后越值得刷）
    w_eq_list = sorted(w_eq_list, key=lambda x: x[3])
    w_pos_list = sorted(w_pos_list, key=lambda x: x[3])
    w_char_list = sorted(w_char_list, key=lambda x: -x[2])   # 角色页按权重从高到低
    # 部位页：先按权重高→低，同一角色内按"一般完成度"低→高
    # 这样排在最上面的是"权重高、又最空"的部位，也就是最该砸体力的地方
    w_slot_list = sorted(w_slot_list, key=lambda x: (-x[1], x[9]))

    return w_eq_list, w_pos_list, w_char_list, w_slot_list


# ==============================================================================
# 输出三：以套装/地点为核心"去重后"的准确概率（v2.0 主口径）
# ==============================================================================
def cal_combine_eq_p(key, type_list):
    """
    计算【一个套装 + 一个部位 + 一个主词条】这件遗器出货后，
    "至少对其中一个使用者有提升"的准确概率（不重复计数）。

    key       = (套装名, 部位名, 主词条名)
    type_list = [(角色名, 该角色有效词条列表, 该角色当前该部位词条数, 该角色单部位提升概率), ...]

    思路：
        多个角色共用同一个"套装+部位+主词条"时，两个角色的有效词条集合可能不同，
        某一组具体副词条组合可能"对A有提升但对B没用"。所以：
          1) 枚举所有可能的 4 条副词条组合（按"有效/无效"分类），
             每组组合在这个部位出现的概率 skill_list_combine_p；
          2) 对这组组合，分别看"对每个角色来说是否提升"（查各自概率分布的尾部）；
          3) 同一组组合只取"最有利的那个角色"的概率（max_p），再按组合出现问题，
             这样同一件遗器不会被重复计数；
          4) 最后乘上主词条、部位、出金 1.05。
        只有一个使用者时，直接返回该角色已算好的概率 p。

    ⚠ 这里沿用 math_cal.cal_p_at_need_combine 的组合枚举，
      该函数从 useful_num=1 开始枚举，漏掉了"4 条副词条全部无效"的情况，
      对"当前 0 条有效词条"的部位会略微低估。
    """
    if len(type_list) == 1:
        # 只有一个人用：直接用现成结果
        name, skill_list, val_num, p = type_list[0]
        return p

    eq_name, eq_pose, main_skill = key

    # —— 与 equip_anylize() 同款的三层系数 ——
    head_p = 1.05                       # 2.1 件/次 × 0.5 套装
    if eq_pose in ['头', '手', '衣服', '鞋子']:
        head_p *= 0.25                  # 外圈 4 个部位
    else:
        head_p *= 0.5                   # 内圈 2 个部位
    head_p *= equip_main[(eq_pose, main_skill)]   # 主词条概率

    # 预先算好"初始有 k 条有效词条时，练满后最终有 m 条"的概率（k=0..4）
    val_num_dict = dict()
    for i in range(5):
        val_num_dict[i] = math_cal.cal_p_when_init_val_num(i)

    # 按"具体副词条组合"归集：同一个组合下，可能挂多个角色的数据
    skill_list_dict = dict()
    for name, skill_list, val_num, p in type_list:
        # ⚠ 这里必须复制一份再用，不能原地改！
        #   type_list 里装的是各角色对象上的 ch.val_name_list **本体**
        #   （同一个 list 对象被 6 个部位、以及 2+2 的两套外圈共享）。
        #   原代码直接 skill_list.remove(main_skill)，于是每算一个部位就把这个角色的
        #   有效词条永久删掉一条：大丽花的「速度、击破」先被鞋子删掉速度、
        #   再被绳子删掉击破，算到后面就变成空列表，后面所有依赖它的计算全错。
        skill_list = [s for s in skill_list if s != main_skill]

        # 每个角色都要先剔除"主词条同名"的词条，再算出无效候选池
        bk_skill_list = copy.deepcopy(res_skill_list)
        if main_skill in bk_skill_list:
            bk_skill_list.remove(main_skill)
        for skill in skill_list:
            if skill in bk_skill_list:
                bk_skill_list.remove(skill)

        # 枚举"4 条副词条里 u 条有效"的所有组合及其概率
        useful_skill_combine_p_list = math_cal.cal_p_at_need_combine(skill_list, bk_skill_list)

        sum_p = 0
        for useful_num, skill_list_combine, weight_list_combine, skill_list_combine_p in useful_skill_combine_p_list:
            # 这个组合"最终练满后超过该角色当前词条数"的概率
            sum_p += sum(val_num_dict[useful_num][val_num:]) * skill_list_combine_p
            if skill_list_combine not in skill_list_dict:
                skill_list_dict[skill_list_combine] = []
            skill_list_dict[skill_list_combine].append((name, useful_num, val_num, skill_list_combine_p, sum(val_num_dict[useful_num][val_num:])))

    # 同一组副词条组合，只计入"最有利的那个角色"的概率，避免重复计数
    better_p = 0
    for key2 in skill_list_dict:
        max_p = 0
        eq_p = 0
        for name, useful_num, val_num, eq_p, b_p in skill_list_dict[key2]:
            if b_p > max_p:
                max_p = b_p
        better_p += max_p * eq_p

    return better_p * head_p


def load_character_to_equipment_by_key(character_info_list):
    """
    【v2.0 主口径】以 (套装, 部位, 主词条) 为 key 汇总，用 cal_combine_eq_p 算去重后的概率，
    再分别汇总到"套装"和"刷取地点"两个层级。

    返回：
        opt_eq_pose_list  [(套装, 部位, 角色集合, 概率)]           按概率升序
        eq_list           [(套装, 角色集合, 套装概率)]              各部位概率相加
        pos_list          [(刷取地点, 地点概率, 套装串, 角色串)]     本内各套装概率相加
    """
    equip_match_character_key_dict = dict()

    # ① 把每个角色的 6 个部位拆成 6 个 key，挂到字典里
    for ch in character_info_list:

        # 保证副词条列表的顺序一致，否则"内容相同但顺序不同"会被当成不同的组合
        ch.val_name_list.sort()

        # 内圈两件
        dict_key_list = [[(ch.eq2, '球', ch.main_qiu), (ch.name, ch.val_name_list, ch.val_num_qiu, ch.p_better[4])],
                         [(ch.eq2, '绳子', ch.main_shengzi), (ch.name, ch.val_name_list, ch.val_num_shengzi, ch.p_better[5])]]
        # 外圈四件
        for eq_name in ch.eq4:
            dict_key_list.append([(eq_name, '头', '小生命'), (ch.name, ch.val_name_list, ch.val_num_tou, ch.p_better[0])])
            dict_key_list.append([(eq_name, '手', '小攻击'), (ch.name, ch.val_name_list, ch.val_num_shou, ch.p_better[1])])
            dict_key_list.append([(eq_name, '衣服', ch.main_yifu), (ch.name, ch.val_name_list, ch.val_num_yifu, ch.p_better[2])])
            dict_key_list.append([(eq_name, '鞋子', ch.main_xiezi), (ch.name, ch.val_name_list, ch.val_num_xiezi, ch.p_better[3])])

        for key_name, val in dict_key_list:
            if key_name not in equip_match_character_key_dict:
                equip_match_character_key_dict[key_name] = []
            equip_match_character_key_dict[key_name].append(val)

    # ② 逐 key 计算准确概率
    eq_match_p_dict = dict()
    for key in equip_match_character_key_dict:
        p = cal_combine_eq_p(key, equip_match_character_key_dict[key])
        eq_match_p_dict[key] = p

    # ③ 整理成 (概率, 使用人数, key, 明细) 并排序
    eq_pose_list = []
    for key in equip_match_character_key_dict:
        eq_pose_list.append((eq_match_p_dict[key], len(equip_match_character_key_dict[key]), key, equip_match_character_key_dict[key]))
    eq_pose_list = sorted(eq_pose_list, key=lambda x: x[0])

    # ④ 压成 (套装, 部位, 角色集合, 概率)
    opt_eq_pose_list = []
    for bp, useful_character_num, eq_pose, ch_list in eq_pose_list:
        character_str = set()
        for ch_name, _, _, _ in ch_list:
            character_str.add(ch_name)
        opt_eq_pose_list.append([eq_pose[0], eq_pose[1], character_str, bp])
    opt_eq_pose_list = sorted(opt_eq_pose_list, key=lambda x: x[3])

    # ⑤ 汇总到套装层级：同一套装各部位概率相加，角色集合取并集
    eq_list = []
    eq_dict = dict()
    for eq_name, eq_pose, ch_str, p in opt_eq_pose_list:
        if eq_name not in eq_dict:
            eq_dict[eq_name] = [ch_str, p]
        else:
            eq_dict[eq_name][0] = eq_dict[eq_name][0] | ch_str
            eq_dict[eq_name][1] += p
    for key in eq_dict:
        eq_list.append([key, eq_dict[key][0], eq_dict[key][1]])
    eq_list = sorted(eq_list, key=lambda x: x[2])

    # ⑥ 汇总到刷取地点层级：一个本里的两个套装概率相加
    pos_dict = dict()
    for eq_name, ch_list, p in eq_list:
        position_name = equip_from_dict[eq_name]
        if position_name not in pos_dict:
            pos_dict[position_name] = ['', set(), 0]
        pos_dict[position_name][0] += eq_name + '、'      # 本内套装名串
        pos_dict[position_name][1] = pos_dict[position_name][1] | ch_list
        pos_dict[position_name][2] += p
    pos_list = []
    for key in pos_dict:
        pos_list.append([key, pos_dict[key][2], pos_dict[key][0], pos_dict[key][1]])
    pos_list = sorted(pos_list, key=lambda x: x[1])

    # ⑦ 把集合类型的角色名转成字符串，便于写进 Excel
    for i in range(len(eq_list)):
        ch_str = ''
        for j in eq_list[i][1]:
            ch_str += j + '、'
        eq_list[i][1] = ch_str

    for i in range(len(opt_eq_pose_list)):
        ch_str = ''
        for j in opt_eq_pose_list[i][2]:
            ch_str += j + '、'
        opt_eq_pose_list[i][2] = ch_str

    for i in range(len(pos_list)):
        ch_str = ''
        for j in pos_list[i][3]:
            ch_str += j + '、'
        pos_list[i][3] = ch_str

    return opt_eq_pose_list, eq_list, pos_list


# ==============================================================================
# 输出五：遗器分解建议（v2.9，一行 = 一个「套装 × 部位」）
# ==============================================================================
def read_dismantle_thresholds(excel_file_path):
    """
    从生成表的「分解阈值」页读三个口径参数（这一页由 01 从配置表原样带过来，
    所以改阈值只改 config\项目配置-豆包.xlsx，不用碰代码）。

    页结构：项目 / 数值 / 说明，按关键词匹配而不是按行号，
    这样以后在中间插一行、调换顺序都不会读错。

    返回 (低阈值, 高阈值, 是否严格超过)，读不到时用 10 / 40 / 1。
    """
    low, high, strict = 10.0, 40.0, 1
    try:
        df = pd.read_excel(excel_file_path, sheet_name='分解阈值')
    except Exception:
        print("!! 生成表里没有「分解阈值」页：按默认 低10% / 高40% / 严格超过 计算")
        return low, high, strict

    for _, r in df.iterrows():
        name = str(r.iloc[0]) if len(r) > 0 else ''
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

    # 防御：阈值填反了会让结论全错，这里直接纠正并提示
    if low > high:
        print("!! 「分解阈值」页里低阈值(%s) 大于高阈值(%s)，已按 10/40 处理，请去配置表改成低<高"
              % (low, high))
        low, high = 10.0, 40.0
    return low, high, strict


def _binom_at_least(n, p, t):
    """
    P(X >= t)，X ~ Bin(n, p)。

    为什么手写而不上 scipy：n 只可能是 4 或 5，直接枚举 5 项比装一个依赖划算，
    而且能和项目里其它"纯 python 算概率"的部分保持一致的风格。
    """
    if t <= 0:
        return 1.0
    if t > n:
        return 0.0
    return float(sum(math.comb(n, k) * (p ** k) * ((1.0 - p) ** (n - k))
                     for k in range(t, n + 1)))


def build_dismantle_sheet(character_info_list, low=10.0, high=40.0, strict=1):
    """
    「分解建议」：回答"刚掉出来的这件，初始 4 条里至少几条有效才值得强化"。

    为什么是「套装 × 部位」而不是「角色 × 套装 × 部位」：
        一件遗器只属于一个套装、一个部位。留不留只看"它能不能帮到这套这个部位里的某个人"，
        所以按套装聚合，门槛取所有使用者里最宽松的那个（对任何一人有用就值得留），
        其余使用者列在最后一列。

    前提口径（用户 2026-11-20 定）：
        · s = 初始 4 条副词条里有效词条的条数。初始 3 条的遗器第 4 条也看得见，
          所以 s 同样按 4 条数（不和 3 条区分）。
        · 强化次数按遗器本体分两档：
            4 条基础（+0 就 4 条）→ 5 次强化，满级 9 投
            3 条基础（第 4 条是 +3 时补的）→ 4 次强化，满级 8 投
        · 每次强化落在哪一条是等概率的 → 命中有效词条的概率 = s / 4
        · 满级有效条数 = s + X，X ~ Bin(强化次数, s/4)

    达标定义：满级后【严格超过】现役最好那条（T+1 条）；
        strict=0 时打平（T 条）也算达标（阈值页里可改）。

    v2.9.1 追加「保底」口径（用户 2026-11-20 补）：
        上面的门槛说的是"运气正常能不能超过现役【最好】的那件"，
        但一件遗器值不值得留，还该看"就算歪 1~2 条，是不是也比现役【最差】的那件强"——
        连这个都做不到，就完全没有保留必要（直接分解）。
        保底假设：最坏合理情形 = 歪 2 条（命中 n-2 次），
        保底可达 = s + (n-2) 条，要 ≥ 现役最低 T_min（strict 时 T_min+1）。
        → 保底门槛 s ≥ T_min(+strict) - (n-2)

    返回的行：套装 / 部位 / 使用者数 / 现役最高 / 现役最低 / 初始可中上限 /
             升级门槛(4条·3条) / 保底门槛(4条·3条) / s=2、s=3 的达标率（各两档）/
             升级结论 / 保底结论 / 能帮到的角色
    """
    pos_names = ['头', '手', '衣服', '鞋子', '球', '绳子']

    # ---- ① 先把「套装 × 部位」聚起来：这部分等价于把「角色部位完成度」按套装分组 ----
    # 注意 2+2 外圈：一个角色的 eq4 有两个套装，同一件外圈遗器只属于其中一个，
    # 所以这里要拆开各记一次（而不是把 "A+B" 当成一个套装名）。
    agg = dict()   # (套装, 部位) -> {'now': [...], 'eff': [...], 'users': [...]}
    for ch in character_info_list:
        # eff_num_list 在 2+2 角色上会有 10 个元素（外圈四部位被记了两遍），
        # 与「角色部位完成度」同一取法：前 4 个 + 最后 2 个 = 真正的六个部位
        if hasattr(ch, 'eff_num_list') and len(ch.eff_num_list) >= 6:
            eff6 = list(ch.eff_num_list[:4]) + list(ch.eff_num_list[-2:])
        else:
            eff6 = [0, 0, 0, 0, 0, 0]

        for i in range(6):
            if i < 4:
                eqs = [str(e).strip() for e in ch.eq4 if str(e).strip()]
            else:
                eqs = [str(ch.eq2).strip()] if str(ch.eq2).strip() else []

            for eq in eqs:
                k = (eq, pos_names[i])
                if k not in agg:
                    agg[k] = {'now': [], 'eff': [], 'users': []}
                agg[k]['now'].append(int(ch.val_num_list[i]))
                agg[k]['eff'].append(int(eff6[i]))
                if ch.name not in agg[k]['users']:
                    agg[k]['users'].append(ch.name)

    # ---- ② 逐行算门槛与达标率 ----
    rows = []
    for (eq, pos) in sorted(agg, key=lambda k: (k[0], pos_names.index(k[1]))):
        d = agg[(eq, pos)]
        T = max(d['now'])                       # 现役最高词条数（这套这部位的天花板）
        T_min = min(d['now'])                   # 现役最低词条数（这套这部位最弱的那件）
        cap = max(min(4, e) for e in d['eff'])  # 初始 4 条里最多能有几条有效（受该部位有效种类数限制）
        target = T + (1 if strict else 0)       # 升级达标线（"超过现役最好"）
        target_min = T_min + (1 if strict else 0)  # 保底达标线（"超过现役最差"）

        # 升级门槛：满级最多 s + n 条，要 ≥ target → s ≥ target - n
        def gate_cell(n):
            k = target - n
            if k <= 0:
                return 0                      # 现役很低，初始 0 条有效也够（理论上）
            return k if k <= cap else '不可达'

        # 保底门槛：假设歪 2 条 → 只命中 n-2 次 → s + (n-2) ≥ target_min
        def gate_floor_cell(n):
            k = target_min - (n - 2)
            if k <= 0:
                return 0
            return k if k <= cap else '不可达'

        # 达标率：初始正好 s 条有效时，满级达到 target 的概率
        def rate_cell(n, s):
            if s > cap:
                return '—'                    # 该部位有效种类不够，s 取不到这个值
            return round(_binom_at_least(n, s / 4.0, target - s) * 100, 2)

        # 升级结论：在 s = 0..cap 里找第一个"概率 ≥ 低阈值"的，作为最小可赌条数
        def verdict(n_label, n):
            if target > cap + n:
                return "%s：直接分解（最多 %d 条，超不过现役 %d 条）" % (n_label, cap + n, T)
            for s in range(0, cap + 1):
                p = _binom_at_least(n, s / 4.0, target - s)
                if p * 100 >= low:
                    tag = '可赌' if p * 100 >= high else '谨慎'
                    return "%s：初始≥%d 条（达标 %.1f%%，%s）" % (n_label, s, p * 100, tag)
            return "%s：直接分解（各档概率都不到 %.0f%%）" % (n_label, low)

        # 保底结论：哪怕歪 2 条，只要最终条数超过现役最低那件，就"有保留价值"；
        #   连这个都做不到 → 没必要留（用户要的正是这一条）
        def verdict_floor(n_label, n):
            k = target_min - (n - 2)
            if k <= 0:
                return "%s：有保留价值（歪2条也够打到现役最低线）" % n_label
            if k > cap:
                return "%s：无保留必要（歪2条最多 %d 条，超不过现役最低 %d 条）" % (n_label, cap + n - 2, T_min)
            return "%s：初始≥%d 条才有保底价值（歪2条仍能超过 %d 条）" % (n_label, k, T_min)

        rows.append([eq,
                     pos,
                     len(d['users']),
                     T,
                     T_min,
                     cap,
                     gate_cell(5),          # 升级门槛：4 条基础，5 次强化
                     gate_cell(4),          # 升级门槛：3 条基础，4 次强化
                     gate_floor_cell(5),    # 保底门槛：4 条基础（歪 2 条 → 命中 3 次）
                     gate_floor_cell(4),    # 保底门槛：3 条基础（歪 2 条 → 命中 2 次）
                     rate_cell(5, 2),
                     rate_cell(4, 2),
                     rate_cell(5, 3),
                     rate_cell(4, 3),
                     verdict('4条基础', 5) + '；' + verdict('3条基础', 4),
                     verdict_floor('4条基础', 5) + '；' + verdict_floor('3条基础', 4),
                     '、'.join(d['users']) + '、'])

    return rows


# ==============================================================================
# 期望次数换算
# ==============================================================================
def cal_exp_numb_from_rate(p, rate):
    """
    已知单次出货概率 p，问"刷多少次，累计出货概率达到 rate"。

    数学上就是 n = ln(1-rate) / ln(1-p)，这里用 while 硬循环实现（作者原版写法）：
        每次不出货的概率 rp = 1-p；连续 n 次都不出货的概率 = rp^n；
        要求 rp^n <= 1-rate，即累计出货概率 >= rate。
    次数上限 9999（超过就返回 9999，表示"基本没戏"）。

    ⚠ 改用对数公式可以直接算，且不受 9999 截断影响。
    """
    rp = 1 - p                 # 一次刷取不出货的概率
    target = 1 - rate          # 最终"不出货"的目标阈值

    num = 1
    res = rp
    while res > target:
        res *= rp
        num += 1
        if num > 9999:
            return num
    return num


# ==============================================================================
# 写 Excel
# ==============================================================================
def get_save_cache_dict(data_list, name_list, sheet_name='default', add_exp_num=0, p_ind=0, star_map=None):
    """
    把"列表的列表"转成"列：值的列表"的字典（pandas 写 Excel 需要这种形状），
    并可选地追加一列"XX%期望出货次数"。

    data_list  每个元素是一行 tuple/list
    name_list  每列的表头
    add_exp_num 期望概率（如 0.95），False 表示不加这一列
    p_ind      概率在第几列（用于算期望次数），指的是**行内下标**（不含星级列）
    star_map   {角色名: 星级}。给了它，且 name_list 里写了「角色星级」，
               这一列就按同一行的角色名去查表填充，不占 data_list 的位置。
               （这样加列不用改任何一个产出 list 的函数，行结构保持原样）

    用法举例：name_list 里把「角色星级」插在「角色名称」后面即可，位置随意，
    因为它是"查出来的"，不参与行内下标计算。
    """
    # 「角色星级」是额外查出来的列，行数据里没有它
    real_cols = [c for c in name_list if c != '角色星级']
    idx_of = {c: i for i, c in enumerate(real_cols)}
    has_star = ('角色星级' in name_list) and (star_map is not None)

    data_dict = dict()
    for i in range(len(real_cols)):
        col_name = real_cols[i]
        data_dict[col_name] = []
        for j in range(len(data_list)):
            data_dict[col_name].append(data_list[j][i])

    if has_star:
        nm_ind = idx_of.get('角色名称', 0)
        data_dict['角色星级'] = [star_map.get(str(data_list[j][nm_ind]).strip(), "")
                             for j in range(len(data_list))]

    if add_exp_num != False:
        col_name = str(int(add_exp_num * 100)) + '%期望出货次数'
        data_dict[col_name] = []
        for j in range(len(data_list)):
            data_dict[col_name].append(cal_exp_numb_from_rate(data_list[j][p_ind], add_exp_num))

    # 按 name_list 的顺序重排 key，保证「角色星级」落在你写的位置上；
    # 再把不在 name_list 里的列（「XX%期望出货次数」）按原顺序补到末尾，
    # 否则会被这个过滤顺手丢掉（曾经丢过一次，整整少了一列）。
    ordered = {c: data_dict[c] for c in name_list if c in data_dict}
    for c in data_dict:
        if c not in ordered:
            ordered[c] = data_dict[c]
    return (ordered, sheet_name)


def save_excel(data_dict_list, save_path):
    """
    一次性写出多张 sheet（覆盖写）。
    data_dict_list 的元素是 (列字典, sheet 名)。
    """
    with pd.ExcelWriter(save_path, engine='openpyxl', mode='w') as writer:
        for data_dict, sheet_name in data_dict_list:
            df = pd.DataFrame(data_dict)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return


def get_eff_num_except_main(use_list, main_skill):
    """
    该部位"理论上最多能有多少条【副词条中的】有效词条"。

    规则：副词条最多 4 条；若"有效词条"里含有与主词条同名的那条，
    它会被主词条占掉（不会出现在副词条里），所以要先减 1。
    最终成品有效词条数上限 = 这个值 + 5（+15 强化最多加 5 次，每次都可能命中有效条）。
    """
    if main_skill in use_list:
        return max(len(use_list) - 1, 0)
    return min(4, len(use_list))


def cal_character_equip_update_max_num(character_info_list):
    """
    给每个角色、每个部位算"理论上限"，存进 ch.eff_num_list，
    供输出表的"理论最高有效词条数量"列使用（顺序：头、手、衣服、鞋子、球、绳子）。
    """
    for ch in character_info_list:
        eff_num_list = []
        for i in ch.eq4:
            eff_num_list.append(get_eff_num_except_main(ch.val_name_list, '小生命'))   # 头
            eff_num_list.append(get_eff_num_except_main(ch.val_name_list, '小攻击'))   # 手
            eff_num_list.append(get_eff_num_except_main(ch.val_name_list, ch.main_yifu))    # 衣服
            eff_num_list.append(get_eff_num_except_main(ch.val_name_list, ch.main_xiezi))   # 鞋子

        eff_num_list.append(get_eff_num_except_main(ch.val_name_list, ch.main_qiu))     # 球
        eff_num_list.append(get_eff_num_except_main(ch.val_name_list, ch.main_shengzi)) # 绳子
        ch.load_eff_num(eff_num_list)

    return


# ==============================================================================
# 主流程
# ==============================================================================
def main(show_log=False, exp_rate=0.95):
    """
    全流程：
        1. 初始化主词条/副词条字典
        2. 读 Excel（遗器映射表 + 角色表）
        3. 算各部位理论上限
        4. 算各部位概率分布
        5. 生成四类结论并写入 分析结果.xlsx

    输入表怎么选：
        默认读 01-一键导出.bat 生成的「星铁装备刷取-自动生成-豆包.xlsx」
        （它自带 角色遗器 / 遗器相关 / 主词条分布 三页，可直接喂本脚本）。
        老手填表「星铁装备刷取.xlsx」已于 2026-11-19 退役、移入 归档-豆包\，不再是默认输入。
        也可以在命令行直接指定别的表：

            python main.py 星铁装备刷取-自动生成-豆包.xlsx
            python main.py 输入表.xlsx 输出表.xlsx
    """
    # 命令行参数：第 1 个 = 输入表，第 2 个 = 输出表；不传就用默认那两个文件名。
    # 本脚本在 code\ 下，默认路径按项目根目录拼，所以在任何工作目录下直接运行都能找到表。
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_file_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        root_dir, '星铁装备刷取-自动生成-豆包.xlsx')
    excel_save_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(root_dir, '分析结果-豆包.xlsx')

    cache_dict_list = []

    # 初始化主词条字典
    init_equip_main_dict()
    # 初始化副词条字典
    init_equip_res_dict()

    # 读取装备相关信息（简称/刷取地点）
    read_excel_equip_info(excel_file_path)

    # 读取角色信息
    character_info_list = read_excel_character_info(excel_file_path)

    # 计算角色各部位理论上限
    cal_character_equip_update_max_num(character_info_list)

    # 核心计算：每角色 × 每部位的概率分布
    update_character_equip_anylize(character_info_list)

    # 星级映射：name → "4"/"5"。只给输出表加列用，计算链一个数都不看它。
    star_map = {ch.name: ch.star for ch in character_info_list}

    # 视角一：角色
    sort_eq_list, sort_eq_once_list = show_character_equipment_list(character_info_list)
    cache_dict_list.append(get_save_cache_dict(sort_eq_list, ['套装名称', '角色名称', '角色星级', '套装提升概率'],
                                               sheet_name='角色套装优化概率', add_exp_num=exp_rate, p_ind=2, star_map=star_map))
    cache_dict_list.append(get_save_cache_dict(sort_eq_once_list, ['部位名称', '角色名称', '角色星级', '套装提升概率', '当前有效词条数量', '理论最高有效词条数量'],
                                               sheet_name='角色套装部位优化概率', add_exp_num=exp_rate, p_ind=2, star_map=star_map))

    # 视角二：套装 / 地点（旧口径，简单相加）
    sort_equip_match_list, sort_pose_list = load_character_to_equipment(character_info_list)

    # 视角三：套装 / 地点（v2.0 准确口径，去重）
    eq_pose_list, eq_list, pos_list = load_character_to_equipment_by_key(character_info_list)

    for i in pos_list:
        print(i)
    cache_dict_list.append(get_save_cache_dict(eq_pose_list, ['套装名称', '部位名称', '该套装使用角色列表', '套装部位提升概率'],
                                               sheet_name='套装部位优化概率', add_exp_num=exp_rate, p_ind=3))
    cache_dict_list.append(get_save_cache_dict(eq_list, ['套装名称', '该套装使用角色列表', '套装提升概率'],
                                               sheet_name='套装优化准确概率', add_exp_num=exp_rate, p_ind=2))
    cache_dict_list.append(get_save_cache_dict(pos_list, ['刷取地点', '地点提升概率', '可刷取套装列表', '该地点可提升角色列表'],
                                               sheet_name='地点优化概率', add_exp_num=exp_rate, p_ind=1))

    cache_dict_list.append(get_save_cache_dict(sort_equip_match_list, ['套装名称', '套装提升概率', '该套装使用角色列表'],
                                               sheet_name='(旧)套装优化概率', add_exp_num=exp_rate, p_ind=1))
    cache_dict_list.append(get_save_cache_dict(sort_pose_list, ['刷取地点', '地点提升概率', '可刷取套装列表', '该地点可提升角色列表'],
                                               sheet_name='(旧)地点优化概率', add_exp_num=exp_rate, p_ind=1))

    # 视角四：加权口径（把上面旧口径的结果 × 角色权重，另开页签，不改动前面任何数值）
    #   顺序不能颠倒：必须先 normalize_character_weight() 把权重缩放成 mean=1，
    #   再进 load_character_to_equipment_weighted()，否则用到的还是 Excel 里填的原始值，
    #   加权得分会随"整体填大填小"而漂移。
    #   这三页是【追加】的，前面七页的数值一个都没动过；
    #   其中「未加权得分」这一列正好可以用来校验：它必须严格等于「(旧)套装优化概率」×100。
    normalize_character_weight(character_info_list)
    w_eq_list, w_pos_list, w_char_list, w_slot_list = load_character_to_equipment_weighted(character_info_list)
    cache_dict_list.append(get_save_cache_dict(w_eq_list,
                                               ['套装名称', '角色列表', '未加权得分', '加权得分', '加权倍数', '涉及角色数'],
                                               sheet_name='(加权)套装得分'))
    cache_dict_list.append(get_save_cache_dict(w_pos_list,
                                               ['刷取地点', '可刷取套装列表', '未加权得分', '加权得分', '加权倍数', '角色列表'],
                                               sheet_name='(加权)地点得分'))
    # 这一页不加 add_exp_num：得分不是概率，套"期望出货次数"的换算公式会算出垃圾数
    cache_dict_list.append(get_save_cache_dict(w_char_list,
                                               ['角色名称', '角色星级', '原始权重', '归一化权重', '当前有效词条合计',
                                                '理论最高合计', '当前完成度%',
                                                '一般最高合计', '一般完成度%'],
                                               sheet_name='角色权重', star_map=star_map))
    # 逐部位明细：一行一个角色的一个部位（55 × 6 = 330 行）
    # 后面的 cost 定向优化就用这张表当输入：标的 = (套装, 部位, 主词条)，基线 = 当前有效词条数
    cache_dict_list.append(get_save_cache_dict(w_slot_list,
                                               ['角色名称', '角色星级', '归一化权重', '部位', '套装', '主词条',
                                                '当前有效词条数', '部位理论上限', '理论完成度%',
                                                '部位一般上限', '一般完成度%',
                                                '剩余空间(理论)', '剩余空间(一般)',
                                                '提升概率(1cost)%', '提升概率(2cost)%', '锁定副词条(2cost)',
                                                '提升概率(5cost)%', '锁定副词条(5cost)'],
                                               sheet_name='角色部位完成度', star_map=star_map))

    # ---- 输出五：遗器分解建议（v2.9）----
    # 阈值来自生成表「分解阈值」页（01 从配置表原样带过来的），改口径只改配置表，不碰代码。
    # 这一页只读 character_info_list 的现状（套装 / 部位 / 当前有效条数 / 有效种类上限），
    # 算的是"强化能不能翻身"，不参与前面任何概率计算，也不影响它们。
    dis_low, dis_high, dis_strict = read_dismantle_thresholds(excel_file_path)
    dis_list = build_dismantle_sheet(character_info_list,
                                     low=dis_low, high=dis_high, strict=dis_strict)
    cache_dict_list.append(get_save_cache_dict(
        dis_list,
        ['套装名称', '部位名称', '使用者数', '现役最高词条数', '现役最低词条数', '初始可中上限',
         '升级门槛 s≥(4条)', '升级门槛 s≥(3条)',
         '保底门槛 s≥(4条)', '保底门槛 s≥(3条)',
         's=2 达标率%(4条/9投)', 's=2 达标率%(3条/8投)',
         's=3 达标率%(4条/9投)', 's=3 达标率%(3条/8投)',
         '升级结论（超现役最高）', '保底结论（歪2条超现役最低）', '能帮到的角色'],
        sheet_name='分解建议'))

    save_excel(cache_dict_list, save_path=excel_save_path)

    # 分解建议的摘要：控制台能一眼看到"有多少套装-部位完全没保留必要"
    n_dump = len([r for r in dis_list if '直接分解' in r[14]])
    n_no_floor = len([r for r in dis_list if '无保留必要' in r[15]])
    print("分解建议 %d 行（套装×部位）：升级口径直接分解 %d 行；保底口径无保留必要 %d 行；阈值 低%.0f%% / 高%.0f%% / %s"
          % (len(dis_list), n_dump, n_no_floor, dis_low, dis_high,
             '要求严格超过现役' if dis_strict else '打平也算'))

    # 可选：控制台打印四类结论
    if show_log:
        print('====================================展示角色套装刷取优化概率====================================')
        for i in sort_eq_list:
            p = i[2]
            print(i, round(exp_rate * 100, 2), '%概率期望出货次数', cal_exp_numb_from_rate(p, exp_rate))

        print('====================================展示角色部位刷取优化概率====================================')
        for i in sort_eq_once_list:
            p = i[2]
            print(i, round(exp_rate * 100, 2), '%概率期望出货次数', cal_exp_numb_from_rate(p, exp_rate))

        print('====================================展示套装刷取优化概率====================================')
        for i in sort_equip_match_list:
            p = i[1]
            print(i, round(exp_rate * 100, 2), '%概率期望出货次数', cal_exp_numb_from_rate(p, exp_rate))

        print('====================================展示刷取地点刷取优化概率====================================')
        for i in range(len(sort_pose_list)):
            p = sort_pose_list[i][1]
            print(i, sort_pose_list[i], round(exp_rate * 100, 2), '%概率期望出货次数', cal_exp_numb_from_rate(p, exp_rate))


if __name__ == "__main__":
    val_show_log = False      # 想在看控制台看明细就改成 True
    # val_show_log=True
    val_exp_rate = 0.95       # 期望出货的置信度：95%
    main(show_log=val_show_log, exp_rate=val_exp_rate)
