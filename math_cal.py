import copy
import itertools
from itertools import combinations
import numpy as np


#副词条权重字典
equip_res=dict()
res_skill_list=[]
e_equip34={3:0.8,4:0.2}

#加速计算的缓存
cal_cache_dict=dict()

#在初始词条为3词条的情况下，4+4，后续4条词条命中的概率分布
#横坐标为初始命中数，纵坐标为后续4条升级后的有效词条命中数量
p_arr3=np.array([[1,0,0,0,0],
        [81/256, 108/256, 54/256, 12/256,1/256],
        [1 / 16, 4/16,6/16,4/16,1/16],
        [1 / 256, 12 / 256, 54 / 256, 108 / 256 , 81 / 256],
        [0,0,0,0,1]
        ])

p_arr4=np.array([
    [1,0,0,0,0,0],
    [243/1024,405/1024,270/1024,90/1024,15/1024,1/1024],
    [1/32,5/32,10/32,10/32,5/32,1/32],
    [1/1024,15/1024,90/1024,270/1024,405/1024,243/1024],
    [0,0,0,0,0,1]
])

perm_order = list(itertools.permutations([0, 1, 2, 3], 4))

def init_equip_res_dict():
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
        ('击破',80)
    ]
    for name, p in res_map:
        equip_res[name] = p
        res_skill_list.append(name)


#在已经确认抽取哪些词条的基础上，计算该词条的不同顺序的概率的总和，作为这个词条的抽取概率
def cal_p_once(skills,w_arr,w_total):
    # print('cal_p_once ',skills,'------------------',w_arr,w_total)
    w_mu=w_arr[0]*w_arr[1]*w_arr[2]*w_arr[3]
    # print(w_mu)
    p_list=[]
    for i in range(len(perm_order)):
        p=w_mu
        res=w_total
        for j in range(4):
            # print(j,p,res,'-',w_arr[perm_order[i][j]],'=',res-w_arr[perm_order[i][j]],p/res)
            p=p/res
            res-=w_arr[perm_order[i][j]]
            # p_arr.append(w_arr[perm_order[i][j]])

        # print(i, perm_order[i],p)
        # print('------')
        # print(p_arr)
        p_list.append(p)
    return sum(p_list)


#通过有效词条和无效词条，输出不同有效词条数量的词条组合的list
def cal_p_at_need_combine(use_name_list,res_name_list):
    ans=[]

    #计算本次词条总权重
    weight_total=0
    for su in use_name_list:
        weight_total+=equip_res[su]
    for sr in res_name_list:
        weight_total+=equip_res[sr]

    for useful_num in range(1,min(5,len(use_name_list)+1)):
        combine_use_list = list(combinations(use_name_list, useful_num))
        combine_res_list = list(combinations(res_name_list, 4-useful_num))

        for cu in combine_use_list:
            for cr in combine_res_list:
                skills = cu + cr


                w_arr = []
                for s in skills:
                    w_arr.append(equip_res[s])
                w_arr.sort()

                # 准备查询是否存在已经计算的结果，如果有就使用缓存
                # 准备部分
                ck = copy.deepcopy(w_arr)
                ck.append(weight_total)
                ck = tuple(ck)
                # 查询部分
                if ck in cal_cache_dict:
                    p = cal_cache_dict[ck]
                    # print('快查 ',skills,ck,p)
                else:
                    p = cal_p_once(skills, w_arr, weight_total)

                    # print('【计算】 ',skills,ck,p)
                    cal_cache_dict[ck] = p

                ans.append((useful_num,skills, ck, p))
    return ans


#在指定命中词条数的情况下，枚举每一种组合情况，并计算每种组合概率的和作为命中词条数的概率
def cal_p_at_need_eff(use_name_list,res_name_list,equip_name_weight_dict,useful_num,res_num):
    # if useful_num>len(use_name_list):
    #     return 0
    combine_use_list = list(combinations(use_name_list, useful_num))
    combine_res_list = list(combinations(res_name_list, res_num))

    # print('----')
    # print('选取数量为 ',useful_num,res_num,'--',use_name_list,res_name_list)
    # print('有效组合为 ',len(combine_use_list),combine_use_list)
    # print('无效组合为 ',len(combine_res_list),combine_res_list)


    t_numb=len(use_name_list)+len(res_name_list)

    #C_(t_numb)_4,从t_numb里抽取4个的组合方式
    comb_total_numb=t_numb/24
    for i in range(1,4):
        comb_total_numb*=t_numb-i


    # print('=================')
    # print('有效词条数量为：',len(use_name_list),use_name_list)
    # print('其他副词条数量为：',len(res_name_list),res_name_list)
    # print('目前需要计算在有效词条数量为 ',useful_num,'，无效词条数量为 ',res_num,' 的情况')
    # print('有效组合数量为 ',len(combine_use_list),' 无效组合数量为 ',len(combine_res_list))
    # print('词条总数为',t_numb,' ,在当前情况下，词条的组合方式总数为: ',int(comb_total_numb),
    #       '符合目前要求的组合数量为',len(combine_use_list)*len(combine_res_list))
    # print(res_skill_list)
    weight_total=0
    for su in use_name_list:
        weight_total+=equip_res[su]
        # print(su,equip_res[su])
    for sr in res_name_list:
        weight_total+=equip_res[sr]
        # print(sr,equip_res[sr])
    # print('总权重为 ',weight_total)

    # p_total_list 储存每个组合的概率
    p_total_list=[]
    for cu in combine_use_list:
        for cr in combine_res_list:
            skills=cu+cr
            w_arr=[]
            # print('cucr ',skills)
            for s in skills:
                w_arr.append(equip_res[s])
            w_arr.sort()
            # print(skills,w_arr,weight_total)
            #准备查询是否存在已经计算的结果，如果有就使用缓存
            #准备部分
            ck=copy.deepcopy(w_arr)
            ck.append(weight_total)
            ck=tuple(ck)
            #查询部分
            if ck in cal_cache_dict:
                p=cal_cache_dict[ck]
                # print('快查 ',skills,ck,p)
            else:
                p=cal_p_once(skills,w_arr,weight_total)

                # print('【计算】 ',skills,ck,p)
                cal_cache_dict[ck]=p

            p_total_list.append((skills,ck,p))

    p_final_total=0
    for sk,w,p in p_total_list:
        # print(sk,w,p)
        p_final_total+=p
    # print('概率为 ',p_final_total)
    # return len(combine_use_list)*len(combine_res_list)/comb_total_numb
    return p_final_total
    # return 0.2

#计算初始四词条中，四词条里有效词条的概率分布
def cal_eff_when_init4(use_name_list,res_name_list,equip_name_weight_dict):
    p_list=[]
    for i in range(5):
    # for i in range(min(5,len(use_name_list)+1)):
        ret=cal_p_at_need_eff(use_name_list,res_name_list,equip_name_weight_dict,i,4-i)
        p_list.append(ret)
        # print('有效词条数量为',i,'的概率分布为--',ret)
    while len(p_list)<5:
        p_list.append(0)
    return np.array(p_list)


#在设置初始有效词条数量下，获取词条最终有效词条的概率分布
def cal_p_when_init_val_num(val_num):
    # 计算3/4词条里，在当前条件下有效词条的概率分布
    # 获取初始三词条中，三词条里有效词条的概率分布
    p_init3_dist = np.array([0,0,0,0,0])
    p_init3_dist[val_num]=1
    # 获取初始四词条中，四词条里有效词条的概率分布，两者概率分布相同，强化数量不同
    p_init4_dist = copy.deepcopy(p_init3_dist)

    # print('=====================')
    # print('初始状态下四词条，有效词条数量概率分布为',p_init4_dist,'总和为',sum(p_init4_dist))

    # region 计算三词条概率分布
    # 计算三词条中不同有效词条最终概率分布,最终表现为可以强化4次
    p_arr3_copy = copy.deepcopy(p_arr3)
    p_dict3 = np.zeros(10)
    # 然后计算在当前条件下最终有效词条数量的概率分布
    for init_hit_num in range(5):
        p_arr3_copy[init_hit_num] *= p_init3_dist[init_hit_num]
    for i in range(5):
        for j in range(5):
            p_dict3[i + j] += p_arr3_copy[i][j]
    # print(p_dict3)
    # endregion

    # region 计算四词条概率分布
    # 计算四词条中不同有效词条最终概率分布,最终表现为可以强化5次
    p_arr4_copy = copy.deepcopy(p_arr4)
    p_dict4 = np.zeros(10)
    # 然后计算在当前条件下最终有效词条数量的概率分布
    for init_hit_num in range(5):
        p_arr4_copy[init_hit_num] *= p_init4_dist[init_hit_num]
    for init_hit_num in range(5):
        for j in range(6):
            p_dict4[init_hit_num + j] += p_arr4_copy[init_hit_num][j]
    # print(p_dict4)
    # endregion

    p_total = e_equip34[3] * p_dict3 + e_equip34[4] * p_dict4

    return p_total
    return 0
'''
输入为：有用的词条列表，其他可选副词条列表、每个词条对应的权重
输出为：【p0,p1,p2,p3,p4,p5,p6,p7,p8,p9】，为最终获取得到的装备中有0-9个有效词条的概率

'''
def cal_p(use_name_list,res_name_list,equip_name_weight_dict):
    #计算3/4词条里，在当前条件下有效词条的概率分布
    #获取初始三词条中，三词条里有效词条的概率分布
    p_init3_dist=cal_eff_when_init4(use_name_list,res_name_list,equip_name_weight_dict)
    #获取初始四词条中，四词条里有效词条的概率分布，两者概率分布相同，强化数量不同
    p_init4_dist=copy.deepcopy(p_init3_dist)

    # print('=====================')
    # print('初始状态下四词条，有效词条数量概率分布为',p_init4_dist,'总和为',sum(p_init4_dist))

    #region 计算三词条概率分布
    #计算三词条中不同有效词条最终概率分布,最终表现为可以强化4次
    p_arr3_copy=copy.deepcopy(p_arr3)
    p_dict3=np.zeros(10)
    #然后计算在当前条件下最终有效词条数量的概率分布
    for init_hit_num in range(5):
        p_arr3_copy[init_hit_num]*=p_init3_dist[init_hit_num]
    for i in range(5):
        for j in range(5):
            p_dict3[i+j]+=p_arr3_copy[i][j]
    # print(p_dict3)
    #endregion

    # region 计算四词条概率分布
    # 计算四词条中不同有效词条最终概率分布,最终表现为可以强化5次
    p_arr4_copy = copy.deepcopy(p_arr4)
    p_dict4 = np.zeros(10)
    # 然后计算在当前条件下最终有效词条数量的概率分布
    for init_hit_num in range(5):
        p_arr4_copy[init_hit_num] *= p_init4_dist[init_hit_num]
    for init_hit_num in range(5):
        for j in range(6):
            p_dict4[init_hit_num + j] += p_arr4_copy[init_hit_num][j]
    # print(p_dict4)
    # endregion

    p_total=e_equip34[3]*p_dict3+e_equip34[4]*p_dict4

    return p_total

def main():
    q=[
        (['速度'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '大攻击', '暴击', '暴伤']),
        (['速度', '大攻击'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '暴击', '暴伤']),
        (['速度', '大攻击', '暴击'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破', '暴伤']),
        (['速度', '大攻击', '暴击', '暴伤'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       # (['速度', '大攻击', '暴击', '暴伤'],['小生命', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       # (['速度', '大攻击', '暴伤'] ,['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       # (['速度', '暴击', '暴伤'] ,['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       # (['速度', '大攻击', '暴击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       # (['速度', '暴击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破'])
    ]

    for quse,qres in q:
        p_total=cal_p(quse,qres,equip_res)
        print('-----------------------------------')
        print(quse,qres)
        print('最终满级不同有效词条数量的概率分布为 ',p_total)
        break

    return


init_equip_res_dict()

if __name__ == "__main__":
    main()