import copy

import numpy as np


#副词条权重字典
equip_res=dict()
res_skill_list=[]
e_equip34={3:0.8,4:0.2}

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
        ('效果抵抗', 100),
        ('效果命中', 100),
        ('击破', 100)
    ]
    for name, p in res_map:
        equip_res[name] = p
        res_skill_list.append(name)


'''
输入为：有用的词条列表，其他可选副词条列表、每个词条对应的权重
输出为：【p0,p1,p2,p3,p4,p5,p6,p7,p8,p9】，为最终获取得到的装备中有0-9个有效词条的概率

'''
def cal_p(use_name_list,res_name_list,equip_name_weight_dict):
    #计算3/4词条里，在当前条件下有效词条的概率分布





    #获取初始三词条中，三词条里有效词条的概率分布
    p_init3_dist=np.array([0.25,0.25,0.25,0.25,0])
    #获取初始四词条中，四词条里有效词条的概率分布
    p_init4_dist=np.array([0.2,0.2,0.2,0.2,0.2])

    #region 计算三词条概率分布
    #计算三词条中不同有效词条最终概率分布
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
    # 计算四词条中不同有效词条最终概率分布
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
    q=[(['速度', '大攻击', '暴击', '暴伤'], ['小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       (['速度', '大攻击', '暴击', '暴伤'],['小生命', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       (['速度', '大攻击', '暴伤'] ,['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       (['速度', '暴击', '暴伤'] ,['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       (['速度', '大攻击', '暴击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破']),
       (['速度', '暴击', '暴伤'], ['小生命', '小攻击', '小防御', '大生命', '大防御', '效果抵抗', '效果命中', '击破'])]

    for quse,qres in q:
        p_total=cal_p(quse,qres,equip_res)

        print(p_total)
        break

    return


if __name__ == "__main__":
    main()