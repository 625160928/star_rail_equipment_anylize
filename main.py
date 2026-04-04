import pandas as pd
import numpy as np
import copy

#装备简称至原名的对照表
equip_name_dict=dict()
#通过装备原名查询到装备的刷取地点的对照表
equip_from_dict=dict()
#通过装备原名到角色需求的放置表
equip_match_character_dict=dict()

class character_equip():
    def __init__(self,name,eq4,eq2,val,val_list,main_yifu,main_xiezi,main_qiu,main_shengzi
                 ,val_num_tou,val_num_shou,val_num_yifu,val_num_xiezi,val_num_qiu,val_num_shengzi):
        self.name=name
        self.eq4=eq4
        self.eq2=eq2
        self.val =val
        self.val_list =val_list
        self.main_yifu =main_yifu
        self.main_xiezi =main_xiezi
        self.main_qiu =main_qiu
        self.main_shengzi =main_shengzi
        self.val_num_tou =val_num_tou
        self.val_num_shou =val_num_shou
        self.val_num_yifu =val_num_yifu
        self.val_num_xiezi =val_num_xiezi
        self.val_num_qiu =val_num_qiu
        self.val_num_shengzi =val_num_shengzi


    @property
    def total_val(self):
        return self.val_num_tou+self.val_num_shou+self.val_num_yifu+self.val_num_xiezi+self.val_num_qiu+self.val_num_shengzi

    #更新信息
    def update_info(self):
        self.pos=[]
        for i in range(len(self.eq4)):
            name=self.eq4[i]
            ori_name=equip_name_dict[name]
            pos=equip_from_dict[ori_name]
            # print(name,ori_name,pos)
            self.eq4[i]=ori_name
            self.pos.append(pos)
        name = self.eq2
        ori_name = equip_name_dict[name]
        pos = equip_from_dict[ori_name]
        self.eq2=ori_name
        self.pos.append(pos)
        return


#读取装备相关信息，比如简称，刷取位置等
def read_excel_equip_info(excel_file_path):

    #读取excel
    df = pd.read_excel(excel_file_path, sheet_name='遗器相关')
    #读取行名称
    attr_list = df.columns
    #转化为numpy便于理解，dataframe没怎么用过
    character_list = df.to_numpy()

    # print(attr_list)
    for ori_name,pos,sp_name in character_list:
        equip_from_dict[ori_name]=pos
        equip_match_character_dict[ori_name]=[]
        if '、' in sp_name:
            sp=sp_name.split('、')
            for i in sp:
                equip_name_dict[i]=ori_name
        else:
            equip_name_dict[sp_name]=ori_name

    # print(equip_from_dict)
    # print('-----')
    # print(equip_name_dict)
    # print('-----')

    return

#转化角色信息
#将list的角色信息转入class类里
def exchang_character_info(c_info):
    c_name=c_info[0]
    eq4=c_info[1]
    if '+'in eq4:
        c_eq4=eq4.split('+')
    else:
        c_eq4=[eq4]
    c_eq2=c_info[2]
    c_val=int(c_info[3])
    c_val_list=c_info[4].split('、')

    main_yifu=c_info[5]
    main_xiezi=c_info[6]
    main_qiu=c_info[7]
    main_shengzi=c_info[8]

    val_num_tou=c_info[9]
    val_num_shou=c_info[10]
    val_num_yifu=c_info[11]
    val_num_xiezi=c_info[12]
    val_num_qiu=c_info[13]
    val_num_shengzi=c_info[14]
    total_val_num=val_num_tou+val_num_shou+val_num_yifu+val_num_xiezi+val_num_qiu+val_num_shengzi

    # print('角色-',c_name,' 的外圈装备是',c_eq4,'； 内圈是',c_eq2,'; 有效词条数量是 ',c_val,'； 有效词条是',c_val_list
    #       ,' 数量是否对应 ',c_val==len(c_val_list),'; \n衣服主词条是 ',main_yifu,'; 鞋子主词条是 ',main_xiezi
    #       ,'; 球主词条是 ',main_qiu,'; 绳子主词条是 ',main_shengzi,)
    # print('角色-',c_name,'的有效词条数量是： 头: ',val_num_tou,'; 手: ',val_num_shou,'; 衣服: ',val_num_yifu ,
    #       '; 鞋子: ',val_num_xiezi,'; 球: ',val_num_qiu,'; 绳子: ',val_num_shengzi,'； 总有效词条数量为: ',total_val_num,total_val_num==int(c_info[15]))

    ch_eq=character_equip(c_name,c_eq4,c_eq2,c_val,c_val_list,main_yifu,main_xiezi,main_qiu,main_shengzi
                 ,val_num_tou,val_num_shou,val_num_yifu,val_num_xiezi,val_num_qiu,val_num_shengzi)
    ch_eq.update_info()
    return ch_eq

#从excel表格中读取角色相关信息并保存到class类的list里
def read_excel_character_info(excel_file_path):
    ret_list=[]

    #读取excel
    df = pd.read_excel(excel_file_path, sheet_name='角色遗器')
    #读取行名称
    attr_list = df.columns
    #转化为numpy便于理解，dataframe没怎么用过
    character_list = df.to_numpy()
    #获取角色姓名
    name_list = character_list[:, 0]
    # print(name_list)
    # print('--------------')
    # print(attr_list)
    # print('--------------')

    #枚举每个角色
    for i in range(len(character_list)):
        c_info=character_list[i]
        # print('=======')
        # print(c_info)
        c_data=exchang_character_info(c_info)
        #将获取的角色信息添加至list
        ret_list.append(c_data)

    return ret_list

#装备词条分析
#输入为主词条名称（用于移除有效词条中重复的），有效词条列表，当前有效词条数量，部位
def equip_anylize(main_skill,useful_skills,now_skill_num,pos_name):
    ret=None
    re_useful = copy.deepcopy(useful_skills)
    if main_skill in re_useful:
        re_useful.remove(main_skill)


    print(pos_name,'主词条需求为',main_skill,' 有效词条为 ',re_useful,' 当前有效数量为 ',now_skill_num)
    return ret

#从角色信息挂在到装备刷取里面
def update_character_equip_match_info(character_info_list):
    skill_list=set()
    for ch in character_info_list:
        print('===================================')
        print(ch.name,ch.eq4,ch.eq2)
        print('该角色有效副词条为',ch.val_list)

        for i in ch.val_list:
            skill_list.add(i)
        skill_list.add(ch.main_yifu)
        skill_list.add(ch.main_xiezi)
        skill_list.add(ch.main_qiu)
        skill_list.add(ch.main_shengzi)

        for eq in ch.eq4:
            print('1. 外圈装备为',eq,' 衣服鞋子主词条为 ',ch.main_yifu,ch.main_xiezi,
                  '当前有效词条数量为',ch.val_num_tou,ch.val_num_shou,ch.val_num_yifu,ch.val_num_xiezi)
            ret=equip_anylize('小生命',ch.val_list,ch.val_num_tou,'tou')
            ret=equip_anylize('小攻击',ch.val_list,ch.val_num_shou,'shou')
            ret=equip_anylize(ch.main_yifu,ch.val_list,ch.val_num_yifu,'yifu')
            ret=equip_anylize(ch.main_xiezi,ch.val_list,ch.val_num_xiezi,'xiezi')

        print('2. 内圈装备为',ch.eq2,' 球和绳子主词条为 ',ch.main_qiu,ch.main_shengzi,
              '当前有效词条数量为',ch.val_num_qiu,ch.val_num_shengzi)

        ret = equip_anylize(ch.main_qiu, ch.val_list, ch.val_num_qiu, 'qiu')
        ret = equip_anylize(ch.main_shengzi, ch.val_list, ch.val_num_shengzi, 'shengzi')

        print('-----')


    lsl=list(skill_list)
    lsl.sort()
    print('词条一览 ',lsl)
    return

def main():
    excel_file_path = '星铁装备刷取.xlsx'

    #读取装备相关信息
    read_excel_equip_info(excel_file_path)

    # 读取Excel文件
    character_info_list=read_excel_character_info(excel_file_path)

    update_character_equip_match_info(character_info_list)
    # print('=====')
    # for ch in character_info_list:
    #     print(ch)


if __name__ == "__main__":
    main()