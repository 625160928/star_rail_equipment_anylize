import pandas as pd
import numpy as np
import copy
import math_cal

#装备简称至原名的对照表
equip_name_dict=dict()
#通过装备原名查询到装备的刷取地点的对照表
equip_from_dict=dict()
#通过装备原名到角色需求的放置表
equip_match_character_dict=dict()
#主词条出货概率字典
equip_main=dict()
#副词条权重字典
equip_res=dict()
res_skill_list=[]

class character_equip():
    def __init__(self,name,eq4,eq2,val,val_list,main_yifu,main_xiezi,main_qiu,main_shengzi
                 ,val_num_tou,val_num_shou,val_num_yifu,val_num_xiezi,val_num_qiu,val_num_shengzi):
        self.name=name
        self.eq4=eq4
        self.eq2=eq2
        self.val_num =val
        self.val_name_list =val_list
        self.main_yifu =main_yifu
        self.main_xiezi =main_xiezi
        self.main_qiu =main_qiu
        self.main_shengzi =main_shengzi

        #储存当前有效词条数量
        self.val_num_tou =val_num_tou
        self.val_num_shou =val_num_shou
        self.val_num_yifu =val_num_yifu
        self.val_num_xiezi =val_num_xiezi
        self.val_num_qiu =val_num_qiu
        self.val_num_shengzi =val_num_shengzi
        self.val_num_list=[val_num_tou,val_num_shou,val_num_yifu,val_num_xiezi,val_num_qiu,val_num_shengzi]

        #储存每个部位装备的分析，里面村的list
        self.eq4_anylize=[]
        self.eq2_anylize=[]


    #计算目前有效词条综述
    @property
    def total_val(self):
        return self.val_num_tou+self.val_num_shou+self.val_num_yifu+self.val_num_xiezi+self.val_num_qiu+self.val_num_shengzi

    #更新信息，excel输入的装备信息都是简称，后续处理都按照原名处理
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

    #获取六件装备里每件能获取更好词条的概率
    def p_get_better6(self):
        ret=[]
        for i in range(4):
            eq_anylize_list=self.eq4_anylize[i]
            eq_val=self.val_num_list[i]
            p_better=sum(eq_anylize_list[eq_val+1:])
            ret.append(p_better)
            # print('eq4 ',i, eq_anylize_list,'---', self.val_num_list, eq_val,',',p_better)
        # sum(p_list[now_skill_num + 1:])

        for i in range(2):
            eq_anylize_list=self.eq2_anylize[i]
            eq_val=self.val_num_list[4+i]
            p_better=sum(eq_anylize_list[eq_val+1:])
            ret.append(p_better)
        #     print('eq2 ',i, eq_anylize_list,'---', self.val_num_list, eq_val,',',p_better)
        # print(ret)
        self.p_better=ret
        return

        # 初始化主词条概率分析字典

def init_equip_main_dict():
    yifu_list = [['大生命', 0.2], ['大攻击', 0.2], ['大防御', 0.2],
                 ['暴击', 0.1], ['暴伤', 0.1], ['治疗量', 0.1], ['效果命中', 0.1]]
    xiezi_list = [['大生命', 0.2917], ['大攻击', 0.2917], ['大防御', 0.2917],
                  ['速度', 0.125]]
    shengzi_list = [['大生命', 0.2633], ['大攻击', 0.2633], ['大防御', 0.2633],
                    ['击破', 0.15], ['充能', 0.06]]
    qiu_list = [['大生命', 0.1233], ['大攻击', 0.1233], ['大防御', 0.1233]
        , ['冰伤', 0.09], ['火伤', 0.09], ['物伤', 0.09], ['虚数伤害', 0.09], ['量子伤', 0.09], ['雷伤', 0.09],
                ['风伤', 0.09]]

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

# 初始化副词条字典
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

repp_list=set()

#根据有效词条列表和无效词条列表计算不同有效词条数量的概率（0-8/9）
#最后返回获取不同数量词条的概率（0-9）
def equip_anylize_cal(use_list,res_list):
    repp_list.add((len(use_list),len(res_list)))
    # print(use_list,res_list)
    ret=math_cal.cal_p(use_name_list=use_list,res_name_list=res_list,equip_name_weight_dict=equip_res)

    return ret

#装备词条分析
#输入为主词条名称（用于移除有效词条中重复的），有效词条列表，当前有效词条数量，部位
#返回是该装备出现不同数量词条的概率（包括主词条+副词条概率）
#然后再把一次刷新2/3金概率折算后的1.05乘进去了
#然后还要把部位的概率下降算进去，四件套*0.25，两件套*0.5
def equip_anylize(main_skill,useful_skills,now_skill_num,pos_name):
    re_useful = copy.deepcopy(useful_skills)
    if main_skill in re_useful:
        re_useful.remove(main_skill)

    tmp_skill_list=copy.deepcopy(res_skill_list)
    if main_skill in tmp_skill_list:
        tmp_skill_list.remove(main_skill)
    for sk in re_useful:
        if sk in tmp_skill_list:
            tmp_skill_list.remove(sk)
    p_list=equip_anylize_cal(re_useful,tmp_skill_list)
    p_list_add_main=p_list*equip_main[(pos_name,main_skill)]
    # print('----------')
    # print(pos_name,p_list_add_main)
    if pos_name in ['头','手','衣服','鞋子']:
        p_list_add_main*=0.25
    else:
        p_list_add_main*=0.5
    # print(pos_name,p_list_add_main)
    # print('------------------')
    # print(pos_name,'主词条需求为',main_skill,' 有效词条为 ',re_useful,' 当前有效数量为 ',now_skill_num,' 主词条概率为 ',equip_main[(pos_name,main_skill)])
    # print(equip_main[(pos_name,main_skill)],len(re_useful),len(tmp_skill_list),main_skill,'------',re_useful,'/',tmp_skill_list)
    # print('在当前情况下，不考虑主词条刷取概率的情况下,该装备获得不同数量词条的概率为 ',p_list)
    # print('主词条出货概率为',equip_main[(pos_name,main_skill)],p_list_add_main)
    # print('想要获取比',now_skill_num,' 更好的数量的词条的概率为 ',sum(p_list[now_skill_num+1:]))
    # print('------------------')

    return p_list_add_main*1.05

#从角色装备信息挂在信息类里面,顺便计算每个角色的装备刷取出更好的概率
def update_character_equip_anylize(character_info_list):
    skill_list=set()
    for ch in character_info_list:
        # print('===================================')
        # print('-----')
        # print(ch.name,ch.eq4,ch.eq2)
        # print('该角色有效副词条为',ch.val_name_list)

        for i in ch.val_name_list:
            skill_list.add(i)
        skill_list.add(ch.main_yifu)
        skill_list.add(ch.main_xiezi)
        skill_list.add(ch.main_qiu)
        skill_list.add(ch.main_shengzi)

        # for eq in ch.eq4:
        #     print('1. 外圈装备为',eq,' 衣服鞋子主词条为 ',ch.main_yifu,ch.main_xiezi,
        #           '当前有效词条数量为',ch.val_num_tou,ch.val_num_shou,ch.val_num_yifu,ch.val_num_xiezi)
        ret_tou=equip_anylize('小生命', ch.val_name_list, ch.val_num_tou, '头')
        ret_shou=equip_anylize('小攻击', ch.val_name_list, ch.val_num_shou, '手')
        ret_yifiu=equip_anylize(ch.main_yifu, ch.val_name_list, ch.val_num_yifu, '衣服')
        ret_xiezi=equip_anylize(ch.main_xiezi, ch.val_name_list, ch.val_num_xiezi, '鞋子')
        ch.eq4_anylize=(ret_tou,ret_shou,ret_yifiu,ret_xiezi)
        # print('2. 内圈装备为',ch.eq2,' 球和绳子主词条为 ',ch.main_qiu,ch.main_shengzi,
        #       '当前有效词条数量为',ch.val_num_qiu,ch.val_num_shengzi)

        ret_qiu = equip_anylize(ch.main_qiu, ch.val_name_list, ch.val_num_qiu, '球')
        ret_shengzi = equip_anylize(ch.main_shengzi, ch.val_name_list, ch.val_num_shengzi, '绳子')
        ch.eq2_anylize=(ret_qiu,ret_shengzi)
        ch.p_get_better6()

    lsl=list(skill_list)
    lsl.sort()
    print('词条一览 ',lsl)
    print('需计算概率 ',repp_list)
    return


#获取最后一个概率为0的词条数量
def get_last_zero(arr):
    if arr[-1]!=0:
        return 9
    i=9
    for i in range(len(arr)-1,0,-1):
        if i!=0:
            return i
    return i
#以角色为核心，进行遗器优化概率的排序
def show_character_equipment_list(character_info_list):
    eq_list=[]
    eq_once_list=[]
    for ch in character_info_list:
        # print('----------------')
        # print(ch.name,ch.eq4,ch.eq2,ch.p_better)
        eq4_better=sum(ch.p_better[:4])/4
        eq2_better=sum(ch.p_better[4:])/2
        # print(ch.name,ch.eq4,ch.eq2,eq4_better,eq2_better)
        # print('外圈分析 ',ch.eq4_anylize)
        # print('内圈分析 ',ch.eq2_anylize)
        for i in ch.eq4:
            eq_list.append(('四件套-'+i,ch.name,eq4_better))
            eq_once_list.append(('四件套-'+i+'-头部',ch.name,ch.p_better[0],ch.val_num_tou,get_last_zero(ch.eq4_anylize[0])))
            eq_once_list.append(('四件套-'+i+'-手部',ch.name,ch.p_better[1],ch.val_num_shou,get_last_zero(ch.eq4_anylize[1])))
            eq_once_list.append(('四件套-'+i+'-衣服',ch.name,ch.p_better[2],ch.val_num_yifu,get_last_zero(ch.eq4_anylize[2])))
            eq_once_list.append(('四件套-'+i+'-鞋子',ch.name,ch.p_better[3],ch.val_num_xiezi,get_last_zero(ch.eq4_anylize[3])))
        eq_list.append(('两件套-'+ch.eq2,ch.name,eq2_better))

        eq_once_list.append(('两件套-' + ch.eq2+ '-球', ch.name, ch.p_better[4],ch.val_num_qiu,get_last_zero(ch.eq2_anylize[0])))
        eq_once_list.append(('两件套-' + ch.eq2+ '-绳子', ch.name, ch.p_better[5],ch.val_num_shengzi,get_last_zero(ch.eq2_anylize[1])))
        # print('球的分析',ch.val_num_qiu,ch.eq2_anylize[0])

    sort_eq_list=sorted(eq_list,key=lambda x:x[2])
    sort_eq_once_list=sorted(eq_once_list,key=lambda x:x[2])



    return sort_eq_list,sort_eq_once_list

#以装备为基础，进行角色提升概率的计算
def load_character_to_equipment(character_info_list):

    for ch in character_info_list:
        eq4_better=sum(ch.p_better[:4])/4
        eq2_better=sum(ch.p_better[4:])/2
        equip_match_character_dict[ch.eq2].append((ch.name,eq2_better))
        for i in ch.eq4:
            equip_match_character_dict[i].append((ch.name,eq4_better))



    equip_match_list=[]
    for key in equip_match_character_dict:
        if len(equip_match_character_dict[key])>0:
            # print(key,equip_match_character_dict[key])
            p_total=0
            name_total=''
            for name,p in equip_match_character_dict[key]:
                p_total+=p
                name_total+=name+'、'
            equip_match_list.append((key,p_total,name_total))
    sort_equip_match_list=sorted(equip_match_list,key=lambda x:x[1])

    pose_dict=dict()
    for pose,p,name in equip_match_list:
        # print(pose,equip_from_dict[pose])
        pf=equip_from_dict[pose]
        if pf in pose_dict:
            pose_dict[pf].append((pose,p,name))
        else:
            pose_dict[pf]=[(pose,p,name)]

    # for key in pose_dict:
    #     print(key,pose_dict[key])

    pose_list=[]
    for key in pose_dict:
        p_total=0
        name_total=''
        eq_name_total=''
        for eq,p,name in pose_dict[key]:
            p_total+=p
            name_total+=name
            eq_name_total+=eq+'+'
        pose_list.append((key,p_total,eq_name_total,name_total))

    sort_pose_list = sorted(pose_list,key=lambda x:x[1])




    return sort_equip_match_list,sort_pose_list

def main(show_log=False):
    excel_file_path = '星铁装备刷取.xlsx'

    #初始化主词条字典
    init_equip_main_dict()
    #初始化副词条字典
    init_equip_res_dict()

    #读取装备相关信息
    read_excel_equip_info(excel_file_path)

    # 读取Excel文件
    character_info_list=read_excel_character_info(excel_file_path)

    #分析角色装备获取的概率信息
    update_character_equip_anylize(character_info_list)

    #展示角色更好的装备刷取概率
    sort_eq_list,sort_eq_once_list=show_character_equipment_list(character_info_list)

    #将角色根据装备信息挂载到装备里面去
    sort_equip_match_list,sort_pose_list=load_character_to_equipment(character_info_list)

    if show_log:
        #展示角色套装刷取优化概率
        print('====================================展示角色套装刷取优化概率====================================')
        for i in sort_eq_list:
            print(i)

        #展示角色部位刷取优化概率
        print('====================================展示角色部位刷取优化概率====================================')
        for i in sort_eq_once_list:
            print(i)

        #展示套装刷取优化概率
        print('====================================展示套装刷取优化概率====================================')
        for i in sort_equip_match_list:
            print(i)

        #展示刷取地点刷取优化概率
        print('====================================展示刷取地点刷取优化概率====================================')
        for i in range(len(sort_pose_list)):
            print(i,sort_pose_list[i])

if __name__ == "__main__":
    # val_show_log=False
    val_show_log=True

    main(show_log=val_show_log)