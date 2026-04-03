import pandas as pd
import numpy as np

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

#转化角色信息
def exchang_character_info(c_info):
    c_name=c_info[0]


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
    print(name_list)
    print('--------------')
    print(attr_list)
    print('--------------')

    #枚举每个角色
    for i in range(len(character_list)):
        c_info=character_list[i]
        print(c_info)
        c_data=exchang_character_info(c_info)
        #将获取的角色信息添加至list
        ret_list.append(c_data)

    return ret_list


def main():
    excel_file_path = '星铁装备刷取.xlsx'

    # 读取Excel文件
    character_info_list=read_excel_character_info(excel_file_path)



if __name__ == "__main__":
    main()