import os
import json
import re
def Get_files(Path):
    Files_list = os.listdir(Path)
    Files_list.remove('script')
    return  Files_list

def handle_information(information):
    information = information.readlines()
    Item = {}
    title = ['中文名字：\n','英文名字：\n','CPU信息：\n','主板信息：\n','内存信息：\n','硬盘信息：\n','网卡MAC：\n','ip地址：\n','显卡信息：\n','显示器信息：\n']
    for Name in title:
        if (Name == "内存信息：\n"):
            Mem_Count_Index = information.index("内存数量：\n")
            Mem_Count = information[Mem_Count_Index+1]
            Mem_Count = int(Mem_Count)
            Mem_Inf_Index_Begin = Mem_Count_Index+2
            Mem_Inf_Index_End = Mem_Inf_Index_Begin+(Mem_Count*3)
            inf = information[Mem_Inf_Index_Begin:Mem_Inf_Index_End]
            Name = Name.replace("\n", "")
            inf = [x.strip() for x in inf if x.strip() != '']
            Item[Name] = inf
        elif (Name == "硬盘信息：\n"):
            Disk_Count_Index = information.index("硬盘数量：\n")
            Disk_Count = information[Disk_Count_Index + 1]
            Disk_Count = int(Disk_Count)
            Disk_Inf_Index_Begin = Disk_Count_Index + 2
            Disk_Inf_Index_End = Disk_Inf_Index_Begin + (Disk_Count * 3)
            inf = information[Disk_Inf_Index_Begin:Disk_Inf_Index_End]
            Name = Name.replace("\n", "")
            inf = [x.strip() for x in inf if x.strip() != '']
            Item[Name] = inf
        elif (Name == "显示器信息：\n"):
            Monitor_Count_Index = information.index("显示器数量：\n")
            try:
                Monitor_Count = information[Monitor_Count_Index + 1]
            except (IndexError) as e:
                print("没有显示器信息")
            else:
                Monitor_Count = information[Monitor_Count_Index + 1]
                Monitor_Count = int(Monitor_Count)
                Monitor_Inf_Index_Begin = Monitor_Count_Index + 2
                Monitor_Inf_Index_End = Monitor_Inf_Index_Begin + (Monitor_Count * 3)
                inf = information[Monitor_Inf_Index_Begin:Monitor_Inf_Index_End]
                Name = Name.replace("\n", "")
                inf = [x.strip() for x in inf if x.strip() != '']
                Item[Name] = inf
        else:
            i = information.index(Name)
            Name = Name.replace("\n","")
            Item[Name] = information[i+1].replace("\n","")
    return Item



if __name__ == '__main__':
    Path = 'E:/information/'
    Files_Name = Get_files(Path)
    Information ={}
    for i in Files_Name:
        File_Path = Path+i
        inf = open(File_Path,encoding='utf-16')
        datas = handle_information(inf)
        inf.close()
        print(datas)

