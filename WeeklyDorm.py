'''
Description: 自律会生成宿舍卫生周总结
Author: Catop
Date: 2021-06-19 22:57:39
LastEditTime: 2021-06-19 23:32:06
'''
#encoding:utf-8

import sys
import re
import xlrd
import xlwt

########################################
#配置班级别名替换
CLASS_ALIAS = {
    "计算机科学与技术" : "计算机",
    "软件工程" : "软  件",
    "数字媒体技术" : "数  媒",
    "网络工程" : "网  络",
    "信息安全" : "信  安",
    "物联网工程" : "物联网",
    "计算机科学与技术-图灵班" : "图灵班",
    "智能科学与技术" : "智  科"

}
########################################


def read_xls(fileName):
    """读取卫检报表，输出结构化dict"""
    
    srcFile = xlrd.open_workbook(fileName)
    table = srcFile.sheets()[0]
    
    rowsNum = table.nrows     #行数
    #获取第一列原始数据(班级)
    class_raw = table.col(0,start_rowx=1)
    #结构化dict
    class_list = []
    
    #处理合并单元格，获取班级索引上下界
    for idx in range(0,len(class_raw)-1):
        #获取班级名称和上界
        if not (class_raw[idx].value == ''):
            class_info = {}
            class_info['class_name'] = class_raw[idx].value
            class_info['begin_idx'] = idx
            class_list.append(class_info)  
    for idx,item in enumerate(class_list):
        #获取下界
        if (idx == len(class_list)-1):
            #最后一个班级
            item['end_idx'] = len(class_raw)-1
        else:
            item['end_idx'] = class_list[idx+1]['begin_idx'] - 1
    

    #获取各班级宿舍信息
    dorm_raw = table.col(8,start_rowx=1)
    for class_info in class_list:
        dorm_cont = {"A":0, "B":0, "C":0, "D":0}
        begin_idx = class_info['begin_idx']
        end_idx = class_info['end_idx']
        

        #for i in range(begin_idx, end_idx-begin_idx):
        i = begin_idx
        while(i>=begin_idx and i<=end_idx):
            dorm_mark = dorm_raw[i].value
            if not(dorm_mark == ''):
                dorm_cont[dorm_mark] += 1
            i += 1
        class_info['dorm_cont'] = dorm_cont

    #print(class_list)

    return class_list



def write_xls(fileName, class_list):
    """写入周总结表格"""

    desFile = xlwt.Workbook(encoding = 'utf-8')
    worksheet = desFile.add_sheet('Sheet1')

    for idx, class_info in enumerate(class_list):
        #班级名称
        class_name = class_info['class_name']
        numRe = re.search("\d", class_name)
        if(numRe):
            numIdx = numRe.start()
            class_name_prefix = class_name[0:numIdx]
            class_name_num = class_name[numIdx:]
        else:
            class_name_prefix = class_name
            class_name_num = ""

        if(class_name_prefix in CLASS_ALIAS.keys()):
            class_name_prefix = CLASS_ALIAS[class_name_prefix]
        worksheet.write(idx, 0, label=class_name_prefix+class_name_num)

        #宿舍个数
        dorm_num = 0
        for k in class_info['dorm_cont'].keys():
            dorm_num += class_info['dorm_cont'][k]
        worksheet.write(idx, 1, label=dorm_num)

        #A等级
        worksheet.write(idx, 2, label=class_info['dorm_cont']['A'])
        #B等级
        worksheet.write(idx, 3, label=class_info['dorm_cont']['B'])
        #C等级
        worksheet.write(idx, 4, label=class_info['dorm_cont']['C'])
        #D等级
        worksheet.write(idx, 5, label=class_info['dorm_cont']['D'])
    
    desFile.save(fileName)



def work(srcFile, desFile):
    try:
        class_list = read_xls(srcFile)
        write_xls(desFile, class_list)
    except FileNotFoundError:
        print("[ERROR]找不到指定文件")
    except:
        print("[ERROR]未知错误，请自行排查")
    else:
        print("[INFO] Well done!")




if __name__ == "__main__":
    #work('./in.xls','./out.xls')
    if(len(sys.argv) == 3):
        work(sys.argv[1],sys.argv[2])
    else:
        print(f"[ERROR]命令行参数有误，使用示例：\npython3 {sys.argv[0]} in.xls out.xls")
