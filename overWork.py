# 统计加班时间
# Coded by Henry 6/12/2018

import xlrd #可读写*.xls 2003格式表格数据
import xlwt
import pandas as pd
import numpy as np
import os
import glob


def parse_overtime(week, on_h, on_m, off_h, off_m):
    # 解算加班时间 星期几 上班时间 下班时间
    std_week = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    std_work_time = 8
    over_time = 0   # 加班时间
    subsidy = False  # 是否有餐补
    
    if on_m/60 >= 0.5:
        # 半点后上班取大的小时数
        on_h += 1
        on_m = 0
    else:
        # 半点前取半点 
        on_m = 0.5
    if off_m/60 >= 0.5:
        # 半点后下班取半点
        off_m = 0.5
    else:
        # 半点前下班取小的小时数
        off_m = 0
    
    # 时间转换为数值
    on_work_time = on_h + on_m
    off_work_time = off_h + off_m
    
    # 当天工作总时间
    if on_work_time >= 13:
        # 13点后上班 不计算中午1小时休息时间
        over_time = off_work_time - on_work_time
    elif 13 > on_work_time > 12:
        # 12点~13点上班 按13点算
        over_time = off_work_time - 13
    elif on_work_time < 8.5:
        # 8点半前来单位不能计入上班时长
        over_time = off_work_time - 8.5 - 1
    else:
        # 12点前上班 算入中午1小时休息时间
        over_time = off_work_time - on_work_time - 1

    #　当天加班时长
    # 周六周日
    if week == std_week[5] or week == std_week[6]:
        # 周末超8小时才有餐补
        if over_time >= std_work_time:
            subsidy = True
        else:
            subsidy = False
    else:
        # 工作日
        if on_work_time > 9:
            # 工作日上班迟到 不影响计算加班时间
            over_time = off_work_time - 9 - 1 - std_work_time
        else:
            # 真实加班时间
            over_time -= std_work_time
        # 加够两个半小时有餐补
        if over_time >= 2.5:
            subsidy = True
        elif over_time < 1:
            # 加班不到1小时不算加班
            over_time = 0
            subsidy = False
        else:
            subsidy = False
        
        # 工作日加班起算时间（打印用）
        if on_work_time <= 8.5:
            on_work_time = 17.5
        else:
            on_work_time = 18

    # 格式化小时数、分钟数
    output_time_list = []
    for time in [on_work_time, off_work_time, over_time]:
        real_h = int(time)
        if real_h == 0:
            real_h = '0'
            real_m = '00'
        else:
            real_m = time % real_h
            if real_m >= 0.5:
                real_m = '30'
            else:
                real_m = '00'
        real_time = '%s:%s' % (real_h, real_m)
        output_time_list.append(real_time)
    
    # 输出：上班时间（格式化）、下班时间（格式化）、加班时间（格式化）、加班时间（数值）、是否有餐补
    return output_time_list[0], output_time_list[1], output_time_list[2], over_time, subsidy

    
def stat_overtime(xls_path):
    # xls_path = r'D:\公司材料\加班明细\0102-0110.xlsx'
    wb = xlrd.open_workbook(xls_path, encoding_override='gb2312')
    df = pd.read_excel(wb, engine='xlrd')
    # 获取部门员工姓名
    name_dic = {}
    for name in df['姓名']:
        if name not in name_dic:
            name_dic[name] = 1
    total_overtime = 0  # 一个xls所有人的加班时间总和
    # 遍历每个人的打卡时间
    for name in name_dic.keys():
        sr1 = df[df['姓名']==name]
        print(name)
        #　打卡时间序列
        dwt_ser = sr1['日期时间']
        dwt_dict = {}
        for dwt in dwt_ser:
            date = dwt.split(' ')[0]
            week = dwt.split(' ')[1]
            time = dwt.split(' ')[2]
            hour = int(time.split(':')[0])
            minute = int(time.split(':')[1])
            if date not in dwt_dict:
                dwt_dict[date] = [[week, hour, minute]]
            else:
                dwt_dict[date].append([week, hour, minute])
        # 遍历一天的
        sum_over_time = 0   # 共计多少小时
        for key in dwt_dict.keys():
            on_work_mark = dwt_dict[key][0]
            off_work_mark = dwt_dict[key][-1]
            week = on_work_mark[0]
            on_work_hour = on_work_mark[1]
            on_work_min = on_work_mark[2]
            off_work_hour = off_work_mark[1]
            off_work_min = off_work_mark[2]
            on_time, off_time, over_time, over_time_scalar, subsidy = parse_overtime(week, on_work_hour, on_work_min, off_work_hour, off_work_min)
            if over_time_scalar >= 1:
                sum_over_time += over_time_scalar
                print('%s\t%s\t%s\t%s' % (key, on_time, off_time, over_time))
            name_dic[name] = sum_over_time
        total_overtime += sum_over_time
        print('共计%.1f小时' % sum_over_time)
    # 最终返回该文件中所有人加班时间总和 以及 每个人的加班时间统计字典
    return total_overtime, name_dic


if __name__ == '__main__':
    '''
    sum_time = 0    # 所有文件所有人加班时间总和
    stat_per_person = {}    #　部门人员加班统计
    for xls in glob.glob('./*.xls*'):
        overtime_per_xls, stat_per_xls = stat_overtime(xls)
        sum_time += overtime_per_xls
        for person in stat_per_xls.keys():
            if person in stat_per_person:
                stat_per_person[person] += stat_per_xls[person]
            else:
                stat_per_person[person] = 0
    print(sum_time, stat_per_person)
    
    for xls in glob.glob('./*-07*.xls*'):
        stat_overtime(xls)
    '''
    # xls = r'./1210-1216.xls'
    xls = r'D:\公司材料\加班明细\overwork_xls\0326-0331.xls'
    stat_overtime(xls)