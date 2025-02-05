#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# File:    0109 待完成 邻区生成脚本.py
# Date:    2025/1/9 09:23
# Author:  HGZ

import xlwings as xw  # 导入xlwings模块
import datetime
import pandas as pd
import os
import time  # 使用 time.time()函数 计算代码执行时间
from geopy.distance import geodesic

"""
根据服务小区和邻小区类型设定邻区距离常量：
宏站_宏站：DISTANCE_OUTDOOR_OUTDOOR 3km
宏站_室分：DISTANCE_OUTDOOR_INDOOR 0.5km
室分_宏站：DISTANCE_INDOOR_OUTDOOR 0.5km
室分_室分：DISTANCE_INDOOR_INDOOR 0.1km
"""
DISTANCE_OUTDOOR_OUTDOOR = 3
DISTANCE_OUTDOOR_INDOOR = 0.5
DISTANCE_INDOOR_OUTDOOR = 0.5
DISTANCE_INDOOR_INDOOR = 0.1

"""
设置筛选邻区对数量常量
"""
COUNT_NEIGHBOR_CELL = 200

start_time = time.time()
current_datatime = datetime.datetime.now().strftime('%Y-%m-%d %H%M')  # 定义当前日期变量，后面加到生成文件名后面，例子：2023-10-19-0940

app = xw.App(visible=False, add_book=False)  # 启动excel程序
app.display_alerts = False  # 关闭用户提示
app.screen_updating = False  # 关闭屏幕刷新

SourceFilePath_cell_info = ''  # 中兴LTE邻区规划和脚本生成工具

# TargetFilePath_FDD = rf'Excel_CM_PLAN_FDD_RADIO_result_{current_datatime}邻区脚本.xlsx'
# TargetFilePath_TDD = rf'Excel_CM_PLAN_TDD_RADIO_result_{current_datatime}邻区脚本.xlsx'
# TargetFilePath_MIMO = rf'addneighborrelation_lte_zh_result_{current_datatime}邻区脚本.xlsx'

prefix_cell_info = '中兴LTE邻区规划和脚本生成工具'

# 根据文件名前缀 确定操作文件
for filename in os.listdir():
    if filename.startswith(prefix_cell_info):
        SourceFilePath_cell_info = filename  # 中兴LTE邻区规划和脚本生成工具

workbook_SourceFilePath_cell_info = app.books.open(SourceFilePath_cell_info)  # 打开来源工作簿文件

# 中兴LTE邻区规划和脚本生成工具
df_SourceFilePath_cell_info = workbook_SourceFilePath_cell_info.sheets['LTE现网小区信息'].range('A1').options(
    pd.DataFrame,
    header=1,
    index=False,
    expand='table').value

df_cell_info_plan = df_SourceFilePath_cell_info[df_SourceFilePath_cell_info['是否LTE规划邻区'] == '是']

# 待规划小区 和 全部小区进行交叉合并
df_cell_info_plan_SourceFilePath_merge = pd.merge(df_cell_info_plan, df_SourceFilePath_cell_info, how='cross')

"""
计算经纬度距离
定义两个点的经纬度，（纬度和经度）
point1 = (29.640702, 116.105341)
point2 = (29.64795, 116.09994)
计算两点之间的距离
distance = geodesic(point1, point2).km
"""
df_cell_info_plan_SourceFilePath_merge['距离'] = ''
for i in df_cell_info_plan_SourceFilePath_merge.index:
    point1 = (df_cell_info_plan_SourceFilePath_merge.loc[i, '天线纬度(小数)_x'],
              df_cell_info_plan_SourceFilePath_merge.loc[i, '天线经度(小数)_x'])
    point2 = (df_cell_info_plan_SourceFilePath_merge.loc[i, '天线纬度(小数)_y'],
              df_cell_info_plan_SourceFilePath_merge.loc[i, '天线经度(小数)_y'])
    df_cell_info_plan_SourceFilePath_merge.loc[i, '距离'] = geodesic(point1, point2).km  # 计算距离

"""
将 交叉合并 生成的df 按照 服务小区和邻区 类型进行分类
"""
df_cross = df_cell_info_plan_SourceFilePath_merge.sort_values(by='距离', ascending=True,
                                                              ignore_index=True)  # 按照距离进行升序排序，并且 重置索引
df_outdoor_outdoor = df_cross[(df_cross['类型（宏站或室分）_x'] == 0) & (df_cross['类型（宏站或室分）_y'] == 0) & (
        df_cross['距离'] <= DISTANCE_OUTDOOR_OUTDOOR)]
df_outdoor_indoor = df_cross[(df_cross['类型（宏站或室分）_x'] == 0) & (df_cross['类型（宏站或室分）_y'] == 1) & (
        df_cross['距离'] <= DISTANCE_OUTDOOR_INDOOR)]
df_indoor_outdoor = df_cross[(df_cross['类型（宏站或室分）_x'] == 1) & (df_cross['类型（宏站或室分）_y'] == 0) & (
        df_cross['距离'] <= DISTANCE_INDOOR_OUTDOOR)]
df_indoor_indoor = df_cross[(df_cross['类型（宏站或室分）_x'] == 1) & (df_cross['类型（宏站或室分）_y'] == 1) & (
        df_cross['距离'] <= DISTANCE_INDOOR_INDOOR)]
"""
对所有经过距离筛选的df 进行 concat
"""
df_distance_concat = pd.concat([df_outdoor_outdoor, df_outdoor_indoor, df_indoor_outdoor, df_indoor_indoor])

"""
将服务小区和邻区 CGI 一致的邻区对 删除 CGI_x CGI_y
"""
for i in df_distance_concat.index:
    if df_distance_concat.loc[i, 'CGI_x'] == df_distance_concat.loc[i, 'CGI_y']:
        df_distance_concat = df_distance_concat.drop(i)

"""
对服务小区 CGI_x 进行分组，获取 前 COUNT_NEIGHBOR_CELL = 200 行
"""
df_distance_concat = df_distance_concat.groupby('CGI_x').head(COUNT_NEIGHBOR_CELL).sort_values('距离', ignore_index=True)
# df_distance_concat = df_distance_concat.groupby('CGI_x').head(5)

# print(df_distance_concat.sort_values('距离', ignore_index=True))

"""
将 符合条件 邻区对 保存
"""
wb_cell_cell = app.books.add()
# sht2 = wb_cell_cell.sheets.add('Sheet2')
wb_cell_cell.sheets['Sheet1'].range('A1').options(index=False).value = df_distance_concat
# wb_cell_cell.sheets['Sheet2'].range('A1').options(index=False).value = df_outdoor_outdoor
wb_cell_cell.save('cell to cell 邻区对表_result.xlsx')
wb_cell_cell.close()


def cell_cell():
    """
    生成 邻区对
    :return: 
    """


def lte_lte_relation():
    """
    生成 4-4 邻区脚本
    :return:
    """


def lte_nr_relation():
    """
    生成 4-5 邻区脚本
    :return:
    """


def nr_lte_relation():
    """
    生成 5-4 邻区脚本
    :return:
    """


def nr_nr_relation():
    """
    生成 5-5 邻区脚本
    :return:
    """


workbook_SourceFilePath_cell_info.close()
app.quit()

print('执行完毕！')
print('代码执行时间：{}'.format(int(time.time() - start_time)), '秒')

