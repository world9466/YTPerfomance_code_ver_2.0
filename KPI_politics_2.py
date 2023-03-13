import pandas as pd
import os

###  節目排行榜  ###
########################  影片清單  ###########################

#從完整的影片清單逐項確認表格是否存在，挑出本月的節目
full_video_list = [
    '正常發揮',
    '國際直球對決',
    '前進戰略高地',
    '大新聞大爆卦',
    '週末大爆卦'
    ]

video_list = []
for video in full_video_list:
    videofilepath = "video_table/table_{}.xlsx".format(video)
    if os.path.isfile(videofilepath):
        video_list.append(video)


########################  bottom  ############################



########################  表格生成  ###########################


trafficKPI = pd.DataFrame()
subsKPI = pd.DataFrame()
tranKPI = pd.DataFrame()


for video in video_list:
    table = pd.read_excel('video_table/table_{}.xlsx'.format(video))

    #先讀取所有表格的流量，然後彙整成表格，下面就可以做排名計算
    traffic_1 = table[['name','Views']].head(1)
    trafficKPI = pd.concat([trafficKPI,traffic_1],ignore_index=True)

    #先讀取所有表格的訂閱，然後彙整成表格，下面就可以做排名計算
    subs_1 = table[['name','Subscribers']].head(1)
    subsKPI = pd.concat([subsKPI,subs_1],ignore_index=True)

    #先讀取所有表格的斗內，然後彙整成表格，下面就可以做排名計算
    tran_1 = table[['name','Your transaction revenue (USD)']].head(1)
    tranKPI = pd.concat([tranKPI,tran_1],ignore_index=True)



######流量排行######
trafficKPI = trafficKPI.sort_values(by = ['Views'],ascending = False)
trafficKPI = trafficKPI.reset_index(drop = True)
trafficKPI.index+=1
trafficKPI = trafficKPI.head(5)
print(trafficKPI)



######訂閱排行######
subsKPI = subsKPI.sort_values(by = ['Subscribers'],ascending = False)
subsKPI = subsKPI.reset_index(drop = True)
subsKPI.index+=1
subsKPI = subsKPI.head(5)
print(subsKPI)



######收益排行######
tranKPI = tranKPI.sort_values(by = ['Your transaction revenue (USD)'],ascending = False)
tranKPI = tranKPI.reset_index(drop = True)
tranKPI.index+=1
tranKPI = tranKPI.head(5)
print(tranKPI)


########################  bottom  ############################



##########################KPI輸出#############################

filepath = "輸出報表"
if os.path.isdir(filepath):
    print('directory CHECK OK')
else:
    os.mkdir("輸出報表")

table = trafficKPI.join(subsKPI,rsuffix = 'subs')
table = table.join(tranKPI,rsuffix = 'tran')
print(table)
table.to_csv('輸出報表/table_KPI_politics_2.csv',encoding = 'utf-8-sig')



##############################################################

print('===============  KPI successful  ===============')