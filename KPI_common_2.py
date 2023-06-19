import pandas as pd
import os

###  節目排行榜  ###
########################  影片清單  ###########################

#從完整的影片清單逐項確認表格是否存在，挑出本月的節目
full_video_list = [
    '小麥的健康筆記',
    '小豪出任務',
    '中天車享家_朱朱哥來聊車',
    '世界越來越盧',
    '民間特偵組',
    '全球政經周報',
    '老Z調查線',
    '你的豪朋友',
    '宏色封鎖線_宏色禁區',
    '金牌特派',
    '阿比妹妹',
    '洪流洞見',
    '食安趨勢報告',
    '真心話大冒險',
    '愛吃星球',
    '新聞千里馬',
    '新聞龍捲風',
    '詩瑋愛健康',
    '詭案橞客室',
    '嗶!就是要有錢',
    '窩星球',
    '綠也掀桌',
    '與錢同行',
    '論文門開箱',
    '鄭妹看世界',
    '螃蟹秀開鍘',
    '獸身男女',
    '靈異錯別字_鬼錯字',
    '琴謙天下事',
    '誰謀殺了言論自由',
    '線上面對面',
    '我是二寶爸',
    '高級酸新聞台',
    '財經風向球',
    '健康點點名'
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
table.to_csv('輸出報表/table_KPI_common_2.csv',encoding = 'utf-8-sig')



##############################################################

print('===============  KPI successful  ===============')