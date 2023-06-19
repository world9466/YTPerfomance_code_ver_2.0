import pandas as pd
import os

###  個人排行榜  ###
########################  影片清單  ###########################

#從完整的影片清單逐項確認表格是否存在，挑出本月的節目
full_video_list = [
    '正常發揮',
    '國際直球對決',
    '前進戰略高地',
    '大新聞大爆卦',
    '週末大爆卦',
    '頭條開講'
    ]

video_list = []
for video in full_video_list:
    videofilepath = "video_table/table_{}.xlsx".format(video)
    if os.path.isfile(videofilepath):
        video_list.append(video)

youtuber_name = ["麥玉潔","簡至豪","朱顯名","盧秀芳","賴麗櫻","李曉玲","周寬展",
"蔡秉宏","金汝鑫","林慧君","洪淑芬","譚若誼","鄭亦真","張若妤","馬千惠","戴立綱",
"方詩瑋","何橞瑢","畢倩涵","李珮瑄","張雅婷","劉盈秀","賴正鎧",'周玉琴','俠姊',
'丹申廣','林嘉源','張介凡','王宜安']

########################  bottom  ############################



########################  表格生成  ###########################

trafficKPI = {'youtuber':[],'Views':[]}
trafficKPI=  pd.DataFrame(trafficKPI)
subsKPI = {'youtuber':[],'Subscribers':[]}
subsKPI=  pd.DataFrame(subsKPI)
tranKPI = {'youtuber':[],'Your transaction revenue (USD)':[]}
tranKPI=  pd.DataFrame(tranKPI)

for video in video_list:
    table = pd.read_excel('video_table/table_{}.xlsx'.format(video))

    #先讀取所有表格的流量，然後彙整成表格，之後下方再迭代這個表格去針對相同名字作加總
    traffic_1 = table[['youtuber','Views']].head(1)
    trafficKPI = pd.concat([trafficKPI,traffic_1],ignore_index=True)

    #先讀取所有表格的訂閱，然後彙整成表格，之後下方再迭代這個表格去針對相同名字作加總
    subs_1 = table[['youtuber','Subscribers']].head(1)
    subsKPI = pd.concat([subsKPI,subs_1],ignore_index=True)

    #先讀取所有表格的斗內，然後彙整成表格，之後下方再迭代這個表格去針對相同名字作加總
    tran_1 = table[['youtuber','Your transaction revenue (USD)']].head(1)
    tranKPI = pd.concat([tranKPI,tran_1],ignore_index=True)



######流量排行######
trafficKPI_finish = {
    'youtuber':["麥玉潔","簡至豪","朱顯名","盧秀芳","賴麗櫻",
    "李曉玲","周寬展","蔡秉宏","金汝鑫","林慧君","洪淑芬",
    "譚若誼","鄭亦真","張若妤","馬千惠","戴立綱","方詩瑋",
    "何橞瑢","畢倩涵","李珮瑄","張雅婷","劉盈秀","賴正鎧",
    '周玉琴','俠姊','丹申廣','林嘉源','張介凡','王宜安'],
    'Views':[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    }
trafficKPI_finish = pd.DataFrame(trafficKPI_finish)

# 因為在統計資料上如果為不同節目但同一主持人，不同節目的數值要加在一起，因為排行榜是以人來排
# 迭代 trafficKPI 的主持人欄位，再迭代主持人總表，如果有符合的話就把trafficKPI_finish裡面的數值做更新(從0開始加上去)
# num 的作法是取在主持人名字相同的條件取得當下的index，然後因為index取出後是一個表格的形式，再用[0][0]取出在0欄位下的第0個值
for traffic in trafficKPI.values:
    for name in youtuber_name:
        if traffic[0] == name:
            num = trafficKPI_finish.index[trafficKPI_finish['youtuber']==name]  # num 是主持人在這個table的編號
            num = pd.DataFrame(num)
            trafficKPI_finish.at[num[0][0],'Views'] = trafficKPI_finish.at[num[0][0],'Views'] + traffic[1]

trafficKPI_finish = trafficKPI_finish.sort_values(by = ['Views'],ascending = False)
trafficKPI_finish = trafficKPI_finish.reset_index(drop = True)
trafficKPI_finish.index+=1
trafficKPI_finish = trafficKPI_finish.head(5)
print(trafficKPI_finish)



######訂閱排行######
subsKPI_finish = {
    'youtuber':["麥玉潔","簡至豪","朱顯名","盧秀芳","賴麗櫻","李曉玲","周寬展",
    "蔡秉宏","金汝鑫","林慧君","洪淑芬","譚若誼","鄭亦真","張若妤","馬千惠",
    "戴立綱","方詩瑋","何橞瑢","畢倩涵","李珮瑄","張雅婷","劉盈秀","賴正鎧",
    '周玉琴','俠姊','丹申廣','林嘉源','張介凡','王宜安'],
    'Subscribers':[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    }
subsKPI_finish = pd.DataFrame(subsKPI_finish)

# 同樣採迭代方式
for subs in subsKPI.values:
    for name in youtuber_name:
        if subs[0] == name:
            num = subsKPI_finish.index[subsKPI_finish['youtuber']==name]
            num = pd.DataFrame(num)
            subsKPI_finish.at[num[0][0],'Subscribers'] = subsKPI_finish.at[num[0][0],'Subscribers'] + subs[1]

subsKPI_finish = subsKPI_finish.sort_values(by = ['Subscribers'],ascending = False)
subsKPI_finish = subsKPI_finish.reset_index(drop = True)
subsKPI_finish.index+=1
subsKPI_finish = subsKPI_finish.head(5)
print(subsKPI_finish)



######收益排行######
tranKPI_finish = {
    'youtuber':["麥玉潔","簡至豪","朱顯名","盧秀芳","賴麗櫻","李曉玲","周寬展","蔡秉宏",
    "金汝鑫","林慧君","洪淑芬","譚若誼","鄭亦真","張若妤","馬千惠","戴立綱","方詩瑋",
    "何橞瑢","畢倩涵","李珮瑄","張雅婷","劉盈秀","賴正鎧",'周玉琴','俠姊','丹申廣',
    '林嘉源','張介凡','王宜安']
    ,'Your transaction revenue (USD)':[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    }
tranKPI_finish = pd.DataFrame(tranKPI_finish)

for tran in tranKPI.values:
    for name in youtuber_name:
        if tran[0] == name:
            num = tranKPI_finish.index[tranKPI_finish['youtuber']==name]
            num = pd.DataFrame(num)
            tranKPI_finish.at[num[0][0],'Your transaction revenue (USD)'] = tranKPI_finish.at[num[0][0],'Your transaction revenue (USD)'] + tran[1]


tranKPI_finish = tranKPI_finish.sort_values(by = ['Your transaction revenue (USD)'],ascending = False)
tranKPI_finish = tranKPI_finish.reset_index(drop = True)
tranKPI_finish.index+=1
tranKPI_finish = tranKPI_finish.head(5)
print(tranKPI_finish)


########################  bottom  ############################



##########################KPI輸出#############################

filepath = "輸出報表"
if os.path.isdir(filepath):
    print('directory CHECK OK')
else:
    os.mkdir("輸出報表")

table = trafficKPI_finish.join(subsKPI_finish,rsuffix = 'subs')
table = table.join(tranKPI_finish,rsuffix = 'tran')
print(table)
table.to_csv('輸出報表/table_KPI_politics.csv',encoding = 'utf-8-sig')



##############################################################

print('===============  KPI successful  ===============')

