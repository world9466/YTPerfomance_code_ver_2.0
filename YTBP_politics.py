import pandas as pd
import os


########################  影片清單  ###########################

#從完整的影片清單逐項確認表格是否存在，挑出本月的節目
full_video_list = [
    '大新聞大爆卦',
    '週末大爆卦',
    '正常發揮',
    '國際直球對決',
    '前進戰略高地',
    '頭條開講'
    ]

video_list = []
for video in full_video_list:
    videofilepath = "video_table/table_{}.xlsx".format(video)
    if os.path.isfile(videofilepath):
        video_list.append(video)

########################  bottom  ############################



########################  通用函式  ##########################

# 抓取需要的欄位第一筆的統計資料
def table_ytbp(table):

    # 抓取需要的表格欄位
    table = table[
        ['youtuber','name','Your estimated ad revenue (USD)','Your YouTube Premium revenue (USD)',
        'Your transaction revenue (USD)','Your estimated revenue (USD)','Views','Watch time (hours)',
        'Subscribers','新進會員數','總影片數']
        ].head(1)

    # 更改欄位順序，測試功能用途，不啟用
    '''
    table = table.reindex(columns=['youtuber','name','Your estimated ad revenue (USD)',
    'Your YouTube Premium revenue (USD)','Your transaction revenue (USD)','Your estimated revenue (USD)',
    'Views','Watch time (hours)','Subscribers','新進會員數','總影片數'])
    '''

    # 特定欄位四捨五入
    table = table.round({
        'Your estimated ad revenue (USD)':0,'Your estimated revenue (USD)':0,
        'Your YouTube Premium revenue (USD)':1,
        'Your transaction revenue (USD)':1,'Watch time (hours)':0})
    return table

########################  bottom  ############################



########################  表格輸出  ###########################

ytbp = {
    'youtuber':[],'name':[],'Your estimated ad revenue (USD)':[],'Your YouTube Premium revenue (USD)':[],
    'Your transaction revenue (USD)':[],'Your estimated revenue (USD)':[],'Views':[],'Watch time (hours)':[],
    'Subscribers':[],'新進會員數':[],'總影片數':[]
    }
ytbp = pd.DataFrame(ytbp)

for video in video_list:
    table = pd.read_excel('video_table/table_{}.xlsx'.format(video))
    ytbp_1 = table_ytbp(table)
    ytbp = pd.concat([ytbp,ytbp_1],ignore_index=True)

ytbp = ytbp.sort_values(by = ['Your estimated revenue (USD)'],ascending = False)
ytbp = ytbp.reset_index(drop = True)
ytbp.index+=1


filepath = "輸出報表"
if os.path.isdir(filepath):
    print('directory CHECK OK')
else:
    os.mkdir("輸出報表")

print(ytbp)
ytbp.to_csv('輸出報表/table_ytb_performance_politics.csv',encoding = 'utf-8-sig')


########################  bottom  ############################


print('===============  ytb_performance successful  ===============')
