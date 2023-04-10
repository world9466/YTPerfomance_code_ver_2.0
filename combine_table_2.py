import pandas as pd
import os,re

bangumi_list = [
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
    '誰謀殺了言論自由_俠姊',
    '誰謀殺了言論自由_金汝鑫',
    '誰謀殺了言論自由_賴麗櫻',
    '大新聞大爆卦',
    '週末大爆卦',
    '頭條開講',
    '正常發揮',
    '前進戰略高地',
    '國際直球對決',
    '線上面對面',
    '我是二寶爸',
    '高級酸新聞台',
    '財經風向球'
    ]

folder_list=[20,19,18,17,16,15,14,13,12,11,10,9,8,7,6,5,4,3,2]   #20個table可以涵蓋1萬部影片，應該夠用了

for bangumi in bangumi_list:
    # 確認有沒有兩個以上要計算，沒有的話就進入下個迴圈
    path_check = "video_table/table_{}(2).xlsx".format(bangumi)
    if os.path.isfile(path_check):
        print('開始計算節目:"{}"的數據'.format(bangumi))
        pass
    else:
        continue    

    # 確認有幾個table要做加總計算
    for number in folder_list:
        filepath = "video_table/table_{}({}).xlsx".format(bangumi,number)
        if os.path.isfile(filepath):
            total = number
            print('確認節目："'+bangumi+'"有'+str(total)+'個table')
            break

    # 開始讀檔進行加總
    #############################################################
    # 地理人數計算

    country_list = ['Total','TW','US','CA','SG','MY','CN','MO','HK']
    count = [0,0,0,0,0,0,0,0,0]
    Geography = {'Geography':country_list,'ViewsGeography':count}
    Geography = pd.DataFrame(Geography)
    
    # 按順序迭代全部的table做相加
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        count=0
        for country in country_list:
            try:
                test = [(data['ViewsGeography'][data['Geography']== country]).reset_index(drop=True)][0][0] + 1
                # 上面沒跳錯，就表示條件篩選有值，否則就跳例外，值為0 
                peo = [(data['ViewsGeography'][data['Geography']== country]).reset_index(drop=True)][0][0]
            except:
                peo = 0
            Geography.at[count,'ViewsGeography']+=peo
            count+=1
    
    #############################################################
    # 廣告收益計算
    ad_revenue = {'Your estimated ad revenue (USD)':[0]}
    ad_revenue = pd.DataFrame(ad_revenue)
    ad_revenue_count = 0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        ad_revenue_count+=data['Your estimated ad revenue (USD)'][0]

    ad_revenue.at[0,'Your estimated ad revenue (USD)']+=round(ad_revenue_count,0)

    #############################################################
    # Premium收益計算
    Premium_revenue = {'Your YouTube Premium revenue (USD)':[0]}
    Premium_revenue = pd.DataFrame(Premium_revenue)
    Premium_revenue_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        Premium_revenue_count+=data['Your YouTube Premium revenue (USD)'][0]

    Premium_revenue.at[0,'Your YouTube Premium revenue (USD)']+=round(Premium_revenue_count,0)

    #############################################################
    # 訂閱數計算
    sub = {'Subscribers':[0]}
    sub = pd.DataFrame(sub)
    sub_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        sub_count+=data['Subscribers'][0]

    sub.at[0,'Subscribers']+=round(sub_count,0)

    #############################################################
    # 觀看數計算
    view = {'Views':[0]}
    view = pd.DataFrame(view)
    view_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        view_count+=data['Views'][0]

    view.at[0,'Views']+=round(view_count,0)

    #############################################################
    # 觀看時數計算
    hour = {'Watch time (hours)':[0]}
    hour = pd.DataFrame(hour)
    hour_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        hour_count+=data['Watch time (hours)'][0]

    hour.at[0,'Watch time (hours)']+=round(hour_count,2)

    #############################################################
    # 總收益計算
    total_revenue = {'Your estimated revenue (USD)':[0]}
    total_revenue = pd.DataFrame(total_revenue)
    total_revenue_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        total_revenue_count+=data['Your estimated revenue (USD)'][0]

    total_revenue.at[0,'Your estimated revenue (USD)']+=round(total_revenue_count,2)

    #############################################################
    # 有/無訂閱的比例計算
    substates_list = ['Total','Not subscribed','Subscribed']
    count_2 = [0,0,0]
    viewsubs = {'Subscription status':substates_list,'ViewsSubs':count_2}
    viewsubs = pd.DataFrame(viewsubs)

    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        count=0
        for states in substates_list:
            try:
                test = [(data['ViewsSubs'][data['Subscription status']== states]).reset_index(drop=True)][0][0] + 1
                # 上面沒跳錯，就表示條件篩選有值，否則就跳例外，值為0 
                peo = [(data['ViewsSubs'][data['Subscription status']== states]).reset_index(drop=True)][0][0]
            except:
                peo = 0
            viewsubs.at[count,'ViewsSubs']+=peo
            count+=1
    
    #############################################################
    # 交易收益計算
    trans = {'Your transaction revenue (USD)':[0]}
    trans = pd.DataFrame(trans)
    trans_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        trans_count+=data['Your transaction revenue (USD)'][0]

    trans.at[0,'Your transaction revenue (USD)']+=round(trans_count,2)

    #############################################################
    # 會員數計算
    member = {'新進會員數':[int(0)]}
    member = pd.DataFrame(member)
    member_count=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        member_count+=data['新進會員數'][0]

    member.at[0,'新進會員數']+=round(member_count,2)    

    #############################################################
    # 年齡計算
    # 原本數值是比例，非人數，所以要先提出比例乘上觀看數變成人數再做總體比例計算

    age_list = ['13–17 years','18–24 years','25–34 years','35–44 years','45–54 years','55–64 years','65+ years']
    count_3 = [0,0,0,0,0,0,0]
    viewage = {'Viewer age':age_list,'Views (%)':count_3}
    viewage = pd.DataFrame(viewage)

    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        count=0
        for age_range in age_list:
            if age_range in data['Viewer age'].values:
                peo_age = [(data['Views (%)'][data['Viewer age'] == age_range]).reset_index(drop=True)][0][0]
                peo_age = round(data['Views'][0]*peo_age,0)
                viewage.at[count,'Views (%)']+=peo_age
            count+=1

    viewage['Views (%)'] = round(viewage['Views (%)']/view_count,2)

    #############################################################
    # 性別的比例計算
    gender_list = ['Female','Male']
    count_4 = [0,0]
    viewgender = {'Viewer gender':gender_list,'Views (%)gender':count_4}
    viewgender = pd.DataFrame(viewgender)

    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        count=0
        for gender in gender_list:
            if gender in data['Viewer gender'].values:
                peo_gender = [(data['Views (%)gender'][data['Viewer gender'] == gender]).reset_index(drop=True)][0][0]
                peo_gender = round(data['Views'][0]*peo_gender,0)
                viewgender.at[count,'Views (%)gender']+=peo_gender
            count+=1

    viewgender['Views (%)gender'] = round(viewgender['Views (%)gender']/view_count,2)

    #############################################################
    # 影片數量總計
    videos = {'總影片數':[0]}
    videos = pd.DataFrame(videos)
    video_amount=0
    for num in range(total):
        data = pd.read_excel('video_table/table_{}({}).xlsx'.format(bangumi,num+1))
        video_amount+=data['總影片數'][0]
    videos.at[0,'總影片數']+=video_amount

    #############################################################
    # 合併表格輸出

    new_table = pd.DataFrame()

    new_table = new_table.join(Geography,how = 'outer',rsuffix='exc')
    new_table = new_table.join(ad_revenue,how = 'outer',rsuffix='exc')
    new_table = new_table.join(Premium_revenue,how = 'outer',rsuffix='exc')
    new_table = new_table.join(sub,how = 'outer',rsuffix='exc')
    new_table = new_table.join(view,how = 'outer',rsuffix='exc')
    new_table = new_table.join(hour,how = 'outer',rsuffix='exc')
    new_table = new_table.join(total_revenue,how = 'outer',rsuffix='exc')
    new_table = new_table.join(viewsubs,how = 'outer',rsuffix='exc')
    new_table = new_table.join(trans,how = 'outer',rsuffix='exc')
    new_table = new_table.join(member,how = 'outer',rsuffix='exc')
    new_table = new_table.join(viewage,how = 'outer',rsuffix='exc')
    new_table = new_table.join(viewgender,how = 'outer',rsuffix='exc')
    new_table = new_table.join(videos,how = 'outer',rsuffix='exc')

    data = pd.read_excel('video_table/table_{}(1).xlsx'.format(bangumi))
    new_table = new_table.join(data[['name','youtuber']],how = 'outer',rsuffix='exc')

    new_table.loc[0,'name'] = bangumi

    new_table.to_excel('video_table/table_{}.xlsx'.format(bangumi))



print('===============  combine_step_2 successful  ===============')

