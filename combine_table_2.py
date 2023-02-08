import pandas as pd
import os,re

bangumi_list = [
    '中天車享家_朱朱哥來聊車',
    '世界越來越盧',
    '民間特偵組',
    '全球政經周報',
    '老Z調查線',
    '阿比妹妹',
    '與錢同行',
    '論文門開箱']

# 目前公式只能處理1000部以內的數量，預計2023下半年再做修正
for bangumi in bangumi_list:
    Folderpath = "頻道資訊/{}(2)".format(bangumi)
    if os.path.isdir(Folderpath):
        data_1 = pd.read_excel('video_table/table_{}(1).xlsx'.format(bangumi))
        data_2 = pd.read_excel('video_table/table_{}(2).xlsx'.format(bangumi))


        #############################################################
        # 地理人數計算

        country_list = ['Total','TW','US','CA','SG','MY','CN','MO','HK']
        count = []

        # 迭代計算
        for country in country_list:
            try:
                test = [(data_1['ViewsGeography'][data_1['Geography']== country]).reset_index(drop=True)][0][0] + 1
                # 上面沒跳錯，就表示條件篩選有值，否則就跳例外為0 
                peo_1 = [(data_1['ViewsGeography'][data_1['Geography']== country]).reset_index(drop=True)][0][0]
            except:
                peo_1 = 0

            try:
                test = [(data_2['ViewsGeography'][data_2['Geography']== country]).reset_index(drop=True)][0][0] + 1
                peo_2 = [(data_2['ViewsGeography'][data_2['Geography']== country]).reset_index(drop=True)][0][0]
            except:
                peo_2 = 0
            peo = peo_1 + peo_2
            count.append(peo)

        Geography = {'Geography':country_list,'ViewsGeography':count}
        Geography = pd.DataFrame(Geography)


        #############################################################
        # 廣告收益計算
        ad_revenue_1 = data_1['Your estimated ad revenue (USD)'][0]
        ad_revenue_2 = data_2['Your estimated ad revenue (USD)'][0]

        ad_revenue = {'Your estimated ad revenue (USD)':[int(ad_revenue_1 + ad_revenue_2)]}
        ad_revenue = pd.DataFrame(ad_revenue)


        #############################################################
        # Premium收益計算
        Premium_revenue_1 = data_1['Your YouTube Premium revenue (USD)'][0]
        Premium_revenue_2 = data_2['Your YouTube Premium revenue (USD)'][0]

        Premium_revenue = {'Your YouTube Premium revenue (USD)':[int(Premium_revenue_1 + Premium_revenue_2)]}
        Premium_revenue = pd.DataFrame(Premium_revenue)


        #############################################################
        # 訂閱數計算
        sub1 = data_1['Subscribers'][0]
        sub2 = data_2['Subscribers'][0]

        sub = {'Subscribers':[int(sub1 + sub2)]}
        sub = pd.DataFrame(sub)


        #############################################################
        # 觀看數計算
        view1 = data_1['Views'][0]
        view2 = data_2['Views'][0]

        view = {'Views':[int(view1 + view2)]}
        view = pd.DataFrame(view)


        #############################################################
        # 觀看時數計算
        hour1 = data_1['Watch time (hours)'][0]
        hour2 = data_2['Watch time (hours)'][0]

        hour = {'Watch time (hours)':[hour1 + hour2]}
        hour = pd.DataFrame(hour)


        #############################################################
        # 總收益計算
        total_revenue_1 = data_1['Your estimated revenue (USD)'][0]
        total_revenue_2 = data_2['Your estimated revenue (USD)'][0]

        total_revenue = {'Your estimated revenue (USD)':[total_revenue_1 + total_revenue_2]}
        total_revenue = pd.DataFrame(total_revenue)


        #############################################################
        # 有/無訂閱的比例計算
        viewsubs_list = ['Total','Not subscribed','Subscribed']
        count_2 = []

        # 迭代計算
        for viewsubs in viewsubs_list:
            try:
                test = [(data_1['ViewsSubs'][data_1['Subscription status']== viewsubs]).reset_index(drop=True)][0][0] + 1
                # 上面沒跳錯，就表示條件篩選有值，否則就跳例外為0 
                peo_1 = [(data_1['ViewsSubs'][data_1['Subscription status']== viewsubs]).reset_index(drop=True)][0][0]
            except:
                peo_1 = 0

            try:
                test = [(data_2['ViewsSubs'][data_2['Subscription status']== viewsubs]).reset_index(drop=True)][0][0] + 1
                peo_2 = [(data_2['ViewsSubs'][data_2['Subscription status']== viewsubs]).reset_index(drop=True)][0][0]
            except:
                peo_2 = 0
            peo = int(peo_1 + peo_2)
            count_2.append(peo)

        viewsubs = {'Subscription status':viewsubs_list,'ViewsSubs':count_2}
        viewsubs = pd.DataFrame(viewsubs)


        #############################################################
        # 交易收益計算
        trans1 = data_1['Your transaction revenue (USD)'][0]
        trans2 = data_2['Your transaction revenue (USD)'][0]

        trans = {'Your transaction revenue (USD)':[trans1 + trans2]}
        trans = pd.DataFrame(trans)


        #############################################################
        # 會員數計算
        member1 = data_1['新進會員數'][0]
        member2 = data_2['新進會員數'][0]

        member = {'新進會員數':[int(member1 + member2)]}
        member = pd.DataFrame(member)


        #############################################################
        # 年齡計算
        # 原本數值是比例，非人數，所以要先提出比例乘上觀看數變成人數再做總體比例計算

        age_list = ['18–24 years','25–34 years','35–44 years','45–54 years','55–64 years','65+ years']
        count_3 = []

        for age_range in age_list:
            if age_range in data_1['Viewer age'].values:
                age1 = [(data_1['Views (%)'][data_1['Viewer age'] == age_range]).reset_index(drop=True)][0][0]
                age1 = int(age1/100*view1)
                if age_range in data_2['Viewer age'].values:
                    age2 = [(data_2['Views (%)'][data_2['Viewer age'] == age_range]).reset_index(drop=True)][0][0]
                    age2 = int(age2/100*view2)
                    count_3.append(
                        round((age1 + age2)/(view1 + view2)*100,2)
                        )
                elif age_range not in data_2['Viewer age'].values:
                    count_3.append(
                        round(age1/(view1 + view2)*100,2)
                        )
            elif age_range not in data_1['Viewer age'].values:
                if age_range in data_2['Viewer age'].values:
                    age2 = [(data_2['Views (%)'][data_2['Viewer age'] == age_range]).reset_index(drop=True)][0][0]
                    age2 = int(age2/100*view2)
                    count_3.append(
                        round(age2/(view1 + view2)*100,2)
                        )
                else:
                    count_3.append(0)

        viewage = {'Viewer age':age_list,'Views (%)':count_3}
        viewage = pd.DataFrame(viewage)


        #############################################################
        # 性別的比例計算
        gender_list = ['Female','Male']
        count_4 = []

        for gender_range in gender_list:
            gender_1 = [(data_1['Views (%)gender'][data_1['Viewer gender'] == gender_range]).reset_index(drop=True)][0][0]
            gender_2 = [(data_2['Views (%)gender'][data_2['Viewer gender'] == gender_range]).reset_index(drop=True)][0][0]
            gender_1 = int(gender_1/100*view1)
            gender_2 = int(gender_2/100*view2)

            count_4.append(
                round((gender_1 + gender_2)/(view1 + view2)*100,2)
                )

        viewgender = {'Viewer gender':gender_list,'Views (%)gender':count_4}
        viewgender = pd.DataFrame(viewgender)


        #############################################################
        # 影片數量總計
        video_amount = int(data_1['總影片數'][0] + data_2['總影片數'][0])
        videos = {'總影片數':[video_amount]}
        videos = pd.DataFrame(videos)



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
        new_table = new_table.join(data_1[['name','youtuber']],how = 'outer',rsuffix='exc')

        new_table.loc[0,'name'] = bangumi

        new_table.to_excel('video_table/table_{}.xlsx'.format(bangumi))


print('===============  combine_step_2 successful  ===============')