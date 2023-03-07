import pandas as pd
import os,zipfile,re

########################  通用函式  ##########################

# 解壓縮所有檔案，如果壓縮檔存在就解壓縮到指定目錄，extractall會自動覆蓋原有檔案
def file_extract(bangumi):
    dirpath = '頻道資訊/{}/壓縮檔'.format(bangumi)
    if os.path.isdir(dirpath):
        files = os.listdir(dirpath)
        for file in files:
            if re.match('Channel.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Channel'.format(bangumi))
            elif re.match('Content.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Content'.format(bangumi))
            elif re.match('Geography.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Geography'.format(bangumi))
            elif re.match('Subscription status.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Subscription status'.format(bangumi))
            elif re.match('Transaction type.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Transaction type'.format(bangumi))
            elif re.match('Viewer age.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Viewer age'.format(bangumi))
            elif re.match('Viewer gender.*.zip',file):
                zf = zipfile.ZipFile('頻道資訊/{}/壓縮檔/{}'.format(bangumi,file), 'r')
                zf.extractall(path = '頻道資訊/{}/Viewer gender'.format(bangumi))



# 同樣的表格格式，載入不同的節目，變數輸入節目名稱，用以更改路徑
def bangumi_load(bangumi):
    global tablechannel
    global tableGeography
    global tableSubs
    global tableTran
    global tableage
    global tablegender
    global tablecontent
    tablechannel = pd.read_csv('頻道資訊/{}/Channel/Table data.csv'.format(bangumi),encoding = 'utf8')
    # 版更後的欄位名稱修正回原本的
    tablechannel = tablechannel.rename(columns={
        'Watch Page ads (USD)': 'Your estimated ad revenue (USD)',
        'YouTube Premium (USD)':'Your YouTube Premium revenue (USD)',
        'Estimated revenue (USD)':'Your estimated revenue (USD)'
        })

    tableGeography = pd.read_csv('頻道資訊/{}/Geography/Table data.csv'.format(bangumi),encoding = 'utf8')
    tableSubs = pd.read_csv('頻道資訊/{}/Subscription status/Table data.csv'.format(bangumi),encoding = 'utf8')

    #如果收益為0，沒有下載csv檔，就自動創建一個收益值為0的csv檔
    tranfilepath = "頻道資訊/{}/Transaction type/Table data.csv".format(bangumi)
    if os.path.isfile(tranfilepath):
        print("tableTran directory CHECK OK")
        tableTran = pd.read_csv('頻道資訊/{}/Transaction type/Table data.csv'.format(bangumi),encoding = 'utf8')
        # YT後台版本更新，欄位名稱有變，改回原本的
        tableTran = tableTran.rename(columns={'Fan funding (USD)': 'Your transaction revenue (USD)'})
    else:
        tranfilepath_1 = "頻道資訊/{}/Transaction type".format(bangumi)
        if os.path.isdir(tranfilepath_1):
            print('tableTran directory CHECK OK')
        else:
            os.mkdir("頻道資訊/{}/Transaction type".format(bangumi))
            print('create folder')
        tran_data = {'Transaction type':['Total'],'Transactions':[0],'Your transaction revenue (USD)':[0],'新進會員數':[0]}
        tran_data = pd.DataFrame(tran_data)
        tran_data.to_csv('頻道資訊/{}/Transaction type/Table data.csv'.format(bangumi),encoding = 'utf-8-sig')
        tableTran = pd.read_csv('頻道資訊/{}/Transaction type/Table data.csv'.format(bangumi),encoding = 'utf8')

    # 年齡分層沒有資料就做一個空白table
    agefilepath = "頻道資訊/{}/Viewer age/Table data.csv".format(bangumi)
    if os.path.isfile(agefilepath):
        print("tableage directory CHECK OK")
        tableage = pd.read_csv('頻道資訊/{}/Viewer age/Table data.csv'.format(bangumi),encoding = 'utf8')
    else:
        ages = {'Viewer age':[],'Views (%)':[]}
        tableage = pd.DataFrame(ages)

    # 性別沒有資料，就做一個空值table
    genderfilepath = "頻道資訊/{}/Viewer gender/Table data.csv".format(bangumi)
    if os.path.isfile(genderfilepath):
        print("tablegender directory CHECK OK")
        tablegender = pd.read_csv('頻道資訊/{}/Viewer gender/Table data.csv'.format(bangumi),encoding = 'utf8')
    else:
        print("tablegender directory False")
        genders = {'Viewer gender':['Female','Male'],'Views (%)':[0,0]}
        tablegender = pd.DataFrame(genders)


    tablecontent = pd.read_csv('頻道資訊/{}/Content/Table data.csv'.format(bangumi),encoding = 'utf8')


# 從交易收入裡的頻道會員抓取新會員數，做成表格，再併到 tableTran 表內
def tran_member():
    global tableTran
    if 'Channel Memberships' in tableTran['Transaction type'].values:
        #條件篩選出來的member_x 是一個 list 的狀態，不能直接取用  必須先取值再以dict形式做成 Dataframe
        member_x = tableTran['Transactions'][tableTran['Transaction type']=='Channel Memberships']        
        member_x = pd.DataFrame(member_x)        
        member_x = member_x.reset_index()
        member = {"新進會員數":[member_x['Transactions'][0]]}
        member = pd.DataFrame(member)
        tableTran = tableTran.join(member,how='outer',rsuffix = 'tran')
    else:
        member = {"新進會員數":[0]}
        member = pd.DataFrame(member)
        tableTran = tableTran.join(member,how='outer',rsuffix = 'tran')
    return tableTran


# 從各個csv檔抓取報表需要的欄位
def table_combine():
    table1 = tableGeography[['Geography','Views']].join(tablechannel,how='outer',lsuffix = 'Geography')

    table2 = table1.join(tableSubs[['Subscription status','Views']],how = 'outer',rsuffix = 'Subs')

    table2 = table2.join(tableTran[['Transaction type','Your transaction revenue (USD)','Transactions','新進會員數']],how='outer',rsuffix = 'Tran')

    table2 = table2.join(tableage[['Viewer age','Views (%)']],how='outer',rsuffix = 'age')

    table = table2.join(tablegender[['Viewer gender','Views (%)']],how='outer',rsuffix = 'gender')

    #加入總影片數，超過1筆的時候會有total值1筆，要扣掉
    videos = len(tablecontent.index)
    if videos == 1:
        videomount = {'總影片數':[1]}
    elif videos == 502:
        videomount = {'總影片數':[500]}
    else:
        videomount = {'總影片數':[videos-1]}
    videomount = pd.DataFrame(videomount)
    table = table.join(videomount,how='outer',rsuffix = 'video')

    return table 
    

# 生成各個節目資訊表(其他函式都包在這裡面)
def table_gen(bangumi,youtuber):
    global tableTran
    bangumi_load(bangumi)
    tableTran = tran_member()
    table = table_combine()
    channel_name = {
        "name":[bangumi],
        "youtuber":[youtuber]
    }
    channel_name = pd.DataFrame(channel_name)
    table = table.join(channel_name,how='outer',rsuffix = 'name')
    table.to_excel('video_table/table_{}.xlsx'.format(bangumi))



########################  bottom  ############################





########################  輸出報表  ##########################


# 建立輸出資料夾
filepath = "video_table"
if os.path.isdir(filepath):
    print('directory CHECK OK')
else:
    os.mkdir("video_table")


# 配對頻道跟youtuber，已經或可能超過500部的節目的資料夾都要找過一遍
channel_youtuber_list = [
    ('小麥的健康筆記','麥玉潔'),
    ('小豪出任務','簡至豪'),
    ('中天車享家_朱朱哥來聊車','朱顯名'),
    ('世界越來越盧','盧秀芳'),
    ('世界越來越盧','盧秀芳'),
    ('民間特偵組','賴麗櫻'),
    ('民間特偵組','賴麗櫻'),
    ('全球政經周報','李曉玲'),
    ('全球政經周報','李曉玲'),
    ('老Z調查線','周寬展'),
    ('老Z調查線','周寬展'),
    ('你的豪朋友','簡至豪'),
    ('宏色封鎖線_宏色禁區','蔡秉宏'),
    ('中天電視','蔡秉宏'),
    ('金牌特派','金汝鑫'),
    ('阿比妹妹','林慧君'),
    ('阿比妹妹','林慧君'),
    ('政治新人榜','李珮瑄'),
    ('洪流洞見','洪淑芬'),
    ('流行星球','譚若誼'),
    ('食安趨勢報告','賴麗櫻'),
    ('真心話大冒險','鄭亦真'),
    ('新聞千里馬','馬千惠'),
    ('新聞龍捲風','戴立綱'),
    ('詩瑋愛健康','方詩瑋'),
    ('詭案橞客室','何橞瑢'),
    ('嗶!就是要有錢','畢倩涵'),
    ('窩星球','何橞瑢'),
    ('綠也掀桌','李珮瑄'),
    ('與錢同行','張雅婷'),
    ('與錢同行','張雅婷'),
    ('論文門開箱','何橞瑢'),
    ('論文門開箱','何橞瑢'),
    ('鄭妹看世界','鄭亦真'),
    ('螃蟹秀開鍘','劉盈秀'),
    ('獸身男女','賴正鎧'),
    ('靈異錯別字_鬼錯字','賴正鎧'),
    ('琴謙天下事','周玉琴'),
    ('愛吃星球','張若妤'),
    ('誰謀殺了言論自由','中天電視'),
    ('誰謀殺了言論自由_俠姊','俠姊'),
    ('誰謀殺了言論自由_金汝鑫','金汝鑫'),
    ('誰謀殺了言論自由_賴麗櫻','賴麗櫻'),
    ('大新聞大爆卦','馬千惠'),
    ('週末大爆卦','李珮瑄'),
    ('正常發揮','丹申廣'),
    ('國際直球對決','林嘉源'),
    ('前進戰略高地','鄭亦真'),
    ('線上面對面','張介凡'),
    ('我是二寶爸','賴正鎧'),
    ('高級酸新聞台','李珮瑄')
    ]


order_list = [
    "(20)","(19)","(18)","(17)","(16)","(15)","(14)","(13)","(12)","(11)","(10)","(9)",
    "(8)","(7)","(6)","(5)","(4)","(3)","(2)","(1)",""
    ]

# 如果資料夾路徑存在就執行公式
for ch_yt in channel_youtuber_list:
    for order in order_list:
        file_extract(ch_yt[0]+order)
        channelpath = "頻道資訊/{}".format(ch_yt[0]+order)
        if os.path.isdir(channelpath):
            print('解壓縮處理<{}>的檔案'.format(str(ch_yt[0]+order)))
            table_gen(ch_yt[0]+order,ch_yt[1])
   


########################  bottom  ############################


print('===============  combine successful  ===============')
