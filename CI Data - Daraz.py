#!/usr/bin/env python
# coding: utf-8

# In[1]:


## import
import pandas as pd
import duckdb
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from bs4 import BeautifulSoup
import re
import win32com.client
from pretty_html_table import build_table
import time        


# In[3]:


## scrape

# accumulators
start_time = time.time()
df_acc = pd.DataFrame()

# particulars
keywords = ['conditioner', 'handwash', 'bodywash', 'facewash', 'lotion', 'cream', 'toothpaste', 'dishwash', 'toilet clean', 'soup', 'shampoo', 'health drink', 'detergent', 'moisturizer', 'soap', 'petroleum jelly', 'hair oil', 'germ kill']
brands = ['Boost Health', 'Boost Drink', 'Boost Jar', 'Clear Shampoo', 'Simple Fac', 'Simple Mask', 'Pepsodent', 'Brylcreem', 'Bru Coffee', 'St. Ives', 'St.Ives', 'Horlicks', 'Sunsilk', 'Sun Silk', 'Lux', 'Ponds', "Pond's", 'Closeup', 'Close Up', 'Cif', 'Dove', 'Maltova', 'Domex', 'Clinic Plus', 'Tresemme', 'Tresemmé', 'GlucoMax', 'Knorr', 'Glow Lovely', 'Fair Lovely', 'Glow Handsome', 'Wheel Wash', 'Axe Body', 'Pureit', 'Lifebuoy', 'Surf Excel', 'Vaseline', 'Vim', 'Rin']

# subsequence
def is_subseq(x, y):
    it = iter(y)
    return all(any(c == ch for c in it) for ch in x)

# preference
options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')

# open window
driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get('https://www.daraz.com.bd/')

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    elem = driver.find_element(By.ID, "q")
    elem.send_keys(Keys.CONTROL + "a")
    elem.send_keys(Keys.DELETE)
    elem.send_keys(k + "\n")

    # initialize
    pg = 1
    pos = 0
    new_skus = set()
    while(1):

        # scroll
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            time.sleep(5)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height

        # soup
        soup_init = BeautifulSoup(driver.page_source, "html.parser")
        soup = soup_init.find_all("div", attrs={"class": "gridItem--Yd0sa"})

        # scrape
        sku = []
        basepack = []
        current_price = []
        original_price = []
        offer = []
        rating = []
        reviews = []
        in_mall = []
        in_mart = []
        position = []
        sku_count = len(soup)
        for i in range(0, sku_count):
            # SKU
            try: val = soup[i].find("div", attrs={"id": "id-title"}).get_text()
            except: val = None
            sku.append(val)
            # basepack
            try: val = sku[i].split(re.compile("\d").findall(sku[i])[0])[0]
            except: val = sku[i]
            basepack.append(val)
            # current price
            try: val = soup[i].find("span", attrs={"class": "currency--GVKjl"}).get_text()
            except: val = None
            current_price.append(val)
            # original price
            try: val = soup[i].find("del", attrs={"class": "currency--GVKjl"}).get_text()[2:]
            except: val = None
            original_price.append(val)
            # offer        
            try: val = soup[i].find("div", attrs={"class": "voucher-wrapper--vCNzH"}).get_text()
            except: val = None
            offer.append(val)
            # rating    
            try: val = soup[i].find("span", attrs={"class": "ratig-num--KNake rating--pwPrV"}).get_text()
            except: val = None
            rating.append(val)
            # reviews
            try: val = soup[i].find("span", attrs={"class": "rating__review--ygkUy"}).get_text()[1:-1]
            except: val = None
            reviews.append(val)
            # mall
            in_mall.append(1)
            try: soup[i].find("i", attrs={"style": "background-image: url(&quot;https://img.alicdn.com/imgextra/i2/O1CN01m9OC6a1UK86X51Dcq_!!6000000002498-2-tps-108-54.png&quot;); width: 32px; height: 16px; vertical-align: text-bottom;"})["class"]
            except: in_mall[i] = 0
            # mart
            in_mart.append(1)
            try: soup[i].find("i", attrs={"style": "background-image: url(&quot;https://img.alicdn.com/imgextra/i1/O1CN01gS7Ros1VI7zYtUDwQ_!!6000000002629-2-tps-64-32.png&quot;); width: 32px; height: 16px; vertical-align: text-bottom;"})["class"]
            except: in_mart[i] = 0
            # position
            pos = pos + 1
            position.append(pos)

        # novelty
        skus_before = len(new_skus)
        for s in sku: 
            if is_subseq(k.replace(' ', ''), s.lower()): 
                new_skus.add(s)
        if len(new_skus) == skus_before: break

        # accumulate 
        df = pd.DataFrame()
        df['sku'] = sku
        df['basepack'] = basepack
        df['grammage'] = [s.split(b)[1] if b!="" else "" for (b, s) in zip(basepack, sku)]
        df['current_price'] = current_price
        df['original_price'] = original_price
        df['offer'] = offer
        df['rating'] = rating
        df['reviews'] = reviews
        df['in_mall'] = in_mall
        df['in_mart'] = in_mart
        df['keyword'] = k
        df['relevance'] = ['relevant' if is_subseq(k.replace(' ', ''), s.lower()) else 'irrelevant' for s in sku]
        df['pg_no'] = pg
        df['position'] = position

        # Unilever
        if_ubl = []
        skus = len(sku)
        for i in range(0, skus):
            if_ubl.append(None)
            for b in brands:
                bb = b.split()
                if len(bb) == 1: bb.append('')
                if bb[0].lower() + ' ' in sku[i].lower() and bb[1].lower() in sku[i].lower(): if_ubl[i] = b
        df['brand_unilever'] = if_ubl

        # record
        df['report_time'] = time.strftime('%Y-%m-%d %H:%M:%S')
        df_acc = df_acc.append(df).fillna('')

        # next page
        elem = driver.find_element(By.CLASS_NAME, "ant-pagination-next")
        ActionChains(driver).move_to_element(elem).click().perform()
        pg = pg + 1

# close window
driver.close()
        


# In[4]:


## separation
def get_gm_bp(skus):

    # accumulators
    grammage = []
    basepack = []

    # patterns
    pattern_gm = re.compile("[\d\.\+\±]+\s*(?:grams|gram|gm|kg|k.g|g|oz)", re.IGNORECASE)
    pattern_ml = re.compile("[\d\.\+\±]+\s*(?:liters|litres|litre|liter|ltr.|ltr|L|ml)", re.IGNORECASE)
    pattern_pc = re.compile("[\d\.\+\±]+\s*(?:pieces|piece|pcs|pc|ps|pics|pic|pes)", re.IGNORECASE)
    pattern_pk = re.compile("[\d\.\+\±]+\s*(?:packs|pack|pair|ply|boxes|box|sachets|sachet|ton|inches|inch|sets|set|sheets|sheet|rolls|roll)", re.IGNORECASE)

    # grammage
    for s in skus:
        sku = re.sub("Get", "", s, flags = re.IGNORECASE)
        vals = pattern_gm.findall(sku)
        if len(vals) == 0: vals = pattern_ml.findall(sku)
        if len(vals) == 0: vals = pattern_pc.findall(sku)
        if len(vals) == 0: vals = pattern_pk.findall(sku)

        # basepack
        try: val = vals[0]
        except: val = "not found"
        grammage.append(val)
        basepack.append(re.sub(" +", " ", s.replace(val, "")).strip())

    # record
    ret_df = pd.DataFrame()
    ret_df['grammage'] = grammage
    ret_df['basepack'] = basepack
    return ret_df


# In[5]:


## previous data
prev_df = pd.read_excel(open("C:/Users/Shithi.Maitra/Downloads/Eagle Eye.xlsx", "rb"), sheet_name="Daraz SoS", header=0, index_col=None).astype(str).replace("nan", "")
df_sep = get_gm_bp(prev_df['sku'].tolist())
prev_df['grammage'] = df_sep['grammage'].tolist()
prev_df['basepack'] = df_sep['basepack'].tolist()
prev_df = duckdb.query('''
select distinct
    sku, basepack, grammage, 
    replace(current_price, ',', '')::float current_price, 
    replace(case when original_price='' then current_price else original_price end, ',', '')::float original_price, 
    keyword, brand_unilever, report_time
from prev_df
where 
    relevance='relevant'
    and current_price!=''
''').df()
display(prev_df)


# In[6]:


## present data
df_sep = get_gm_bp(df_acc['sku'].tolist())
df_acc['grammage'] = df_sep['grammage'].tolist()
df_acc['basepack'] = df_sep['basepack'].tolist()
pres_df = duckdb.query('''
select distinct
    sku, basepack, grammage, 
    replace(current_price, ',', '')::float current_price, 
    replace(case when original_price='' then current_price else original_price end, ',', '')::float original_price, 
    keyword, brand_unilever, report_time 
from df_acc
where 
    relevance='relevant'
    and current_price!=''
''').df()
display(pres_df)


# In[7]:


## changes

# compare
qry = '''
-- price
select distinct basepack, grammage, 'price' attr_changed, attr_previous, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, grammage attr_unchanged, original_price attr_previous, keyword from prev_df) tbl1 
    inner join 
    (select basepack, grammage attr_unchanged, original_price  attr_now, keyword, grammage, brand_unilever, report_time report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_previous!=attr_now

-- offer
union all
select distinct basepack, grammage, 'offer price' attr_changed, attr_previous, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, grammage attr_unchanged, current_price attr_previous, keyword from prev_df) tbl1 
    inner join 
    (select basepack, grammage attr_unchanged, current_price  attr_now, keyword, grammage, brand_unilever, report_time report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_previous!=attr_now

-- grammage
union all
select distinct basepack, grammage, 'grammage' attr_changed, attr_previous, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, original_price attr_unchanged, grammage attr_previous, keyword from prev_df) tbl1 
    inner join 
    (select basepack, original_price attr_unchanged, grammage  attr_now, keyword, grammage, brand_unilever, report_time report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_previous!=attr_now

-- dropped
union all
select distinct basepack, grammage, 'dropped from results' attr_changed, '-' attr_previous, '-' attr_now, keyword, brand_unilever, (select max(report_time) from pres_df) report_time_to
from prev_df
where (basepack, grammage) not in(select (basepack, grammage) from pres_df)

-- new
union all
select distinct basepack, grammage, 'new in results' attr_changed, '-' attr_previous, '-' attr_now, keyword, brand_unilever, report_time report_time_to
from pres_df
where (basepack, grammage) not in(select (basepack, grammage) from prev_df)
'''
change_df = duckdb.query(qry).df()
change_df = duckdb.query('''select keyword, basepack, grammage, attr_changed, attr_previous, attr_now, brand_unilever, (select min(report_time) from prev_df) report_time_from, report_time_to from change_df order by keyword, attr_changed''').df()

# summary - sheet
qry = '''
select 
    keyword,
    attr_changed, 
    sum(case when brand_unilever!='' then 1 else 0 end) changes_ubl,
    sum(case when brand_unilever='' then 1 else 0 end) changes_nonubl
from change_df
group by 1, 2
order by 1, 2
'''
piv = duckdb.query(qry).df()
summ_df_sheet = piv.pivot(index="keyword", columns="attr_changed")

# store
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/CI Data - Daraz.xlsx") as writer:
    change_df.to_excel(writer, sheet_name="CI Data", index=False)
    summ_df_sheet.to_excel(writer, sheet_name="Summary", index=True)


# In[8]:


## summary - email
qry = '''
select 
    attr_changed "Attr. Changed", 
    count(case when brand_unilever!='' then 1 else null end) "Changes - UBL",
    count(case when brand_unilever='' then 1 else null end) "Changes - nonUBL",
    max(report_time_from) "Reporting From",
    max(report_time_to) "Reporting Till"
from change_df
group by 1
'''
summ_df = duckdb.query(qry).df()
display(summ_df)


# In[11]:


## email

# object
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = "CI Daraz: " + time.strftime("%d-%b-%y")
# newmail.To = "shithi.maitra@unilever.com"
newmail.CC = "avra.barua@unilever.com; safa-e.nafee@unilever.com; rafid-al.mahmood@unilever.com; zoya.rashid@unilever.com; samsuddoha.nayeem@unilever.com; sudipta.saha@unilever.com; mehedi.asif@unilever.com; asif.rezwan@unilever.com; shithi.maitra@unilever.com"

# body
newmail.HTMLbody = '''
Dear concern,<br><br>
Thanks for sharing the datapoints to monitor for <b>Competitive Intelligence (CI)</b>. As discussed, the data have been fetched and the changes have been reported, as summarized below:
''' + build_table(summ_df, 'orange_dark', font_size='13px') + '''
Note that, the statistics presented above and in the attachment are reflections from <a href="daraz.com.bd">daraz.com.bd</a>, within the timeframe of scraping. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, CSE<br>
Unilever BD Ltd.<br>
'''
# attachment
filename = "C:/Users/Shithi.Maitra/Downloads/CI Data - Daraz.xlsx"
newmail.Attachments.Add(filename)

# send
newmail.Send()


# In[10]:


## stats
display(change_df.head())
print("Changes in result: " + str(change_df.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))


# In[ ]:




