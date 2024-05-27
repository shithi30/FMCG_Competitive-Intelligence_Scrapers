#!/usr/bin/env python
# coding: utf-8

# In[3]:


## import
import pandas as pd
import duckdb
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import win32com.client
from pretty_html_table import build_table


# In[2]:


## scrape

# accumulators
start_time = time.time()
df_acc = pd.DataFrame()

# particulars
keywords = ['conditioner', 'handwash', 'bodywash', 'facewash', 'lotion', 'cream', 'toothpaste', 'dishwash', 'toilet clean', 'soup', 'shampoo', 'health drink', 'washing powder', 'wash liquid', 'detergent', 'moisturizer', 'soap', 'germ kill']
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

# url
for k in keywords:
    print("Scraping for keyword: " + k)
    url = "https://chaldal.com/search/" + k
    driver.get(url)

    # scroll
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height: break
        last_height = new_height

    # soup
    soup_init = BeautifulSoup(driver.page_source, 'html.parser')
    soup = soup_init.find_all("div", attrs={"class": "product"})

    # scrape
    skus = []
    quants = []
    prices = []
    prices_if_discounted = []
    options = []
    if_ubl = [] 
    for s in soup:
        # sku
        try: val = s.find("div", attrs={"class": "name"}).get_text()
        except: val = None
        skus.append(val)
        # quantity
        try: val = s.find("div", attrs={"class": "subText"}).get_text()
        except: val = None
        quants.append(val)
        # price
        try: val = float(s.find("div", attrs={"class": "price"}).get_text().split()[1].replace(',', ''))
        except: val = None
        prices.append(val)
        # discount
        try: val = float(s.find("div", attrs={"class": "discountedPrice"}).get_text().split()[1].replace(',', ''))
        except: val = None
        prices_if_discounted.append(val)
        # option
        try: val = s.find("p", attrs={"class": "buyText"}).get_text() 
        except: val = None
        options.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['quantity'] = quants
    df['price'] = prices
    df['price_if_discounted'] = prices_if_discounted
    df['option'] = options
    df['pos_in_pg'] = list(range(1, df.shape[0]+1))
    df['keyword'] = k
    df['relevance'] = ['relevant' if is_subseq(k.replace(' ', ''), s.lower()) else 'irrelevant' for s in skus]
    
    # Unilever
    sku_count = len(skus)
    for i in range(0, sku_count):
        if_ubl.append(None)
        for b in brands:
            bb = b.split()
            if len(bb) == 1: bb.append('')
            if bb[0].lower() + ' ' in skus[i].lower() and bb[1].lower() in skus[i].lower(): if_ubl[i] = b
    df['brand_unilever'] = if_ubl

    # record
    df['report_time_to'] = time.strftime('%Y-%m-%d %H:%M:%S')
    df_acc = df_acc.append(df.fillna('').astype(str))

# close window
driver.close()


# In[3]:


## previous data

# read
prev_df = pd.read_excel(open("C:/Users/Shithi.Maitra/Downloads/Eagle Eye.xlsx", "rb"), sheet_name="Chaldal SoS", header=0, index_col=None).astype(str).replace("nan", "")
# separate
prev_basepack = prev_df['sku'].tolist()
prev_qty = prev_df['quantity'].tolist()  
for i in range(0, len(prev_basepack)): prev_basepack[i] = prev_basepack[i].replace(prev_qty[i], "").strip()
prev_df['basepack'] = prev_basepack
# relevant
prev_df = duckdb.query('''select basepack, quantity grammage, price, price_if_discounted, keyword, brand_unilever, report_time from prev_df where relevance='relevant' and price!='' ''').df()
display(prev_df)


# In[4]:


## present data
pres_df = duckdb.query('''select basepack, quantity grammage, price, price_if_discounted, keyword, brand_unilever, report_time_to from df_acc where relevance='relevant' and price!='' ''').df()
display(pres_df)


# In[5]:


## compare

# changes 
qry = '''
-- price
select basepack, grammage, 'price' attr_changed, attr_prev, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, grammage attr_unchanged, price attr_prev, keyword from prev_df) tbl1 
    inner join 
    (select basepack, grammage attr_unchanged, price attr_now, keyword, grammage, brand_unilever, report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_prev!=attr_now

-- offer
union all
select basepack, grammage, 'offer price' attr_changed, attr_prev, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, grammage attr_unchanged, price_if_discounted attr_prev, keyword from prev_df) tbl1 
    inner join 
    (select basepack, grammage attr_unchanged, price_if_discounted attr_now, keyword, grammage, brand_unilever, report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_prev!=attr_now

-- grammage
union all
select basepack, grammage, 'grammage' attr_changed, attr_prev, attr_now, keyword, brand_unilever, report_time_to
from 
    (select basepack, price attr_unchanged, grammage attr_prev, keyword from prev_df) tbl1 
    inner join 
    (select basepack, price attr_unchanged, grammage attr_now, keyword, grammage, brand_unilever, report_time_to from pres_df) tbl2 using(basepack, attr_unchanged, keyword)
where attr_prev!=attr_now

-- new
union all
select basepack, grammage, 'new in results' attr_changed, '-' attr_prev, '-' attr_now, keyword, brand_unilever, report_time_to
from pres_df
where (basepack, grammage) not in(select (basepack, grammage) from prev_df)
    
-- dropped
union all
select basepack, grammage, 'dropped from results' attr_changed, '-' attr_prev, '-' attr_now, keyword, brand_unilever, (select max(report_time_to) from pres_df) report_time_to
from prev_df
where (basepack, grammage) not in(select (basepack, grammage) from pres_df)
'''
change_df = duckdb.query(qry).df()
change_df = duckdb.query('''select keyword, basepack, grammage, attr_changed, attr_prev, attr_now, brand_unilever, (select min(report_time) from prev_df) report_time_from, report_time_to from change_df order by keyword, attr_changed''').df()

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
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/CI Data - Chaldal.xlsx") as writer:
    change_df.to_excel(writer, sheet_name="CI Data", index=False)
    summ_df_sheet.to_excel(writer, sheet_name="Summary", index=True)


# In[6]:


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


# In[9]:


## email

# object
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = "CI Chaldal: " + time.strftime("%d-%b-%y")
# newmail.To = "shithi.maitra@unilever.com"
newmail.CC = "avra.barua@unilever.com; safa-e.nafee@unilever.com; rafid-al.mahmood@unilever.com; zoya.rashid@unilever.com; samsuddoha.nayeem@unilever.com; sudipta.saha@unilever.com; mehedi.asif@unilever.com; asif.rezwan@unilever.com; shithi.maitra@unilever.com"

# body
newmail.HTMLbody = '''
Dear concern,<br><br>
Thanks for sharing the datapoints to monitor for <b>Competitive Intelligence (CI)</b>. As discussed, the data have been fetched and the changes have been reported, as summarized below:
''' + build_table(summ_df, 'yellow_dark', font_size='13px') + '''
Note that, the statistics presented above and in the attachment are reflections from <a href="https://chaldal.com/">Chaldal.com</a>, within the timeframe of scraping. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, CSE<br>
Unilever BD Ltd.<br>
'''
# attachment
filename = "C:/Users/Shithi.Maitra/Downloads/CI Data - Chaldal.xlsx"
newmail.Attachments.Add(filename)

# send
newmail.Send()


# In[8]:


## stats
display(change_df.head())
print("Changes in result: " + str(change_df.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))


# In[ ]:




