#!/usr/bin/env python
# coding: utf-8

# In[1]:


## import
import pandas as pd
import duckdb
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import re
import win32com.client
from pretty_html_table import build_table
import time        


# In[2]:


## scrape

# accumulators
start_time = time.time()
df_acc = pd.DataFrame()

# particulars
keywords = ['conditioner', 'handwash', 'bodywash', 'facewash', 'lotion', 'face cream', 'toothpaste', 'dishwash', 'toilet clean', 'soup', 'shampoo', 'health drink', 'detergent', 'moisturizer', 'soap', 'petroleum jelly', 'hair oil', 'germ kill']
brands = ['Boost Health', 'Boost Drink', 'Boost Jar', 'Clear Shampoo', 'Simple Fac', 'Simple Mask', 'Pepsodent', 'Brylcreem', 'Bru Coffee', 'St. Ives', 'St.Ives', 'Horlicks', 'Sunsilk', 'Sun Silk', 'Lux', 'Ponds', "Pond's", 'Closeup', 'Close Up', 'Cif', 'Dove', 'Maltova', 'Domex', 'Clinic Plus', 'Tresemme', 'Tresemmé', 'GlucoMax', 'Knorr', 'Glow Lovely', 'Fair Lovely', 'Glow Handsome', 'Wheel Wash', 'Axe Body', 'Pureit', 'Lifebuoy', 'Surf Excel', 'Vaseline', 'Vim', 'Rin']

# preference
options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')

# open window
driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.get("https://www.foodpanda.com.bd/darkstore/w2lx/pandamart-gulshan-w2lx")

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    elem = driver.find_element(By.XPATH, '//*[@id="groceries-menu-react-root"]/div/div/div[2]/div/section/div[3]/div/div/div/div/div[1]/input')
    elem.send_keys(k + "\n")

    # scroll
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        time.sleep(5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height: break
        last_height = new_height
    
    # soup
    soup_init = BeautifulSoup(driver.page_source, 'html.parser')
    soup = soup_init.find_all('div', attrs={'class', 'box-flex product-card-attributes'})

    # scrape
    sku = []
    current_price = []
    original_price = []
    offer = []
    if_ubl = [] 
    for s in soup:
        # sku
        try: val = s.find('p', attrs={'class', 'product-card-name'}).get_text()
        except: val = None
        sku.append(val)
        # current price
        try: val = s.find("span", attrs={"data-testid", "product-card-price"}).get_text().split()[1]
        except: val = None
        current_price.append(val)
        # original price
        try: val = s.find("span", attrs={"data-testid", "product-card-price-before-discount"}).get_text().split()[1]
        except: val = None
        original_price.append(val)
        # offer
        try: val = s.find("span", attrs={"class", "bds-c-tag__label"}).get_text()
        except: val = None
        offer.append(val)

    # accumulate
    df = pd.DataFrame()
    df['sku'] = sku
    df['current_price'] = current_price
    df['original_price'] = original_price
    df['offer'] = offer
    df['pos_in_pg'] = list(range(1, df.shape[0]+1))
    df['keyword'] = k
  
    # Unilever
    sku_count = len(sku)
    for i in range(0, sku_count):
        if_ubl.append(None)
        for b in brands:
            bb = b.split()
            if len(bb) == 1: bb.append('')
            if bb[0].lower() + ' ' in sku[i].lower() and bb[1].lower() in sku[i].lower(): if_ubl[i] = b
    df['brand_unilever'] = if_ubl

    # record
    df['report_time'] = time.strftime('%Y-%m-%d %H:%M:%S')
    df_acc = df_acc.append(df).fillna('')
    
    # back
    driver.back()

# close window
driver.close()


# In[3]:


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


# In[4]:


## previous data
prev_df = pd.read_excel(open("C:/Users/Shithi.Maitra/Downloads/Eagle Eye.xlsx", "rb"), sheet_name="Pandamart SoS", header=0, index_col=None).astype(str).replace("nan", "")
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
where current_price!=''
''').df()
display(prev_df)


# In[5]:


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
where current_price!=''
''').df()
display(pres_df)


# In[6]:


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
with pd.ExcelWriter("C:/Users/Shithi.Maitra/Downloads/CI Data - Pandamart.xlsx") as writer:
    change_df.to_excel(writer, sheet_name="CI Data", index=False)
    summ_df_sheet.to_excel(writer, sheet_name="Summary", index=True)


# In[7]:


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


# In[10]:


## email

# object
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = "CI Pandamart: " + time.strftime("%d-%b-%y")
# newmail.To = "shithi.maitra@unilever.com"
newmail.CC = "avra.barua@unilever.com; safa-e.nafee@unilever.com; rafid-al.mahmood@unilever.com; zoya.rashid@unilever.com; samsuddoha.nayeem@unilever.com; sudipta.saha@unilever.com; mehedi.asif@unilever.com; asif.rezwan@unilever.com; shithi.maitra@unilever.com"

# body
newmail.HTMLbody = '''
Dear concern,<br><br>
Thanks for sharing the datapoints to monitor for <b>Competitive Intelligence (CI)</b>. As discussed, the data have been fetched and the changes have been reported, as summarized below:
''' + build_table(summ_df, 'red_dark', font_size='13px') + '''
Note that, the statistics presented above and in the attachment are reflections from <a href="https://www.foodpanda.com.bd/darkstore/">Pandamart</a>, within the timeframe of scraping. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, CSE<br>
Unilever BD Ltd.<br>
'''
# attachment
filename = "C:/Users/Shithi.Maitra/Downloads/CI Data - Pandamart.xlsx"
newmail.Attachments.Add(filename)

# send
newmail.Send()


# In[9]:


## stats
display(change_df.head())
print("Changes in result: " + str(change_df.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))


# In[ ]:





# In[ ]:




