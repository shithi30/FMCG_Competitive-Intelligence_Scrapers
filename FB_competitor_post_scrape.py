## import
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import pandas as pd
import duckdb
from googleapiclient.discovery import build
from google.oauth2 import service_account
import win32com.client
from datetime import datetime
import time

## scrape

# accumulators
start_time = time.time()
pages = [
    'https://www.facebook.com/godrejno.1bd',
    'https://www.facebook.com/parachuteskinpure',
    'https://www.facebook.com/himalayabd',
    'https://www.facebook.com/dettolbd',
    'https://www.facebook.com/savlonbangladesh'
]

# preference
options = webdriver.ChromeOptions()
options.add_argument("ignore-certificate-errors")
options.add_argument("disable-notifications")

# open window
driver = webdriver.Chrome(options = options)
driver.maximize_window()

# url
for j in range(0, len(pages)):
    print("Scraping: " + pages[j])
    driver.get(pages[j])

    # cross
    time.sleep(3)
    ActionChains(driver).send_keys("\n").perform()

    # scroll smooth
    y = 500
    for timer in range(0, 20):
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, " + str(y) + ")")
        y = y + 500

    # soup
    soup_init = BeautifulSoup(driver.page_source, "html.parser")
    soup = soup_init.find_all("div", attrs={"role": "article"})

    # data
    page = []
    post_text = []
    date = []
    date_refined = []
    url = []
    logos = ['''<img src="data:image/png;base64,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" style="width:12px;height:12px;">''']
    posts = len(soup)
    for i in range(0, posts):

        # page
        try: val = soup[i].find("strong").get_text()
        except: val = ""
        page.append(val)

        # date
        try: val1 = soup[i].find("a", attrs={"class": "x1i10hfl xjbqb8w x6umtig x1b1mbwd xaqea5y xav7gou x9f619 x1ypdohk xt0psk2 xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x16tdsg8 x1hl2dhg xggy1nq x1a2a7pz x1heor9g xt0b8zv xo1l8bm"})["aria-label"]
        except: val1 = ""
        try: val2 = soup[i].find("a", attrs={"class": "x1i10hfl xjbqb8w x6umtig x1b1mbwd xaqea5y xav7gou x9f619 x1ypdohk xt0psk2 xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x16tdsg8 x1hl2dhg xggy1nq x1a2a7pz x1heor9g xt0b8zv xo1l8bm"}).get_text()
        except: val2 = ""
        date.append(val1 if val1!="" else val2)

        # url
        try: val = soup[i].find("a", attrs={"class": "x1i10hfl xjbqb8w x6umtig x1b1mbwd xaqea5y xav7gou x9f619 x1ypdohk xt0psk2 xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x16tdsg8 x1hl2dhg xggy1nq x1a2a7pz x1heor9g xt0b8zv xo1l8bm"})["href"].split("?")[0]
        except: val = ""
        url.append(val)

        # caption + hashtag = text
        try: cap = soup[i].find("div", attrs={"class": "xdj266r x11i5rnm xat24cr x1mh8g0r x1vvkbs x126k92a"}).get_text()
        except: cap = ""
        try: tag = soup[i].find("div", attrs={"class": "x11i5rnm xat24cr x1mh8g0r x1vvkbs xtlvy1s x126k92a"}).get_text()
        except: tag = ""
        txt = (cap + "\n" + tag).replace("'", "")
        post_text.append(txt)
        
#         # media
#         try: val = soup[i].find("img", attrs={"class": "x1ey2m1c xds687c x5yr21d x10l6tqk x17qophe x13vifvy xh8yej3 xl1xv1r"})["src"]
#         except: val = ""
#         print(val)
        
    # cover
    val = soup_init.find("img", attrs={"data-imgperflogname": "profileCoverPhoto"})["src"]
    val = '''<img src="''' + val + '''" style="width:270px;height:100px"><br>'''
    logos.append(val)

    # accumulate
    df = pd.DataFrame()
    df["post_text"] = post_text
    df["post_date"] = date
    df["post_url"] = url
    df["page"] = page
    df["page_url"] = pages[j]
    df["report_time"] = str(time.strftime("%Y-%m-%d %H:%M"))
    df = duckdb.query('''select * from df where post_url!='' ''').df()
    
    # refine date
    date = df["post_date"].tolist()
    for d in date:
        date_modified = d.split(" at")[0]
        last_of_date = date_modified.split()[-1]
        
        # format
        if len(last_of_date) == 1: date_modified = duckdb.query("select left(now()::timestamp  + interval '6 h' - interval '" + date_modified + "', 10) date_modified").df()["date_modified"].tolist()[0]
        elif last_of_date.isnumeric() == True: date_modified = datetime.strptime(date_modified, "%d %B %Y").strftime("%Y-%m-%d")
        else: date_modified = datetime.strptime(date_modified + " " + duckdb.query("select left(current_date, 4) current_year").df()["current_year"].tolist()[0], "%d %B %Y").strftime("%Y-%m-%d")
        
        # record
        date_refined.append(date_modified)
    df["post_date_refined"] = date_refined
    
    ## GSheet

    # credentials
    SERVICE_ACCOUNT_FILE = "read-write-to-gsheet-apis-1-04f16c652b1e.json"
    SAMPLE_SPREADSHEET_ID = "1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0"
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    # APIs
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # extract
    values = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='FB Posts!A1:H').execute().get('values', [])
    fb_df_prev = pd.DataFrame(values[1:] , columns = values[0])

    # transform
    qry = '''
    -- others
    select * from fb_df_prev where page_url!=''' + "'" + pages[j] + "'" + '''
    union all
    -- old
    select post_text, post_date, post_date_refined, post_url, page, page_url, 0 if_new, report_time from fb_df_prev where page_url=''' + "'" + pages[j] + "'" + '''
    union all
    -- new
    select post_text, post_date, post_date_refined, post_url, page, page_url, 1 if_new, report_time from df where concat(post_text, post_date_refined, page_url) not in(select concat(post_text, post_date_refined, page_url) from fb_df_prev) -- assuming there won't be > 1 blank/same-text posts on a day on a page
    '''
    fb_df_pres = duckdb.query(qry).df()

    # load
    sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='FB Posts').execute()
    sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="'FB Posts'!A1", valueInputOption='USER_ENTERED', body={'values': [fb_df_pres.columns.values.tolist()] + fb_df_pres.fillna('').values.tolist()}).execute()

    ## novelty
    df_new = duckdb.query('''select * from fb_df_pres where if_new=1 and page_url=''' + "'" + pages[j] + "'").df()
    new_heads = df_new['post_text'].tolist()
    new_heads = [h if len(h)<65 else h[0:65] + " ..." for h in new_heads]
    new_links = df_new['post_url'].tolist()
    new_dates = df_new['post_date'].tolist()
    new_len = df_new.shape[0]
    
    ## MSTeams

    # email
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    
    # report
    new = logos[1] + logos[0] + " The following " + str(new_len) + " post(s) are newly found on <b>" + page[0] + "</b>"
    for i in range(0, new_len): new = new + '''<br>&nbsp;&nbsp;&nbsp;‣ ''' + new_heads[i] + ''' [<a href="''' + new_links[i] + '''">''' + new_dates[i] + '''</a>]''' 

    # Teams
    newmail.To = "Facebook Updates - Auto Monitoring <e9a950e7.Unilever.onmicrosoft.com@emea.teams.ms>"
    newmail.HTMLbody = new + "<br><br>"
    if new_len > 0: newmail.Send()
        
# close window
driver.close()

## stats
print("Total posts in result: " + str(fb_df_pres.shape[0]))
print("Elapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))