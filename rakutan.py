import csv
import os
import sys
import time
import warnings
from datetime import datetime

import schedule
from bs4 import BeautifulSoup
from linebot import LineBotApi
from linebot.models import TextSendMessage
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait


class RakuTan:
    def __init__(self):
        warnings.filterwarnings("ignore", category=DeprecationWarning)
        self.op = Options()
        self.op.add_argument("--headless")
        self.op.add_argument("--no-sandbox")
        self.op.add_argument("--disable-extensions")
        self.op.add_argument('--disable-gpu')  
        self.op.add_argument('--ignore-certificate-errors')
        self.op.add_argument('--allow-running-insecure-content')
        self.op.add_argument('--disable-web-security')
        self.op.add_argument('--disable-desktop-notifications')
        self.op.add_argument("--disable-extensions")
        self.op.add_argument('--lang=ja')
        self.op.add_argument('--blink-settings=imagesEnabled=false')
        self.op.add_argument('--disable-dev-shm-usage')
        self.op.add_argument('--proxy-server="direct://"')
        self.op.add_argument('--proxy-bypass-list=*')
        self.op.add_argument('--disable-logging')
        self.op.add_argument('--log-level=3')
        self.op.add_argument("--disable-dev-shm-usage")
        self.op.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.op.add_experimental_option("excludeSwitches", ['enable-automation'])
        
        self.eraa = None
        self.freq = 1
        self.url = "https://eduweb.sta.kanazawa-u.ac.jp/Portal/Public/Regist/RegistrationStatus.aspx?year=2023&lct_term_cd=22"
        self.lba = LineBotApi("YOUR_ACCESS_TOKEN_HERE")

    def syokika(self):
        sabisu = Service(executable_path='path/to/chromedriver')
        doraiba = webdriver.Chrome(service=sabisu, options=self.op)
        doraiba.implicitly_wait(30)
        return doraiba
    
    def strato(self):
        try:
            self.driver.get(self.url)
            dropdown = self.driver.find_element(By.ID, "ctl00_phContents_ucRegistrationStatus_ddlLns_ddl")
            selected = Select(dropdown).first_selected_option.get_attribute("value")
            if selected != "0":
                Select(dropdown).select_by_value("0")
            WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_all_elements_located)
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            return soup
        except Exception as e:
            print(f"[strato] An error occurred during startup.: {e}")

    def chenji(self, soup):
        if soup:
            cd = os.getcwd()
            koshin = soup.find("span", id="ctl00_phContents_ucRegistrationStatus_lblDate").text.strip()
            nau = datetime.strptime(koshin, "%Y/%m/%d %H:%M:%S")
            naus = nau.strftime("%Y年%m月%d日%H:%M:%S現在")
            naui = nau.strftime("%Y%m%d%H%M%S")
            sa = datetime.now() - nau
            if sa.total_seconds() > 60:
                print("[chenji] Exceeds 60 seconds.")
                raise Exception
            else:
                csvd = os.path.join(cd, 'hosei')
                name = os.path.join(csvd, f"hosei_{naui}.csv")
                os.makedirs(csvd, exist_ok=True)
                with open(name, 'a', newline='', encoding='utf-8') as csv_file:
                    lighter = csv.writer(csv_file)
                    tr = soup.select('#ctl00_phContents_ucRegistrationStatus_gv tr')
                    for r in tr:
                        columns = r.find_all(['td', 'th'])
                        data = [col.text.strip() for col in columns]
                        lighter.writerow(data)
                print(f"hosei_{naui}.csv({datetime.now()})")
                return naui, naus
        else:
            print(f"[chenji] An error in chenji.")
            raise

    def fuga(self, naui, naus):
        ls = os.listdir('hosei')
        if len(ls) < 2:
            print("[fuga] There is only one file.")
            return
        ls.sort(key=lambda x: os.path.getmtime(os.path.join('hosei', x)), reverse=True)
        prime = os.path.join('hosei', ls[0])
        sub = os.path.join('hosei', ls[1])
        with open(prime, 'r') as pf:
            pls = pf.read().strip().split('\n')
        with open(sub, 'r') as sf:
            sls = sf.read().strip().split('\n')
        for pl, sl in zip(pls, sls):
            pv = pl.split(',')
            sv = sl.split(',')
            if sv[1] == "ＧＳ科目":
                if int(sv[8]) == 0:
                    if int(pv[8]) != 0:
                        print(f"{pv}({datetime.now()})")
                        messe = f"{pv[2]}\n{pv[3]} / {pv[4]} / {pv[0]}\n適正人数:{pv[6]} / 登録人数:{pv[7]} / 残数:{pv[8]}\n{naus}\nｉｉｉｉ今がチャソス！！！！"
                        print(naus)
                        self.okuru(messe)
                        self.rogu(naui, pl)

    def okuru(self, messe):
         self.lba.broadcast(TextSendMessage(text=messe))

    def rogu(self, naui, pl):
        cd = os.getcwd()
        diffd = os.path.join(cd, 'diff')
        path = os.path.join(diffd, "diff.csv")
        os.makedirs(diffd, exist_ok=True)
        values = pl.strip().split(',')
        tsuiki = [naui] + values
        with open(path, 'a', newline='', encoding='utf-8') as csvfile:
            csv.writer(csvfile).writerow(tsuiki)

    def liset(self):
        print("[liset] Restart your browser.")
        self.driver.quit()
        self.driver = self.syokika()
        time.sleep(3)
        self.strato()

    def rupu(self):
        limit = 5
        for i in range(1, limit + 1):
            try:
                soup = self.strato()
                naui, naus = self.chenji(soup)
                self.fuga(naui, naus)
                if self.eraa is not None:
                    self.eraa = None
                break
            except KeyboardInterrupt:
                print("[rupu] Interrupted by Ctrl+C.")
                self.driver.quit()
                sys.exit()
            except Exception as e:
                print(f"[rupu] {i}times error occurred: {e}")
                if i >= limit:
                    if e == self.eraa:
                        print(f"[rupu] Forced to terminate due to the same error in 2 consecutive loops.")
                        self.driver.quit()
                        sys.exit()
                    self.eraa = e
                    self.liset()
                    break

    def lan(self):
        self.driver = self.syokika()
        self.rupu()
        schedule.every().hour.at("00:28").do(self.liset)
        for i in range(60 // self.freq):
            schedule.every().hour.at(f"{str(i * self.freq).zfill(2)}:03").do(self.rupu)
        while True:
            schedule.run_pending()
            time.sleep(1)

if __name__ == "__main__":
    RakuTan().lan()
