import tkinter
import traceback
import sys,os
import time
import logging
import tkinter.messagebox as msg
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,TimeoutException

################## COMMON_FUNCTION ##################

def createDir():
    '''Check if directory exists, if not, create it'''
    dir1=os.path.isdir('resource')
    dir2=os.path.isdir('templates')
    dir3=os.path.isdir('templates')
    
    if not dir1 or not dir2 or not dir3:
        os.makedirs('resource')
        os.makedirs('templates')
        os.makedirs('log')

def makeComponent():
    '''tkinter INIT'''
    root = tkinter.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    tkinter.Tk.report_callback_exception = callback_error#ERROR HANDLER FOR TKINTER
    
    root.title("ダウンロードするファイルを選んでください。")
    root.geometry("500x78+"+str(int(screen_width/2)-200)+"+"+str(int(screen_height/2)-200))
    tkinter.Button(root,text="y_ALL*.csv",width=190,command=getCsv_y_ALL).pack()
    tkinter.Button(root,text="tp_*.xlsx",width=190,command=getExcel_tp).pack() 
    tkinter.Button(root,text="ippanmeishohoumaster.xlsx",width=190,command=getExcel_ippan).pack()
    root.mainloop()

def callback_error(self, *args):
    '''tkinter Exception Handle '''
    err = traceback.format_exception(*args)
    set_logging(err)
    releaseDriver()

def set_logging(ex):
    '''LOG'''
    # logger instance
    logger = logging.getLogger(__name__)
    # formatter create
    formatter = logging.Formatter('[%(asctime)s][%(levelname)s|%(filename)s:%(lineno)s] >> %(message)s')
    # handler
    streamHandler = logging.StreamHandler()
    fileHandler = logging.FileHandler(os.getcwd()+"/log.txt")
    # fomatter init
    streamHandler.setFormatter(formatter)
    fileHandler.setFormatter(formatter)
    # logger add handler
    logger.addHandler(streamHandler)
    logger.addHandler(fileHandler)
    #出力
    logger.critical(ex)
    msg.showerror(title='エラー', message='エラー履歴を確認してください。')

def initWebDriver():
    '''check Chrome driver version '''
    chromedriver_autoinstaller.install(cwd=True)
    #set downLoad path
    path = os.path.dirname(os.path.abspath(__file__))+'\\resource'
    options = webdriver.ChromeOptions() 
    options.add_experimental_option("prefs", {"download.default_directory":path})
    options.add_argument('--no-sandbox')
    
    # driver Global
    global driver
    driver=webdriver.Chrome(options = options) #return driver

def releaseDriver():
    '''release memory of Driver'''
    driver.close()
    driver.quit()

def setUrl_wait(url):
    '''set url for click and get Wait '''
    driver.get(url) #change URL
    return WebDriverWait(driver , 10)#wait 10sec


################## USER_METHOD ##################

def getCsv_y_ALL():
    initWebDriver()
    url='https://www.ssk.or.jp/seikyushiharai/tensuhyo/kihonmasta/kihonmasta_04.html'
    wait=setUrl_wait(url)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '全件分ファイル(CSV:')]"))).click()
    time.sleep(4)
    releaseDriver()

def getExcel_tp():
    initWebDriver()
    url='https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000078916.html'
    setUrl_wait(url)
    liTags=driver.find_elements_by_css_selector('ul.m-listLink>li')
    # aTag=driver.find_elements_by_css_selector('ul.m-listLink>li>a')

    for liTag in liTags:
        aTag=liTag.find_element_by_tag_name("a")
        child=liTag.find_elements_by_xpath(".//*")
        if '薬価基準収載品目リストについて' in aTag.text and len(child)>1:
            aTag.click()
            break

    aTags_nextPage=driver.find_elements_by_css_selector('div.section>ul>li.ico-excel>a')
    for i in range(5):
        aTags_nextPage[i].click()
        time.sleep(1)
    time.sleep(1)
    releaseDriver()

def getExcel_ippan():
    initWebDriver()
    url='https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000078916.html'
    setUrl_wait(url)
    liTags=driver.find_elements_by_css_selector('ul.m-listLink>li')
    # aTag=driver.find_elements_by_css_selector('ul.m-listLink>li>a')
    
    for liTag in liTags:
        aTag=liTag.find_element_by_tag_name("a")
        child=liTag.find_elements_by_xpath(".//*")
        if '処方箋に記載する一般名処方の標準的な記載（一般名処方マスタ）について' in aTag.text and len(child)>1:
            aTag.click()
            break

    aTags_nextPage=driver.find_elements_by_css_selector('div.section>ul>li.ico-excel>a')
    for i in range(1):
        aTags_nextPage[i].click()
        time.sleep(1)
    time.sleep(1)
    releaseDriver()

if __name__ == "__main__":
    try:
        createDir()
        makeComponent()
        # driver.find_elements_by_xpath("//*[contains(text(), '全件分ファイル(CSV:')]")[0].click()
    except NoSuchElementException as ex:
        releaseDriver()
        msg.showerror(title='エラー', message='該当URLが変更されました。チェックしてください。')
    except Exception as ex:
        set_logging(ex)
