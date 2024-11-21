from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime, time, os, re, win32com.client, shutil, telepot
from selenium.webdriver.chrome.service import Service as ChromeService #//240221_1 : To Fix options Error
from selenium.webdriver.chrome.options import Options as ChromeOptions #//240221_1
from webdriver_manager.chrome import ChromeDriverManager #//240221_1
import time

###################################################################
#   Define List, Variable
###################################################################
url_list = []
diff = []
Cur_price_column = 8
Old_price_column = 9
Cur_URL_column = 24 #//241016
Old_URL_column = 25 #//241016
Delivery_price_column = 10
row_number = 4
item_numbers = 0
sleep_time = 1 ###2
wait_time = 2 ###5
Cur_item_name_column = 3

PATH = 'D:/01_MS_Work/02_Biz/01_MS_Global/02_구매대행/'
#PATH = 'D:/02_MS/01_MS_Work/02_Office/01_MS_Global/02_구매대행/'

###################################################################
#   Working Directory
###################################################################
### Log File 함수 ###
def write_log(msg):
    #print(msg)
    f = open(PATH + 'auto_aboard.log', 'a', encoding='UTF-8')
    f.write('[%s] %s\n' % (str(datetime.datetime.now()), msg))

### amazon 로그인 ####
def login_amazon():
    AMAZON_ID='starrynig99@gmail.com'
    AMAZON_PW='Mins50577578!!'

    ### 1. 아마존을 연다. https://www.amazon.com/-/us/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F%3Fref_%3Dnav_ya_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=usflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&
    try:
        driver.get('https://www.amazon.com/-/ko/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F%3Fref_%3Dnav_custrec_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=usflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&')

    ### 2. ID 창을 찾아서 선택 후 ID 입력
        elem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//input[@type="email"]')))
        elem.click()
        
        elem.send_keys(AMAZON_ID)
        elem.send_keys(Keys.RETURN)
        time.sleep(5)
        
    ### 3. PW 창을 찾아서 선택 수 PW 입력
        elem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//input[@type="password"]')))
        elem.click()
        
        elem.send_keys(AMAZON_PW)
        elem.send_keys(Keys.RETURN)
                
    except Exception as e:
        write_log(e)
        write_log("login Amazon Error!!!")
    
    finally:
        time.sleep(sleep_time)

    try:
        Approval_page_on = True
        elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-size-medium transaction-approval-word-break a-text-bold"]')))
        write_log('아마존 로그인 정보 verify 단계')
    except Exception as e:
        write_log(e)
        Approval_page_on = False
        write_log('아마존 로그인 정보 verify 없음')
    finally:
        Approval_pass = True

    if (Approval_page_on):
        Approval_pass = False
        write_log('Approval page 진입 후 Approval 대기')
        try:
            elem = WebDriverWait(driver, 60*5).until(EC.presence_of_element_located((By.XPATH, '//div[@class="nav-prime-1 nav-progressive-attribute"]')))
            Approval_pass = True
        except Exception as e:
            write_log(e)
            Approval_pass = False
        finally:
            if(Approval_pass):
                write_log('아마존 verify 성공')
            else:
                write_log('아마존 verify 실패')

    if not Approval_pass:
        write_log('프로그램 종료 !!!!!')
        bot = telepot.Bot('1146194999:AAED43PhvHMme3ibW80Fnlgq9XiIXqvugHI')
    
        msg = '##구매대행 : 프로그램 강제 종료!!!\n'
        
        write_log(msg)
        bot.sendMessage('714653402', msg)

        sys.exit()

    write_log('아마존 log in 성공')

## Amazon에서 Used Item 가격 찾기
def find_used_item_price_from_amazon(driver):
    print("Check Used Item from Amazon - Start")
    try:
        ## 중고인지 확인
        elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id,'usedBuySection') or contains(@id,'usedAccordionCaption_feature_div')]"))) ## Used case
        #print("u1")
        ## 중고면 Find Option Window and Click
        elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, './/span[@data-action="show-all-offers-display"]')))
        elem.click()
        #print("u2")
        # New라는 text만 갖는 요소를 찾는다.
        try:
            # New이면서 Prime인 요소를 찾는다.
            elem = WebDriverWait(driver, wait_time+4).until(EC.presence_of_element_located((By.XPATH, "//*[@id='aod-offer']//*[normalize-space(text())='New']/ancestor::*[5]//i[contains(@class, 'a-icon-prime')]")))
            #print("u3")
        except Exception as e:
            write_log(e)
            print(e)
            #print("u4")
            elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, '//*[@id="aod-offer"]//*[normalize-space(text())="New"]')))
            #print("u5")
        finally:
            e=None
            del e
        #print("u6")
        # New라는 Text만 갖는 요소에서 id가 aod-offer라는 것을 갖는 가장 가까운 부모를 찾는다.
        elem = elem.find_element(By.XPATH, "./ancestor::*[@id='aod-offer']//span[contains(@class, 'a-price aok-align-center centralizedApexPricePriceToPayMargin') or contains(@class, 'a-price aok-align-center reinventPricePriceToPayMargin priceToPay' or contains(@class, 'a-price a-text-price a-size-medium apexPriceToPay'))]")
        
        # 찾은 요소 출력
        print(elem.text)
        
        item_price = elem.text.replace('\n', '.')
        item_price = item_price.replace('$', '')
        
        item_price = float(item_price) ## 달러는 float 타입
    except Exception as e:
        write_log(e)
        print(e)
        item_price = -1
    finally:
        e=None
        del e    
    
    print("Check Used Item from Amazon - End")

    return item_price

## Amazon에서 New Item 가격 찾기
def find_new_item_price_from_amazon(driver):
    print("Check New Item from Amazon - Start")
    try:
        ## New 제품인지 확인 ## New or New 이지만 별도의 표기가 없는 Item 찾기
        elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, 'newAccordionRow_0') or contains(@id, 'qualifiedBuybox')]")))
        #elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-price aok-align-center"]')))
        try: ### 가격이 priceblock_ourprice / priceblock_dealprice / priceblock_saleprice 3곳중 한 곳에 표기 되어서, 하기와 같이 최소 1회/최대 3회 체크
                print("Amazon - 1 Start")
                elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-price a-text-normal aok-align-center reinventPriceAccordionT2"]')))
                                
                item_price = elem.text.replace('\n', '.')
                item_price = item_price.replace('$', '')
                
                item_price = float(item_price) ## 달러는 float 타입
                
                ws.Cells(row_number, Cur_price_column).Value = item_price
                
                write_log("Amazon - 1 End")
        except Exception as e:
            try:
                print("Amazon - 2 Start")

                elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-price aok-align-center reinventPricePriceToPayMargin priceToPay"]')))
                #elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-price aok-align-center"]')))
                
                item_price = elem.text.replace('\n', '.')
                item_price = item_price.replace('$', '')
                
                item_price = float(item_price) ## 달러는 float 타입
                
                ws.Cells(row_number, Cur_price_column).Value = item_price

                write_log("Amazon - 2 End")
            except Exception as e:
                try:
                    print("Amazon - 3 Start")
                    
                    elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, '//span[@class="a-price a-text-price a-size-medium apexPriceToPay"]')))
                    #elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//span[@id="priceblock_saleprice"]')))
                    
                    item_price = elem.text.replace(',', '')
                    item_price = item_price.replace('$', '')
                    
                    item_price = float(item_price) ## 달러는 float 타입
                    
                    ws.Cells(row_number, Cur_price_column).Value = item_price
                    
                    write_log("Amazon - 3 End")
                except Exception as e:
                    item_price = -1 ### 가격 정보를 찾지 못했을 때,

                    write_log("Amazon Error!!!")
                finally:
                    e = None
                    del e
            
            finally:
                e = None
                del e

        finally:
            write_log(item_price)            
                        
    except Exception as e:
        write_log(e)
        item_price = -1
    finally:
        e=None
        del e

    print("Check New Item from Amazon - End")
    return item_price

## Amazon에서 unqualified buy item(Option을 눌러야 하는) 가격 찾기
def find_unqualified_price_from_amazon(driver):
    print("find_unqualified_price_from_amazon : Start")
    try:
        ## unqualified buy item 찾기
        elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, 'unqualifiedBuyBox')]")))
        #print("t1")
        # See All Buying Options
        elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, ".//a[contains(text(), 'See All Buying Options')]")))
        #print("t2")
        # Click the button        
        elem.click()
        
        # New라는 text만 갖는 요소를 찾는다.
        try:
            # New이면서 Prime인 요소를 찾는다.
            elem = WebDriverWait(driver, wait_time+4).until(EC.presence_of_element_located((By.XPATH, "//*[@id='aod-offer']//*[normalize-space(text())='New']/ancestor::*[5]//i[contains(@class, 'a-icon-prime')]")))#"//*[@id='aod-offer']//*[normalize-space(text())='New'].//i[contains(@class, 'a-icon a-icon-prime')]")
            #print("t3")
        except Exception as e:
            write_log(e)
            #print(e)
            #print("t4")
            #elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, '//*[@id="aod-offer"]//*[normalize-space(text())="New"]')))
            elem = WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, "//*[@id='aod-offer']//*[normalize-space(text())='New']/ancestor::*[5]//span[contains(text(), 'FREE delivery')]")))
            #print("t5")
        finally:
            e=None
            del e
        #print("t6")
        # New라는 Text만 갖는 요소에서 id가 aod-offer라는 것을 갖는 가장 가까운 부모를 찾는다.
        elem = elem.find_element(By.XPATH, "./ancestor::*[@id='aod-offer']//span[contains(@class, 'a-price aok-align-center centralizedApexPricePriceToPayMargin') or contains(@class, 'a-price aok-align-center reinventPricePriceToPayMargin priceToPay' or contains(@class, 'a-price a-text-price a-size-medium apexPriceToPay'))]")

        # 찾은 요소 출력
        print(elem.text)
        
        item_price = elem.text.replace('\n', '.')
        item_price = item_price.replace('$', '')
        
        item_price = float(item_price) ## 달러는 float 타입
    except Exception as e:
        write_log(e)
        item_price = -1
    finally:
        e=None
        del e
    
    print("find_unqualified_price_from_amazon : End")
    return item_price

## Amazon 가격 찾기
def find_item_price_from_amazon(driver, mode_num):
    try:
        if mode_num == 1: ## New 제품일때 가격처리
            item_price = find_new_item_price_from_amazon(driver)
        elif mode_num == 2: ## 가격정보가 아예 없는 경우
            item_price = find_unqualified_price_from_amazon(driver)
        elif mode_num == 3: ## 중고일때 가격 처리        
            item_price = find_used_item_price_from_amazon(driver)
    
    except Exception as e:
        write_log(e)
        item_price = -1
    finally:
        e=None
        del e

    return item_price

'''def login_mrrebates():
    mrrebates_ID='decemberdream08@gmail.com'
    mrrebates_PW='tt64097578'
    driver.get('https://www.mrrebates.com/merchant.asp?id=9678')
    
    try: 
        elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//a[@class="button join-button alert"][@href="/login.asp"]')))
        elem.click()
        #print("login_mrrebates - Click OK!!!")
        
        ### 2. ID 창을 찾아서 선택 후 ID 입력
        elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@type="text"][@name="t_email_address"]')))
        elem.click()
        
        elem.send_keys(mrrebates_ID)
        elem.send_keys(Keys.RETURN)

        ### 3. PW 창을 찾아서 선택 수 PW 입력
        elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@type="password"][@name="t_password"]')))
        elem.click()

        elem.send_keys(mrrebates_PW)
        elem.send_keys(Keys.RETURN)

        elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//a[@class="small button ShopNowBtn"][@href="/click/nw.asp?merchant_id=9678"]')))
        elem.click()

        #print("login_mrrebates - ID/PW OK!!!")
        
        driver.switch_to.window(driver.window_handles[1])
        #print(driver.title)
        
        elem = WebDriverWait(driver, 60*2).until(EC.presence_of_element_located((By.XPATH, '//a[@class="g_a s_a g_c l_a"][@title="Walmart Homepage"]')))
        #print("login_mrrebates - move Walmart!!!")

    except Exception as e:
        write_log(e)
        #print("login_mrrebates() : Error!!!") # walmart는 로봇인지를 가끔 체크하여 Error 발생 !!!

    finally:
        write_log("Exit login_mrrebates()") '''


### Main 함수 Start ####

try:
    ### Excel File 정보 ###
    File_Name = '구매대행_판매상품관리'
    File_extension = '.xlsx'
    Excel_PATH = PATH + File_Name + File_extension
    print(Excel_PATH)
    #print(datetime.date.today())
    date = str(datetime.date.today())
    Excel_PATH2 = PATH + File_Name + '_' + date + File_extension
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(Excel_PATH)
    ws = wb.Worksheets('판매상품목록')

    ### Looking for number of item from Excel file ###
    for i in range(row_number, 100000):
        cell_value = ws.Cells(i, Cur_URL_column).Value
        
        if cell_value == None:
            break
        else:
            url_list.append(cell_value)
            ws.Cells(i, Old_price_column).Value = ws.Cells(i, Cur_price_column).Value
            #ws.Cells(i, Old_URL_column).Value = ws.Cells(i, Cur_URL_column).Value
            ws.Cells(i, Cur_price_column).Value = ''
            #ws.Cells(i, Cur_URL_column).Value = ''
    
    ''' ### 크롬 드라이버 설정
    #options = webdriver.ChromeOptions() #//240221_1
    options = ChromeOptions() #//240221_1
    # options.add_argument("start-maximized")
    # options.add_argument('--disable-blink-features=AutomationControlled')
    
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--incognito")
    options.add_argument("--start-maximized")

    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu")

    ### 크롬 드라이버 설정 - walmart는 headless 모드 지원하지 않으므로 Option 없이 Chrome 실행)
    service = ChromeService(executable_path=ChromeDriverManager().install()) #//240221_1
    driver = webdriver.Chrome(service=service, options=options) #//240221_1
    #driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver', options=options)
    #driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver')
    #driver = webdriver.Chrome('D:/02_MS/02_Study/01_Python/01_Code/02_Auto/chromedriver')

    ### walmart인 경우 우회하기 위해 하기 사이트를 이용 ####
    #login_mrrebates()
    driver.get("https://www.walmart.com/")
    print("Open web site !!")
    ### url list의 크기 만큼 크롬에서 url을 검색 - Walmart 만 검색
    for url in url_list:
        if 'walmart' in url:
            #driver.switch_to.window(driver.window_handles[2])
            write_log("Walmart web site !!")
            print("Walmart web site !!")
            driver.get(url)
            write_log(url)
            print(url)

            try: ### 가격이 priceblock_ourprice / priceblock_dealprice / priceblock_saleprice 3곳중 한 곳에 표기 되어서, 하기와 같이 최소 1회/최대 3회 체크
                ##elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[@id="price"]/div/span[@class="hide-content display-inline-block-m"]/span/span[@class="visuallyhidden"]')))
                elem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//span[@class="inline-flex flex-column"]/span[@itemprop="price"]')))
                
                item_price = elem.text.replace(',', '')
                item_price = item_price.replace('Now','')
                item_price = item_price.replace('$', '')
                item_price = float(item_price) ## 달러는 float 타입
                                        
                ws.Cells(row_number, Cur_price_column).Value = item_price

            except Exception as e:
                write_log(e)
                write_log("Walmart web Error!!!") # walmart는 로봇인지를 가끔 체크하여 Error 발생 !!!

                ### 가격 정보를 찾지 못했을 때, 현재 가격은 기존 가격으로 설정
                item_price = 0 ### 가격 정보를 찾지 못했을 때, 가격을 0으로 초기화

            finally:
                write_log(item_price)
                print(item_price)

                
            if ws.Cells(row_number, Old_price_column).Value != item_price:
                old_price = ws.Cells(row_number, Old_price_column).Value
                cur_price = item_price
                print("Old : ", old_price, "New : ", item_price)
                diff.append((row_number-3, ws.Cells(row_number, Cur_item_name_column).Value, ws.Cells(row_number, Old_price_column).Value, item_price))

        row_number += 1
        #time.sleep(1)

    row_number = 4

    ### Walmart에서 사용하던 Driver를 닫고 새로운 Driver로 시작
    driver.close()
    #driver.switch_to.window(driver.window_handles[0])
    driver.quit() '''
    
    ### 크롬 드라이버 설정 - Amazon 등 기타 웹사이트는 headless 모드 지원하므로 Option 설정 후 Chrome 실행)
    ''' options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("disable-gpu") '''

    #driver = webdriver.Chrome('D:/02_MS/02_Study/01_Python/01_Code/02_Auto/chromedriver')
    #driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver', options=options)
    #driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver')
    options = webdriver.ChromeOptions() #//240221_1

    service = ChromeService(executable_path=ChromeDriverManager().install()) #//240221_1
    driver = webdriver.Chrome(service=service, options=options) #//240221_1

    ### Amazon prime 가격을 얻기 위해 Log in을 실시 ###
    login_amazon()
    
    ### 새로운 탭을 열어서 이 탭에서 저장된 url을 호출하여 가격을 비교할 수 있게 한다.
    driver.execute_script('window.open("https://www.google.com/","", "_blank")')
    driver.switch_to.window(driver.window_handles[1])

    ### url list의 크기 만큼 크롬에서 url을 검색
    for url in url_list:
        
        ### walmart인 경우는 위에서 이미 실행 했으므로 PASS ####
        if 'walmart' in url:
            row_number += 1
            
            continue

        ### Amazon인 경우 ####
        elif 'amazon' in url:
            write_log("Amazon Web site !!")

            ### 새탭을 열어 품목 url을 입력
            driver.get(url)
            write_log(url)
            print("No.", row_number)
            
            ### 가격정보 가져오기 첫번째 시도
            item_price = find_item_price_from_amazon(driver, 1)
            
            ### find_result가 -1이면 Retry 시도.
            if item_price != -1:
                ws.Cells(row_number, Cur_price_column).Value = item_price
            else: ### find_result가 -1이면 다른 xpath로 Retry 시도.
                item_price = find_item_price_from_amazon(driver, 2)
                if item_price != -1:
                    ws.Cells(row_number, Cur_price_column).Value = item_price
                else: ### find_result가 -1이면 다른 xpath로 Retry 시도.
                    item_price = find_item_price_from_amazon(driver, 3)
                    if item_price != -1:
                        ws.Cells(row_number, Cur_price_column).Value = item_price
                    else:
                        item_price = 0 ### 가격 정보를 찾지 못했을 때, 현재 가격은 기존 가격으로 설정
            
            write_log(item_price)
            ### Debug
            print(item_price)
            #input()
                
            ### 기타 사이트 경우 ####
        else:
            write_log("This web site is other site !!")

            driver.get(url)            
            write_log(url)
                        
            ### Amazon 이외의 사이트의 가격 정보는 아직 구현이 되지 않았으므로, 현재 가격은 기존 가격으로 설정
            item_price = ws.Cells(row_number, Old_price_column).Value
            ws.Cells(row_number, Cur_price_column).Value = item_price
                            
        if ws.Cells(row_number, Old_price_column).Value != item_price:
            old_price = ws.Cells(row_number, Old_price_column).Value
            cur_price = item_price
            #print("Old : ", old_price, "New : ", item_price)
            diff.append((row_number-3, ws.Cells(row_number, Cur_item_name_column).Value, ws.Cells(row_number, Old_price_column).Value, item_price))

        row_number += 1
        time.sleep(sleep_time)
        
except Exception as e:
    write_log(e) 
    write_log("Program Error!!!")

finally:
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.quit()

### 엑셀 파일을 저장 후 종료
wb.Save()
wb.Close()
#excel.Quit()

### 엑셀 파일을 오늘 날짜로 업데이트
date = str(datetime.date.today())
New_Excel_PATH = PATH + File_Name + '_' + date + File_extension
shutil.copy(Excel_PATH, New_Excel_PATH)

### 기존 가격과 다른 경우, 해당 정보를 취합하여 텔레그램으로 발송한다.
if diff:
    write_log('변동된 가격 정보를 텔레그램으로 전송 합니다.')
    bot = telepot.Bot('1146194999:AAED43PhvHMme3ibW80Fnlgq9XiIXqvugHI')
    msg = '##구매대행 : \n'
    for info in diff:
        msg += '- %s.%s\n%s => %s\n' % info

    write_log(msg)
    bot.sendMessage('714653402', msg)

write_log("Done !!!")