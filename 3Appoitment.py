from selenium import webdriver 
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import pytesseract
import traceback
from PIL import Image
import re
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os
from openpyxl import load_workbook
from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image, ImageEnhance
from selenium.common.exceptions import WebDriverException, TimeoutException ,NoSuchElementException
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
wb = load_workbook(r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inboundDB.xlsx")
ws = wb["Inbound"]
PO_create1 = ws["B22"].value
PO_close = ws["B23"].value
PO_cancel = ws["B24"].value
PO_create2 = ws["B25"].value
startTime = ws["B26"].value
endTime = ws["B27"].value
date_Current = ws["B28"].value
dockNo_Avaliable = ws["B29"].value
date_Past = ws["B30"].value
dockNo_Inuse = ws["B31"].value
ArriveTime1_past = ws["B32"].value
ArriveTime2_past = ws["B33"].value
ArriveTime1_current = ws["B34"].value
ArriveTime2_current = ws["B35"].value
PO_create3 = ws["B36"].value
PO_create4 = ws["B37"].value
card_check = ws["B38"].value
card_notcheck = ws["B39"].value
expect_arrive_time = ws["B40"].value
owner_po = ws["B41"].value


wb = load_workbook(r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inboundDB.xlsx")
login = wb["URL WMS"]
url_uat = login["B2"].value
user = login["B3"].value
password = login["B4"].value
url_rf_uat = login["B5"].value





driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)
driver.get(url_uat)
driver.find_element(By.ID,"details-button").click()
driver.find_element(By.ID,"proceed-link").click()
driver.find_element(By.ID,"CURUSERID").send_keys(user)
driver.find_element(By.ID,"CURURPD").send_keys(password)
time.sleep(1)
len_cap = 0
loop_count = 0
MAX_LOOP = 10

while loop_count < MAX_LOOP:
    print(f"\n=== LOOP {loop_count + 1} ===")

    try:
        img = driver.find_element(By.ID, "checknumImg")
        time.sleep(0.5)
        img.screenshot(
            r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\capthca\capthca.png"
        )

        image = Image.open(
            r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\capthca\capthca.png"
        )
        image = image.convert("L")
        image = ImageEnhance.Contrast(image).enhance(3.0)
        image = image.point(lambda x: 0 if x < 140 else 255, '1')

        captcha = pytesseract.image_to_string(
            image,
            config="--psm 7 --oem 1 -c tessedit_char_whitelist=0123456789"
        )

        Revisecaptcha = re.sub(r"\D", "", captcha)
        len_cap = len(Revisecaptcha)

        print("captcha OCR =", Revisecaptcha, "| len =", len_cap)

        if len_cap != 4:
            print("OCR not 4 ")
            driver.execute_script("arguments[0].click();", img)
            loop_count += 1
            continue

        inputbox = driver.find_element(By.ID, "CHECKNUM")
        inputbox.clear()
        inputbox.send_keys(Revisecaptcha)

        driver.find_element(By.ID, "submit").click()
        time.sleep(1)


        errors = driver.find_elements(By.ID, "login_validationMessage")

        if errors and errors[0].is_displayed():
            print("wrong captcha ")
            driver.execute_script("arguments[0].click();", img)
            loop_count += 1
            continue

        print("captcha passing ")
        break

    except Exception as e:
        print("ERROR Loop:", e)
        loop_count += 1

if loop_count >= MAX_LOOP:
    print("loop is max try again")
    driver.quit()

btn = driver.find_element(By.XPATH, "//input[@value='Force Login']")
driver.execute_script("arguments[0].click();", btn)
#funtion 

def createapp1():
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create1)
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ2\1Poยังไม่ได้สร้างAppointment.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab

    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    time.sleep(3)
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.send_keys(PO_create1)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ2\2กดoperation.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ5\1po.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "udf06")))
    POWMS.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ8\1คำนวณเวลาPo.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ20\1dockcapacityน้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ24\1คำนวนเวลาประตูถูกต้อง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create1)
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ2\3สร้างappได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ5\2สร้างappได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ8\1คำนวณเวลาประตูapp.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ20\2ได้1dock.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ24\1คำนวนเวลาประตูถูกต้อง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    detailapp = driver.find_element(By.XPATH,"//div[contains(@class,'dhxtabbar_tab_text') and normalize-space()='Details']")
    driver.execute_script("arguments[0].click();", detailapp)  #เปิด Details dock open time
    wait = WebDriverWait(driver, 20)
    time.sleep(1)
    add_btn = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH,"//div[@id='A0588D2_cell']//div[contains(@class,'dhx_toolbar_btn')]""[.//div[@class='dhxtoolbar_text' and normalize-space()='Add']]"))) #ปุ่ม add
    add_btn.click()
    time.sleep(1)
    add_PO = ActionChains(driver)
    add_PO.send_keys(PO_create2)
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ6\1สามมารถaddpoได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    add_PO.send_keys(Keys.TAB)
    add_PO.send_keys(Keys.ENTER)
    add_PO.perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ6\2addpoได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ7\1poที่add.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    time.sleep(0.5)
    #ปิดไม่ได้
    time.sleep(3)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab
    time.sleep(1)

def createapp2():
    wait = WebDriverWait(driver, 20)
    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.send_keys(PO_create1)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ3\1สร้างappที่สร้างไปแล้วไม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    confirm = driver.find_element(By.ID, "layer-index-btn--0").click() #confirm duplicated
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    close_PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    close_PO.clear()
    close_PO.send_keys(PO_close) #po close
    time.sleep(1)
    PO_SERCH = ActionChains(driver)
    PO_SERCH.send_keys(Keys.ENTER)
    PO_SERCH.perform()
    time.sleep(1)
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ4\1สร้างappสถานะcloseไม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    confirm = driver.find_element(By.ID, "layer-index-btn--0").click() #confirm duplicated
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    cancel_PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    cancel_PO.clear()
    cancel_PO.send_keys(PO_cancel) #po cancel
    time.sleep(1)
    PO_SERCH = ActionChains(driver)
    PO_SERCH.send_keys(Keys.ENTER)
    PO_SERCH.perform()
    time.sleep(1)
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ4\2สร้างappสถานะcancelไม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    confirm = driver.find_element(By.ID, "layer-index-btn--0").click() #confirm duplicated
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO2status  = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO2status.clear()
    PO2status.send_keys(PO_create1,",",PO_create2) #po 2 status
    time.sleep(1)
    PO_SERCH = ActionChains(driver)
    PO_SERCH.send_keys(Keys.ENTER)
    PO_SERCH.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ7\2POสถานะไม่เปลี่ยน.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab


def appTime():
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create1)
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    startTimewms = wait.until(EC.element_to_be_clickable((By.NAME, "startTime")))
    startTimewms.clear()
    startTimewms.send_keys(startTime)
    endTimewms = wait.until(EC.element_to_be_clickable((By.NAME, "endTime")))
    endTimewms.clear()
    endTimewms.send_keys(endTime)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ9\1กำหนดเวลาstartน้อยกว่าendtimeได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    time.sleep(1)
    AriveTime_current= wait.until(EC.element_to_be_clickable((By.NAME, "appointmentDate")))
    AriveTime_current.clear()
    AriveTime_current.send_keys(date_Current)
    dockNo_avl = wait.until(EC.element_to_be_clickable((By.NAME, "dockNo")))
    time.sleep(1)
    dockNo_avl.clear()
    dockNo_avl.send_keys(dockNo_Avaliable)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ12\1กำหนดวันเวลาและประตูนัดหมายที่ว่าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    AriveTime_past = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentDate")))
    AriveTime_past.clear()
    AriveTime_past.send_keys(date_Past)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ11\1กำหนดวันเวลาที่นัดหมายที่ผ่านมาแล้ว.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    dockNoinuse = wait.until(EC.element_to_be_clickable((By.NAME, "dockNo")))
    time.sleep(1)
    dockNoinuse.clear()
    dockNoinuse.send_keys(dockNo_Inuse)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ13\1กำหนดวันเวลาและประตูนัดหมายที่ไม่ว่าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    startTimewms = wait.until(EC.element_to_be_clickable((By.NAME, "startTime")))
    startTimewms.clear()
    startTimewms.send_keys(endTime)
    endTimewms = wait.until(EC.element_to_be_clickable((By.NAME, "endTime")))
    endTimewms.clear()
    endTimewms.send_keys(startTime)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ10\1กำหนดเวลาstartมากกกว่าendtimeไม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(1)
    cell = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//td[normalize-space()='{PO_create1}']")))
    ActionChains(driver).context_click(cell).perform()
    ap_cell = cell.find_element(
    By.XPATH, "./ancestor::tr//td[starts-with(normalize-space(),'AP')]")
    ap_value = ap_cell.text


    cancel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_cancelDocument') and contains(@class,'sub_item')]")))
    driver.execute_script("arguments[0].click();", cancel) #cancel appointment
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ17\1ทำการCancelAppoint.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ16\2ทำการCancelAppoint.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ17\2ทำการCancelAppointสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ16\3ทำการCancelAppointสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab


    function = 'A0590 Admission registration - 0590'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    time.sleep(1)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    ArrNo = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentNo")))
    ArrNo.send_keys(ap_value)
    ArrSerch = ActionChains(driver)
    ArrSerch.send_keys(Keys.ENTER)
    ArrSerch.perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ16\1ยังไม่ได้สร้างARR.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิดแทบ


def appTime2():
    wait = WebDriverWait(driver, 20)
    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.send_keys(PO_create1)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    expectedArriveTime1_past = wait.until(EC.element_to_be_clickable((By.NAME, "expectedArriveTime1")))
    expectedArriveTime1_past.clear()
    expectedArriveTime1_past.send_keys(ArriveTime1_past)
    expectedArriveTime2_past = wait.until(EC.element_to_be_clickable((By.NAME, "expectedArriveTime2")))
    expectedArriveTime2_past.clear()
    expectedArriveTime2_past.send_keys(ArriveTime2_past)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ11\2กำหนดวันที่ผ่านมาแล้วที่PO.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ11\3ไม่สามารถสร้างได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ26\1ไม่สามารถสร้างได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm = driver.find_element(By.ID, "layer-index-btn--0").click() 
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    expectedArriveTime1_current = wait.until(EC.element_to_be_clickable((By.NAME, "expectedArriveTime1")))
    expectedArriveTime1_current.clear()
    expectedArriveTime1_current.send_keys(ArriveTime1_current)
    expectedArriveTime2_current = wait.until(EC.element_to_be_clickable((By.NAME, "expectedArriveTime2")))
    expectedArriveTime2_current.clear()
    expectedArriveTime2_current.send_keys(ArriveTime2_current)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ17\3สร้างAppointใหม่.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ17\4สร้างAppointใหม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ22\1สร้างAppointแบบข้ามวัน.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิดแทบ
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create1)
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ22\2ประตูคำนวณเวลาลงข้ามวันถูก.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิดแทบ



def appTime3():
    wait = WebDriverWait(driver, 20)
    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.send_keys(PO_create3)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create3}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "udf06"))) #ดู carton quantity
    POWMS.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ21\1POมีcqมากกว่าdock.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ25\1POลง2ประตู.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ27\1POใส่cqมากกว่าของที่สั่งจริง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create3}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click()

    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create3)
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ21\2POสร้าง2App.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ25\2ระบบคำนวณเวลาลงประตูถูกต้อง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ27\2ระบบเจนเวลาลงสินค้าและประตูมาให้เกินเวลาที่สินค้าใช้จริง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ29\1POเจน2app.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")



    dock_elements = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//td[starts-with(normalize-space(),'DOCK')]")))
    dock_text = ",".join(d.text.strip() for d in dock_elements)


    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() 
    
    function = 'A1010 Dock - 1010'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    dock_search = wait.until(EC.element_to_be_clickable((By.NAME, "DOCK")))
    dock_search.send_keys(dock_text)
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    row = wait.until(EC.element_to_be_clickable((By.XPATH,"(//tr[contains(@class,'ev_material') ""and .//td[starts-with(normalize-space(),'DOCK')]])[1]"))) #double click row แรก
    ActionChains(driver).double_click(row).perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ21\3dockมีcapacityน้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    row_down = ActionChains(driver)
    row_down.send_keys(Keys.ARROW_DOWN)
    row_down.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ21\4dockมีcapacityน้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() 


    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create3)
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
  
    cell = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//td[normalize-space()='{PO_create3}']")))
    ActionChains(driver).context_click(cell).perform() #คลิ๊กขวารอ cancel
    cancel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_cancelDocument') and contains(@class,'sub_item')]")))
    driver.execute_script("arguments[0].click();", cancel) #cancel appointment
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ29\2POสร้าง2appcancel1.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ29\3appcancel1เหลือ1.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() 
    wait = WebDriverWait(driver, 20)

    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.send_keys(PO_create3)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ29\4POสถานะเป็นcreate.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 

def appTime4():
    wait = WebDriverWait(driver, 20)
    #เปิด PO
    function = 'A2001 PO - 2001'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.clear()
    PO.send_keys(PO_create1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    SKU_elements1 = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//td[starts-with(normalize-space(),'0000000')]"))) #เก็บ SKU ใน PO
    SKU_PO1 = ",".join(d.text.strip() for d in SKU_elements1)
    print("สินค้า PO1", SKU_PO1)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ28\1POให้ดูSKU.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    #23 แคป detail ให้เห็น zone 
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()


    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A2001']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    time.sleep(1)
    el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@id='a_cell_cont']//div[contains(@class,'dhxform_img') and @title='Show Close/Cancel PO']")))
    driver.execute_script("""arguments[0].dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));""", el) #ติ๊ก Close/cancel
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ serch
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "poNo")))
    PO.clear()
    PO.send_keys(PO_create4)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create4}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    SKU_elements = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//td[starts-with(normalize-space(),'0000000')]"))) #เก็บ SKU ใน PO
    SKU_text = ",".join(d.text.strip() for d in SKU_elements)
    print("สินค้า", SKU_text)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ23\1POมีสินค้าหลายโซน.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{PO_create4}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()






    operation = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Operation']").click() #เปิด operation
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    auto_assign = wait.until(EC.element_to_be_clickable(( By.XPATH, "//h1[normalize-space()='Auto Assign Dock']"))) #auto assign doc
    driver.execute_script("arguments[0].click();", auto_assign) 
    time.sleep(1)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 

    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create4)
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    Dock_elements = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//td[starts-with(normalize-space(),'DOCK')]"))) #เก็บประตู
    dock_text = ",".join(d.text.strip() for d in Dock_elements)
    print('ประตู'  ,dock_text)
    #แก้ double click
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{dock_text}']"))) #double click list detail PO
    ActionChains(driver).context_click(cell).perform()
    #เก็บค่า appoint po_create4
    app_element = cell.find_element(
    By.XPATH, "./ancestor::tr//td[starts-with(normalize-space(),'AP')]")
    app_po4 = app_element.text
    print("appont po 4 is ",app_po4)
    #เก็บค่า appoint po_TDSC
    app_element2 = cell.find_element(
    By.XPATH, "./ancestor::tr//td[starts-with(normalize-space(),'TDSC_')]")
    po_tdsc = app_element2.text





    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ23\4appointลงประตูที่โซนสินค้าเยอะสุด.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.send_keys(PO_create1)
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    Dock_elements2 = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//td[starts-with(normalize-space(),'DOCK')]"))) #เก็บประตู
    dock_text2 = ",".join(d.text.strip() for d in Dock_elements2)
    print('ประตู'  ,dock_text2)

   #แก้ double click
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{dock_text2}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    app_element = cell.find_element(
    By.XPATH, "./ancestor::tr//td[starts-with(normalize-space(),'AP')]")
    app_po1 = app_element.text
    print("appont po 1 is ",app_po1)



    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ28\2appointไปประตูโซนเดียวกับบ้านหยิบ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #แคปประตู

    #23 แคปประตู
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 

    function = 'A1010 Dock - 1010'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    dock_search = wait.until(EC.element_to_be_clickable((By.NAME, "DOCK")))
    dock_search.send_keys(dock_text)
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ23\3ลงประตูของโซนที่จำนวนเยอะสุด.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A1010']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    dock_search = wait.until(EC.element_to_be_clickable((By.NAME, "DOCK")))
    dock_search.send_keys(dock_text2)
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    time.sleep(0.5)   
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ28\3ประตูโซนเดียวกับบ้านหยิบ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 

    

    function = 'C0104_ACHECKMASTER report  check master'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    dock_search = wait.until(EC.element_to_be_clickable((By.NAME, "sku")))
    dock_search.send_keys(SKU_text)
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ23\2โซนของSKU.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 

    function = 'A1028 SKU - 1028'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    dock_search = wait.until(EC.element_to_be_clickable((By.NAME, "skuQ")))
    dock_search.send_keys(SKU_PO1)
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    time.sleep(0.5)
    cell = wait.until(EC.presence_of_element_located((By.XPATH,f"//td[normalize-space()='{SKU_PO1}']"))) #double click list detail PO
    ActionChains(driver).double_click(cell).perform()
    forwardingloc = driver.find_element(By.XPATH,"//div[contains(@class,'dhxtabbar_tab_text') and normalize-space()='Forwarding LOC']")
    driver.execute_script("arguments[0].click();", forwardingloc)  #เปิด Details dock open time
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ28\4โซนบ้านหยิบ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 
    return app_po1,app_po4



def arr1(app_po1,app_po4):    
    print("return app_po1",app_po1)
    print("return app_po4",app_po4)
    time.sleep(1)
    function = 'A0590 Admission registration - 0590'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(0.5)
    #เปิดแทบ arr
    wait = WebDriverWait(driver, 20)
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentNo")))
    PO.send_keys(app_po4)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'obj')]/tbody"))) #คลิ๊กขวาหน้าว่าง
    ActionChains(driver).context_click(tbody).perform()
    withdrawal = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_AbstractAppArrival') and contains(@class,'sub_item')]" )))
    withdrawal.click()
    appointment = WebDriverWait(driver,10).until(
    EC.element_to_be_clickable((By.XPATH,"//form[@id='A0590A05DIG90_APPARRIVALsearchDialogForm']//input[@name='appointmentNo']"))) #กรอก app ที่ arr
    appointment.clear()
    appointment.send_keys(app_po4)
    time.sleep(0.5)
    enter = ActionChains(driver)
    enter.send_keys(Keys.ENTER)
    enter.perform()
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    confirmbtn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Confirm']"))) #กดปุ่ม confrim create arr
    confirmbtn.click() #สร้าง arr ทำ check in
    #ใส่การ์ด
    cell = wait.until( EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'ARR')]")))
    ActionChains(driver).double_click(cell).perform() #double click arr
    time.sleep(0.5)
    addCard_btn = wait.until(EC.element_to_be_clickable((By.ID,"tplCheck")))
    addCard_btn.click()
    time.sleep(0.5)
    card_serial = wait.until(EC.element_to_be_clickable((By.NAME,"cardSerialNo")))
    card_serial.clear()
    card_serial.send_keys(card_check)
    card_serial = ActionChains(driver)
    card_serial.send_keys(Keys.ENTER)
    card_serial.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\1ใส่การ์ด.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@id='A0590documentsIssued']//button[normalize-space()='Confirm']"))) #ปุ่ม confirm card
    confirm.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\2ใส่การ์ดเรียบร้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0590']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    time.sleep(1)
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(0.5)
    #เปิดแทบ arr
    wait = WebDriverWait(driver, 20)
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentNo")))
    PO.send_keys(app_po1)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'obj')]/tbody"))) #คลิ๊กขวาหน้าว่าง
    ActionChains(driver).context_click(tbody).perform()
    withdrawal = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_AbstractAppArrival') and contains(@class,'sub_item')]" )))
    withdrawal.click()
    appointment = WebDriverWait(driver,10).until(
    EC.element_to_be_clickable((By.XPATH,"//form[@id='A0590A05DIG90_APPARRIVALsearchDialogForm']//input[@name='appointmentNo']"))) #กรอก app ที่ arr
    appointment.clear()
    appointment.send_keys(app_po1)
    time.sleep(0.5)
    enter = ActionChains(driver)
    enter.send_keys(Keys.ENTER)
    enter.perform()
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    confirmbtn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Confirm']"))) #กดปุ่ม confrim create arr
    confirmbtn.click() #สร้าง arr ทำ check in
    #ใส่การ์ด
    cell = wait.until( EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'ARR')]")))
    ActionChains(driver).double_click(cell).perform() #double click arr
    time.sleep(0.5)
    addCard_btn = wait.until(EC.element_to_be_clickable((By.ID,"tplCheck")))
    addCard_btn.click()
    time.sleep(0.5)
    card_serial = wait.until(EC.element_to_be_clickable((By.NAME,"cardSerialNo")))
    card_serial.clear()
    card_serial.send_keys(card_notcheck)
    card_serial = ActionChains(driver)
    card_serial.send_keys(Keys.ENTER)
    card_serial.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ18\1ใส่การ์ด.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@id='A0590documentsIssued']//button[normalize-space()='Confirm']"))) #ปุ่ม confirm card
    confirm.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ18\2ใส่การ์ดเรียบร้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]"))) #ปิด tab
    close_btn.click() 






def newtab2old():
    old_tab = driver.current_window_handle   # จำ tab ปัจจุบัน
    # เปิด tab ใหม่
    driver.switch_to.new_window('tab')
    # เข้าเว็บ
    driver.get(url_rf_uat)
    ActionChains(driver).send_keys("thisisunsafe").perform()
    time.sleep(3)
    driver.find_element(By.ID,"but_Setting_1").click()
    time.sleep(1)
    driver.find_element(By.ID,"but_SetLanguage_3").click()
    time.sleep(1)
    driver.find_element(By.ID,"but_EN_5").click()
    time.sleep(1)
    driver.find_element(By.ID,"but_Cancel_3").click()
    time.sleep(1)
    driver.find_element(By.ID,"txt_UserId_1").send_keys(user)
    driver.find_element(By.ID,"txt_UserPassword_1").send_keys(password)

    time.sleep(1)
    captcha_folder = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\capthcaRF"
    captcha_file = os.path.join(captcha_folder, "capthca.png")

    # สร้าง folder ถ้ายังไม่มี

    if not os.path.exists(captcha_folder):
        os.makedirs(captcha_folder)

    len_cap = 0
    loop_count = 0
    MAX_LOOP = 25

    while loop_count < MAX_LOOP:
        print(f"\n=== LOOP {loop_count + 1} ===")

        try:
            img = driver.find_element(By.ID, "img_CheckNum_1")
            time.sleep(0.5)

            img.screenshot(captcha_file)

            image = Image.open(captcha_file)
            image = image.convert("L")
            image = ImageEnhance.Contrast(image).enhance(3.0)
            image = image.point(lambda x: 0 if x < 140 else 255, '1')

            captcha = pytesseract.image_to_string(
                image,
                config="--psm 7 --oem 1 -c tessedit_char_whitelist=0123456789"
            )

            Revisecaptcha = re.sub(r"\D", "", captcha).strip()
            len_cap = len(Revisecaptcha)

            print("captcha OCR =", Revisecaptcha, "| len =", len_cap)

            # ===== OCR ไม่ครบ 4 =====
            if len_cap != 4:
                print("OCR not 4 → refresh captcha")
                driver.execute_script("arguments[0].click();", img)
                time.sleep(1.5)
                loop_count += 1
                continue

            inputbox = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "txt_VerificationCode_1"))
            )

            inputbox.clear()
            inputbox.send_keys(Revisecaptcha)

            driver.find_element(By.ID, "but_Login_1").click()
            time.sleep(2)

            errors = driver.find_elements(
                By.XPATH, "//*[contains(text(),'Invalid Verification Code')]"
            )

            if any(e.is_displayed() for e in errors):
                print("wrong captcha")
                driver.execute_script("arguments[0].click();", img)
                time.sleep(1.5)
                loop_count += 1
                continue

            # ===== captcha ผ่าน =====
            print("captcha passing → exit loop")
            break

        except (NoSuchElementException, TimeoutException) as e:
            print("element not ready → retry loop", e)
            loop_count += 1
            continue

        except Exception as e:
            print("unexpected error → retry loop", e)
            loop_count += 1
            continue

    else:
        raise Exception("Captcha failed after MAX_LOOP")
    time.sleep(1)
    driver.find_element(By.ID,"popup_Confirm_Button_1").click() #กดคอนเฟิร์ม RF login
    time.sleep(1)
    enter_confirm = ActionChains(driver)
    enter_confirm.send_keys(Keys.ENTER)
    enter_confirm.perform()
    time.sleep(1)
    inbound = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[text()='Inbound']"))) #กดเปิด inbound
    driver.execute_script("arguments[0].click();", inbound)
    time.sleep(0.5)
    driver.find_element(By.XPATH,"//a[contains(text(),'CheckIn')]").click() #กดฟังก์ชัน checkIn
    time.sleep(2)
    checkin_rf = driver.find_element(By.ID,"TXT_PASSCARD_1")
    checkin_rf.clear()
    checkin_rf.send_keys(card_check)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\3รอcheck.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    enter_confirm2 = ActionChains(driver)
    enter_confirm2.send_keys(Keys.ENTER)
    enter_confirm2.perform()
    #RF ต้องชี้ปุ่มก่อนแล้วส่ง enter
    wait = WebDriverWait(driver, 20)
    confirm = wait.until(EC.presence_of_element_located((By.ID, "BUT_CONFIRM_1"))) 
    driver.execute_script("arguments[0].focus();", confirm)
    confirm.send_keys(Keys.ENTER)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\4checkInเรียบร้อย.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    driver.close()
    # กลับ tab เดิม
    driver.switch_to.window(old_tab)

def arr2(app_po1,app_po4):
    function = 'A0590 Admission registration - 0590'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    time.sleep(1)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    time.sleep(2)
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ arr
    wait = WebDriverWait(driver, 20)
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentNo")))
    PO.send_keys(app_po4)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab
    
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "APPNO")))
    POWMS.clear()
    POWMS.send_keys(app_po4 + Keys.TAB + "Y" + Keys.ENTER)
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(1)    
    cell = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//td[normalize-space()='{app_po4}']")))
    ActionChains(driver).context_click(cell).perform() #คลิ๊กขวารอ cancel
    cancel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_cancelDocument') and contains(@class,'sub_item')]")))
    driver.execute_script("arguments[0].click();", cancel) #cancel appointment
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\5ทำการcancel.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ19\6cancelสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    time.sleep(1)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab


    function = 'A0590 Admission registration - 0590'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    time.sleep(2)
    btn = driver.find_element(By.CSS_SELECTOR,"div.moreSearchBtn div.btn.btn-raised.btn-default") #ปุ่ม serch หน้าแรก
    ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
    time.sleep(1)
    #เปิดแทบ arr
    wait = WebDriverWait(driver, 20)
    header = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'dhx_cell_hdr_text') and .//span[text()='Query Condition']]")))
    ActionChains(driver).move_to_element(header).click().perform() 
    PO = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentNo")))
    PO.send_keys(app_po1)
    time.sleep(1)
    POSERCH = ActionChains(driver)
    POSERCH.send_keys(Keys.ENTER)
    POSERCH.perform()
    time.sleep(1)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab

    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "APPNO")))
    POWMS.clear()
    POWMS.send_keys(app_po1 + Keys.TAB + "Y" + Keys.ENTER)

    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(1)    
    cell = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//td[normalize-space()='{app_po1}']")))
    ActionChains(driver).context_click(cell).perform() #คลิ๊กขวารอ cancel
    cancel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//tr[contains(@id,'_cancelDocument') and contains(@class,'sub_item')]")))
    driver.execute_script("arguments[0].click();", cancel) #cancel appointment
    time.sleep(1)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ18\3ทำการcancel.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ18\4cancelสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab


def createapp0(app_po1,app_po4):
    function = 'A0588 Appointment Form - 0588'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(3)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "docNo")))
    POWMS.clear()
    POWMS.send_keys(PO_create3 + Keys.TAB+Keys.TAB+Keys.TAB +Keys.TAB   + "Y" + Keys.ENTER)
    #ปุ่ม serch
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\1serch_po_wms.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    tdsc_el = wait.until(EC.presence_of_element_located((By.XPATH, "//td[starts-with(normalize-space(.), 'TDSC_')]"))) #เก็บค่า PO_TDSC
    TDSC_po4 = tdsc_el.text.strip()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\2serch_po_wmsสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "SAPPO")))
    POWMS.clear()
    POWMS.send_keys(TDSC_po4 + Keys.TAB +Keys.TAB +Keys.TAB + "Y" + Keys.ENTER)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\3serch_po_tdsc.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\4serch_po_tdscสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "appointmentDate_startappointmentDate_end")))
    POWMS.clear()
    POWMS.send_keys(expect_arrive_time + Keys.TAB +Keys.TAB + "Y" + Keys.ENTER)
    #ปุ่ม serch
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\5serch_expectTime.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\6serch_expectTimeสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    
    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "APPNO")))
    POWMS.clear()
    POWMS.send_keys(app_po4 + Keys.TAB + "Y" + Keys.ENTER)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\7serch_App.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\8serch_Appสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "APPNO")))
    POWMS.clear()
    POWMS.send_keys(Keys.TAB + "Y" + Keys.ENTER)
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\9serch_status.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    #ปุ่ม serch
    search_btn = wait.until(EC.element_to_be_clickable((By.NAME, "A0588searchForHeader_searchBtn")))
    search_btn.click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\10serch_statusสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


    #คลิ๊กขวาที่ tab A2001 
    actions = ActionChains(driver)
    tab = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@href='#A0588']")))
    actions.move_to_element(tab).context_click().perform()
    refresh = wait.until(EC.visibility_of_element_located((By.XPATH,"//div[starts-with(@id,'polygon_dhxId')]//tr[contains(@id,'_refresh')]")))
    actions.move_to_element(refresh).pause(0.2).click().perform() 
    #คลิ๊กขวารีเฟรช ^^^^^^^^

    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "APPNO")))
    POWMS.clear()
    POWMS.send_keys(Keys.TAB + "Y")
    wait = WebDriverWait(driver, 20)
    POWMS = wait.until(EC.element_to_be_clickable((By.NAME, "hedi04")))
    POWMS.clear()
    POWMS.send_keys(owner_po)
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\11serch_owner_po.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    #ปุ่ม serch
    search_btn = ActionChains(driver)
    search_btn.send_keys(Keys.ENTER)
    search_btn.perform()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\Appointment\ข้อ1\12serch_owner_poสำเร็จ.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)



























   






    


























    



    

    




    










    

























































































   









    time.sleep(5)



    





    

















#end of function
time.sleep(2)
print('after',driver.current_url)
WebDriverWait(driver, 60).until(lambda d: "#/home" in d.current_url)
print('before',driver.current_url)

functions = [ 
    createapp1,
    createapp2,
    appTime,
    appTime2,
    appTime3,
    appTime4,
    arr1,
    newtab2old,
    arr2,
    createapp0,
]

app_po1 = None
app_po4 = None


for i, func in enumerate(functions, start=1):
    print(f"\n=== running... function {i}: {func.__name__} ===")

    try:
        # 🔹 ลองเรียกแบบส่งพารามิเตอร์ก่อน
        try:
            result = func(app_po1, app_po4)
        except TypeError as e:
            # ถ้า error เพราะรับ arg ไม่ได้ → เรียกใหม่แบบไม่ส่ง
            if "positional arguments" in str(e) or "given" in str(e):
                result = func()
            else:
                raise

        # 🔹 ถ้ามี return → เก็บ
        if result is not None:
            if isinstance(result, tuple) and len(result) == 2:
                app_po1, app_po4 = result
                print(" ได้ค่าใหม่ (app_po1, app_po4)")
            else:
                app_po4 = result
                print(" ได้ค่าใหม่ (app_po4)")

        print(" done")

    except (TimeoutException, WebDriverException) as e:
        print("\n Selenium Error")
        print(f"loop     : {i}")
        print(f"function : {func.__name__}")
        print(f"type     : {type(e).__name__}")
        print(f"message  : {e}")

        tb = traceback.extract_tb(e.__traceback__)
        last = tb[-1]
        print(f"file     : {last.filename}")
        print(f"line     : {last.lineno}")
        print(f"code     : {last.line}")
        break

    except Exception as e:
        print("\n Logic / Python Error")
        print(f"loop     : {i}")
        print(f"function : {func.__name__}")
        print(f"type     : {type(e).__name__}")
        print(f"message  : {e}")

        tb = traceback.extract_tb(e.__traceback__)
        last = tb[-1]
        print(f"file     : {last.filename}")
        print(f"line     : {last.lineno}")
        print(f"code     : {last.line}")
        break

    time.sleep(1)

time.sleep(1)
print("****ทำงานจบเรียบร้อยครับ Code Selenuim By Ahisit Senatam(Aof)****")
driver.quit()