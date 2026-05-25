from selenium import webdriver 
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import pytesseract
from PIL import Image
import re
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os
from openpyxl import load_workbook
from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image, ImageEnhance
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
wb = load_workbook(r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inboundDB.xlsx")
ws = wb["Inbound"]
supplierId = ws["B2"].value
vehicleType = ws["B3"].value
efficiency = ws["B4"].value
customerId = ws["B5"].value
vehicleType2 = ws["B6"].value
efficiency2= ws["B7"].value
wb = load_workbook(r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inboundDB.xlsx")
login = wb["URL WMS"]
url_uat = login["B2"].value
user = login["B3"].value
password = login["B4"].value



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
print("Waiting...")
def supplierUnloadingefficiency1():
    function = 'C0104_AC0104_A1011'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    time.sleep(2)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(1)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    add = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Add']").click()
    time.sleep(1)
    driver.find_element(By.NAME, "supplierId").send_keys(str(supplierId))
    driver.find_element(By.NAME, "vehicleType").send_keys(str(vehicleType))
    driver.find_element(By.NAME, "efficiency").send_keys(str(efficiency))
    #screenshot
    time.sleep(1)
    #filename = ("C:\\Users\\aphisit.sen\\Desktop\\งานโปรเจค\\DC5 Project\\selenuim bot project UNT\\inbound\\supplierUnloadingEfficiency\\ใส่suppierIdรอสร้าง.png")
    #driver.save_screenshot(filename)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ1\2ใส่suppierIdรอสร้าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    closecallpse = driver.find_element(By.XPATH, "//div[contains(@class,'layout_collapse_btn')]").click()
    #screenshot
    time.sleep(1)
    #filename = ("C:\\Users\\aphisit.sen\\Desktop\\งานโปรเจค\\DC5 Project\\selenuim bot project UNT\\inbound\\supplierUnloadingEfficiency\\สร้างsuppierId.png")
    #driver.save_screenshot(filename)
    #screenshot
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ1\3สร้างsuppierId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    #filename = ("C:\\Users\\aphisit.sen\\Desktop\\งานโปรเจค\\DC5 Project\\selenuim bot project UNT\\inbound\\supplierUnloadingEfficiency\\suppierIdสามารถสร้างได้.png")
    #driver.save_screenshot(filename)
    #print('Screenshot 1 Success')
    #ข้อ 2
    time.sleep(1)
    add = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Add']").click()
    time.sleep(1)
    driver.find_element(By.NAME, "supplierId").send_keys(str(customerId))
    driver.find_element(By.NAME, "vehicleType").send_keys(str(vehicleType))
    driver.find_element(By.NAME, "efficiency").send_keys(str(efficiency))
    #screenshot
    time.sleep(1)
    #filename = ("C:\\Users\\aphisit.sen\\Desktop\\งานโปรเจค\\DC5 Project\\selenuim bot project UNT\\inbound\\supplierUnloadingEfficiency\\ใส่suppierIdรอสร้าง.png")
    #driver.save_screenshot(filename)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ2\2ใส่suppierIdรอสร้าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    closecallpse = driver.find_element(By.XPATH, "//div[contains(@class,'layout_collapse_btn')]").click()
    #screenshot
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ2\3สร้างsuppierId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)


    #ข้อ 3
    time.sleep(1)
    add = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Add']").click()
    time.sleep(1)
    driver.find_element(By.NAME, "supplierId").send_keys(str(customerId))
    driver.find_element(By.NAME, "vehicleType").send_keys(str(vehicleType))
    driver.find_element(By.NAME, "efficiency").send_keys(str(efficiency))
    #screenshot
    time.sleep(1)
    #filename = ("C:\\Users\\aphisit.sen\\Desktop\\งานโปรเจค\\DC5 Project\\selenuim bot project UNT\\inbound\\supplierUnloadingEfficiency\\ใส่suppierIdรอสร้าง.png")
    #driver.save_screenshot(filename)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ3\2ใส่suppierIdรอสร้าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    #closecallpse = driver.find_element(By.XPATH, "//div[contains(@class,'layout_collapse_btn')]").click()
    #screenshot
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ3\3สร้างsuppierIdซ้ำไม่ได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
   
    #ข้อ4
    time.sleep(1)
    closecallpse = driver.find_element(By.ID, "layer-index-btn--0").click() #ปิดซ้ำ
    driver.find_element(By.NAME, "Sup").send_keys(str(customerId))
    driver.find_element(By.NAME, "C0104_AC0104_A1011queryForm_searchBtn").click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ4\2suppierIdมีแล้ว.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    add = driver.find_element(By.XPATH, "//div[@class='dhxtoolbar_text' and text()='Add']").click()
    time.sleep(1)
    driver.find_element(By.NAME, "supplierId").send_keys(str(customerId))
    driver.find_element(By.NAME, "vehicleType").send_keys(str(vehicleType2))
    driver.find_element(By.NAME, "efficiency").send_keys(str(efficiency))
    #screenshot
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ4\3ใส่suppierIdรอสร้าง.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\1กำลังAdd.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    closecallpse = driver.find_element(By.XPATH, "//div[contains(@class,'layout_collapse_btn')]").click()
    #screenshot
    time.sleep(2)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ4\4สร้างsuppierId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    #ข้อ 5
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\2สามารถAddได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")

    cell = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//td[normalize-space()='500']")))
    ActionChains(driver).double_click(cell).perform()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\3สามารถEditได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    time.sleep(1)
    driver.find_element(By.NAME, "efficiency").send_keys(str(efficiency2))
    time.sleep(1)
    save = driver.find_element(By.XPATH, "//button[.//i[contains(@class,'icon-save')]]").click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\4สามารถEditได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    wait = WebDriverWait(driver, 10)
    delete_btn = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[contains(@class,'dhx_toolbar_btn')][.//div[text()='Delete']]")))
    delete_btn.click()
    time.sleep(1)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\5สามารถDeleteได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")
    confirm_btn = wait.until(EC.element_to_be_clickable((By.ID,"layer-index-btn--0"))).click()
    time.sleep(2)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ5\6Deleteได้.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    driver.save_screenshot(filename)
    print("screenshot save...")


def customer():
    function = 'A1017 Customer - 1017'
    buttonnv = driver.find_element(By.ID, "menu-btn").click()
    #time.sleep(1)
    box = driver.find_element(By.XPATH,"//input[@placeholder='Function Retrieval']")
    box.clear()
    time.sleep(2)
    box.send_keys(function + Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.NAME, "customerId").send_keys(str(supplierId))
    driver.find_element(By.NAME, "A1017searchForHeader_searchBtn").click()
    #screenshot
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ1\1ไม่มีcustomerId.png"
    # แยก path โฟลเดอร์ออกมา
    folder = os.path.dirname(filename)
    # ถ้าไม่มีโฟลเดอร์สร้างให้ครบทุกชั้น
    os.makedirs(folder, exist_ok=True)
    # ค่อยเซฟรูป
    driver.save_screenshot(filename)
    print("screenshot save...")
    #ข้อ 2
    customer = driver.find_element(By.NAME, "customerId")
    customer.clear()
    customer.send_keys((str(customerId)))
    driver.find_element(By.NAME, "A1017searchForHeader_searchBtn").click()
    time.sleep(0.5)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ2\1มีcustomerId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    #time.sleep(0.5)
    driver.save_screenshot(filename)
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ3\1มีcustomerId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    #time.sleep(0.5)
    driver.save_screenshot(filename)
    print("screenshot save...")
    filename = r"C:\Users\aphisit.sen\Desktop\งานโปรเจค\DC5 Project\selenuim bot project UNT\inbound\supplierUnloadingEfficiency\ข้อ4\1มีcustomerId.png"
    folder = os.path.dirname(filename)
    os.makedirs(folder, exist_ok=True)
    #time.sleep(1)
    driver.save_screenshot(filename)
    print("screenshot save...")
    wait = WebDriverWait(driver, 10)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'ui-icon-close')]")))
    close_btn.click() #ปิด tab
#เริ่มทำงาน
time.sleep(2)
print('after',driver.current_url)
WebDriverWait(driver, 60).until(lambda d: "#/home" in d.current_url)
print('before',driver.current_url)
print("running... function 1")
customer()
print("function 1 Done")
time.sleep(1)
print("running... function 2")
supplierUnloadingefficiency1()
print("function 2 Done")
time.sleep(1)
print("****ทำงานจบเรียบร้อยครับ Code Selenuim By Ahisit Senatam(Aof)****")
driver.quit()
