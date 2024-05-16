from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui
import time
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# 엑셀 파일 로드
workbook = load_workbook(filename="test.xlsx")

sheet = workbook.active

header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
}

chrome_options = Options()
#브라우저 꺼짐 방지
chrome_options.add_experimental_option("detach", True)
#불필요한 에러 메시지 없애기
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"]) # 셀레니움 로그 무시

chrome_options.add_argument(r'load-extension=C:\Users\jaewon\AppData\Local\Google\Chrome Dev\User Data\Profile 1\Extensions\cgococegfcmmfcjggpgelfbjkkncclkf\1.1.9.3_0')

service = webdriver.ChromeService(executable_path='C:\Program Files\Google\Chrome Dev\Application\chrome.exe')

def remove_none_img(imgs):
        filer_img = []
        
        for img in imgs:
            if img.get_attribute("src") != None:
                filer_img.append(img)
                
        return filer_img
    
def find_imgs(tag):
    imgs = tag.find_elements(By.TAG_NAME, 'img')
    
    filer_img = []
    
    for img in imgs:
        if img.get_attribute("src") != None:
            filer_img.append(img)
    
    return filer_img

def find_imgs_container(img_elements):
    
    for img_tag in img_elements:
        parent_tag = img_tag.find_element(By.XPATH, '..')
        parent_tag_imgs = find_imgs(parent_tag)
        
        while len(parent_tag_imgs) < len(imgs):
            parent_tag = parent_tag.find_element(By.XPATH, '..')
            parent_tag_imgs = find_imgs(parent_tag)
        
        return parent_tag.find_elements(By.XPATH, './*')
    


if __name__ == '__main__':
    
    
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        data_start_row = 1
        
        url = ""
        attribute = ""
        attribute_value = ""
        if i > data_start_row:
            print("Row", i, ":", row)
            
            url = row[2]
            
            if row[0]:
                attribute = "id"
                attribute_value = row[0]
            else:
                attribute = "class"
                attribute_value = row[1]

            b = webdriver.Chrome(options=chrome_options)
            
            b.get(f"{url}")

            containers = b.find_elements(By.XPATH, f"//*[contains(@{attribute},'{attribute_value}')]")

            for tag in containers:
                try:
                    imgs = find_imgs(tag)
                    imgs_container = find_imgs_container(imgs)
                    
                    list_obj = []
                    
                    for tag in imgs_container:
                        print (tag.text)
                        obj = {}
                        
                        obj["url"] = ""
                        obj["element"] = tag.find_element(By.TAG_NAME, 'img')
                        
                        try:
                            a = tag.find_element(By.TAG_NAME, 'a')
                            obj["url"] = a.get_attribute("href")
                            
                        except Exception as e:
                            print (f"{tag.text} 상품 a 태그 없음")

                        list_obj.append(obj)


                    # ==================== 아이템 페이지로 이동 후 자료 수집 ====================
                    
                    for product in list_obj:
                        
                        # 상품 페이지로 이동
                        if product["url"] == "":
                            b.execute_script("arguments[0].click();", product["element"])
                        
                        else:
                            b.get(f"{product["url"]}")
                        
                        
                        
                        # 사이즈 태그 리스트 container 찾기
                        all_elements = b.find_elements(By.XPATH, "//*[contains(text(), '95') or contains(text(), '100') or contains(text(), 'x') or contains(text(), 'xs')]")

                        # 2개 이상의 단어를 포함하는 요소 필터링
                        matching_elements = []
                        for element in all_elements:
                            text = element.text
                            if text.count('95') + text.count('100') + text.count('x') + text.count('xs') >= 2:
                                matching_elements.append(element)

                        # ifram_tag = b.find_elements(By.CLASS_NAME, "revenue")
                        # b.switch_to_frame(ifram_tag)
                        
                        
                        

                        # # sellerlife-keyword-analysis-frame

                        # price_list = b.find_elements(By.XPATH, "//[@class='revenue']")

                        # for price in price_list:
                        #     print(price.text)
                    
                    
                        # 작업 완료 후 이전페이지가기
                        b.back()
                      
                        
                        
                    
                except Exception as e:
                    print(e)
                    
                    
                    
            # WebDriver를 닫습니다.
            b.quit()









# wb.save(f"test.xlsx")