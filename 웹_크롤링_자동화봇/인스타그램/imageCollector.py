from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib.request

# 파이썬으로 브라우저부터 띄워보기
driver = webdriver.Chrome('chromedriver.exe')
driver.get("https://www.instagram.com/")

time.sleep(1)
id_input = driver.find_element_by_css_selector("input[name='username']")
pw_input = driver.find_element_by_css_selector("input[name='password']")
login_btn = driver.find_element_by_css_selector("button[type='submit']")

id_input.send_keys('****')
pw_input.send_keys('****')
login_btn.click()
time.sleep(5)

# 해쉬태그 검색
tags = "영상"
driver.get(f"https://www.instagram.com/explore/tags/{tags}/")

# 첫번째 사진 클릭
driver.implicitly_wait(30)
firstContent = driver.find_element_by_css_selector("._9AhH0").click()

# 반복문 자동화
for i in range(50):
    driver.implicitly_wait(10)
    try:
        src = driver.find_element_by_css_selector(".PdwC2 .FFVAD").get_attribute('src')
        urllib.request.urlretrieve(src, f'{tags}{i + 1}.jpg')
        nextBtn = driver.find_element_by_css_selector(
            "body > div._2dDPU.CkGkG > div.EfHg9 > div > div > a._65Bje.coreSpriteRightPaginationArrow").click()
        print(f'{tags}{i + 1}.jpg 저장 완료!!')
    except:
        print(f'동영상 패스')
        nextBtn = driver.find_element_by_css_selector(
            "body > div._2dDPU.CkGkG > div.EfHg9 > div > div > a._65Bje.coreSpriteRightPaginationArrow").click()