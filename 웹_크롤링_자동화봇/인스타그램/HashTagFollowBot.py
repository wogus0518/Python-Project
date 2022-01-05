from selenium import webdriver
import time

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
tags = "コリヨジャ"
driver.get(f"https://www.instagram.com/explore/tags/{tags}/")

# 첫번째 사진 클릭
driver.implicitly_wait(30)
firstContent = driver.find_element_by_css_selector("._9AhH0").click()

# 반복문 자동화
success = 0

for i in range(500):
    try:
        driver.implicitly_wait(30)
        followBtn = driver.find_element_by_css_selector("div.bY2yH > button.sqdOP")
        followBtn_status = followBtn.text
        likeBtn = driver.find_element_by_css_selector("div.QBdPU.rrUvL > span > svg")
        nextBtn = driver.find_element_by_css_selector("body > div._2dDPU.QPGbb.CkGkG > div.EfHg9 > div > div > div.l8mY4.feth3 > button > div > span > svg")
        if followBtn_status == '팔로우':
            followBtn.click()
            print(likeBtn.text)
            likeBtn.click()
            success += 1
            print(f'현재까지 {success}명 작업 완료!!')
            time.sleep(5)
            nextBtn.click()
        else:
            print(f'이미 팔로우 하고 있습니다! 패스~')
            nextBtn.click()
    except:
        print(f'Error 발생!!')
        nextBtn.click()

# print(f'''
# ************************
# 작업이 끝났습니다!!
# 작업한 해시태그 : #{tags}
# 팔로우한 수 : {success}
# ************************
# ''')
# print(f'''
# ************************
# 작업이 끝났습니다!!
# 작업한 해시태그 : #{tags}
# 좋아요 누른 수 : {success}
# ************************
# ''')