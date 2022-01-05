from selenium import webdriver
import time
import pandas as pd
import csv

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

# 읽을 csv파일명 지정
fileName = 'flosting___언팔리스트.csv'

# # csv파일 읽어오기
f = open(fileName, 'r')
reader = csv.reader(f)
backup_list = []
follower_list = []
for row in reader:
    follower_list.append(row[0])
    backup_list.append(row[0])

def BackupList(backup_list, follower, fileName):
    backup_list.remove(follower)
    follower_df = pd.DataFrame(backup_list)
    follower_df.to_csv(fileName, index=False, header=False)

def CheckBtnStatus(followBtn, backup_list, follower, fileName):
    followBtn.click()
    time.sleep(2)
    confirmBtn = driver.find_element_by_css_selector(".-Cab_")
    confirmBtn.click()
    BackupList(backup_list, follower, fileName)

for follower in follower_list:
    driver.get(f"https://www.instagram.com/{follower}/")
    time.sleep(3)
    try:
        followBtn = driver.find_element_by_css_selector("._6VtSN")
        CheckBtnStatus(followBtn, backup_list, follower, fileName)
    except:
        None