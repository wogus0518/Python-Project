import requests
import json
from selenium import webdriver
import time
import pandas as pd
import csv

# 파이썬으로 브라우저부터 띄워보기
driver = webdriver.Chrome('chromedriver.exe')
driver.get("https://www.instagram.com/")
time.sleep(1)

# 로그인하기
id_input = driver.find_element_by_css_selector("input[name='username']")
pw_input = driver.find_element_by_css_selector("input[name='password']")
login_btn = driver.find_element_by_css_selector("button[type='submit']")

id_input.send_keys('****')
pw_input.send_keys('****')

login_btn.click()
time.sleep(5)

# 검색 대상 user_id 가져오기
userName = "8784_24"
cookie = 'mid=YRZ-0QALAAETcgnxE3ySRIqpDRag; ig_did=420234A7-76C2-408F-ABC5-9DFD238BF661; ig_nrcb=1; fbm_124024574287414=base_domain=.instagram.com; shbid="2629\05446518722167\0541663846868:01f7b8fb70d1ebd5db177f500a0767006453b431d3c0a046aa6170959cad7102a68fbaf3"; shbts="1632310868\05446518722167\0541663846868:01f7c9f5db7aec475d2d381351c268224e881b803fc9537d85de05df84804904f549344e"; ds_user_id=45678201261; fbsr_124024574287414=zS9loeYsHNEhgqLpig8ehMwhEk2l-gvRlfKIoEaF710.eyJ1c2VyX2lkIjoiMTAwMDAzOTI0MjEyNDUxIiwiY29kZSI6IkFRQ1ctV0FnVlZ0clY0OGtfT2hMWW1TZGFMVC11OExVdDZZcFRBN0hQd3dmc1FiNDVPX0Z5eXVjZEo0SlBUX1RiRU9Dc3MybmlaSzNRU1ZZZlc1eUFMM3VnVmxrTnprMFZBR1g4SE9HMHNjMmpPUjV0c1lNSlBzaERsMjJ6NDZhSWFFTmt5V0VkdThPdmpVT3VHV0tDVW1CQy0wNWFlWkREUjVnSjVmOXNZOWRDS2dfWnp3dThoZ1BBTWt1VkRoVkpjSC05aFhMODhYMWRKWFBlamo1LW5jM0V6TnBEM2JHWVBhUEpFLV96a0tqLWQ5VlhsRG9paWtIU0VEci1HUEI3RlE3V2VsQVhWaUg0VVFlMDB0d0pZaHRVSnVTcFROYTlFV3NVMmh4STIyN253b19aODVtUXRhR0NXb05xUHZKNTIyaW95OXppWEgtbjdJdTJJZkFHN0U0Iiwib2F1dGhfdG9rZW4iOiJFQUFCd3pMaXhuallCQUFGVTNMMkQ4NEZINmdIaEo1SzJaQzU4NXJDcFRWbE9ReEZaQVhBZXh4WkJCM1pDOHd3Q3pYbjNOaVU5MjJ6MFpCaElJczVxTXBQWEhQSUlhT3JkVkFVcWpTZ3hjTXVhNEJubldTakFxbEI1d3p2MkRiWVhPbE9MZnJJdzJGYXBaQlpDSTNrOHR1R01ZcFpCYUF4Q0RjeDBPSGNKeEpZYlRoUERVVGhXMmZCR0JlTFIyUEhyd2xjWkQiLCJhbGdvcml0aG0iOiJITUFDLVNIQTI1NiIsImlzc3VlZF9hdCI6MTYzMjUwNTk3MH0; csrftoken=ijEIXIQCqTMiLDSW2lRkme7EXzIPiEMb; sessionid=45678201261%3A34tYxuhSZcl46N%3A7; rur="EAG\05445678201261\0541664042024:01f76c1b522b53c74249b179092c73b73f478ee4c5cc9e0008d568f6baf3f2cdaa3e3267"'
get_id_url = f"https://www.instagram.com/{userName}/?__a=1"
header = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'cache-control': 'max-age=0',
    'cookie': cookie,
    'sec-ch-ua': '"Google Chrome";v="93", " Not;A Brand";v="99", "Chromium";v="93"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36'
}
res = requests.get(get_id_url, headers=header).content
res = json.loads(res)

edge_followed_by = res['graphql']['user']['edge_followed_by']['count']
user_id = res['graphql']['user']['id']

# 반복문 몇번?
requests_num = round(edge_followed_by / 12)
print(f'{requests_num}번 네트워크 요청이 필요합니다.')


header={
    'accept': '*/*',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'cookie': cookie,
    'origin': 'https://www.instagram.com',
    'referer': 'https://www.instagram.com/',
    'sec-ch-ua': '"Google Chrome";v="93", " Not;A Brand";v="99", "Chromium";v="93"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36',
    'x-asbd-id': '198387',
    'x-ig-app-id': '936619743392459',
    'x-ig-www-claim': 'hmac.AR1T67faovwWzH9rMZsBBfQsWnf5f8HERJSraHl-Pefgj6Eo',
}
follower_list = []
for idx in range(requests_num):
    max_id = idx * 12
    if(max_id==0):
        url = f"https://i.instagram.com/api/v1/friendships/{user_id}/followers/?count=12&search_surface=follow_list_page"
    else:
        url = f"https://i.instagram.com/api/v1/friendships/{user_id}/followers/?count=12&max_id={max_id}&search_surface=follow_list_page"
    res = requests.get(url, headers=header).content
    res = json.loads(res)
    for i in range(len(res['users'])):
        follower_username = res['users'][i]['username']
        follower_list.append(follower_username)

# follower list => .csv 파일로 저장
follower_df = pd.DataFrame(follower_list)
follower_df.to_csv(f'{userName}.csv', index=False, header=False)






