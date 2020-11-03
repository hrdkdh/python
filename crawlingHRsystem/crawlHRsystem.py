
import urllib
import requests
from bs4 import BeautifulSoup as bs
from datetime import datetime

first_url = "http://swpsso.posco.net/idms/U61/jsp/login/login.jsp"
login_url = "http://swpsso.posco.net/idms/U61/jsp/login/loginProc.jsp"
login_data = {
    "username": "hrdkdh",
    "password": "echoes78(",
    "login-form-type": "pwd"
}
# now_timestamp = str(int(datetime.now().timestamp()))+"765"
now_timestamp = "1604304759945"
cookie_obj1 = requests.cookies.create_cookie(domain=".posco.net", name="swpsso_logintime", value=now_timestamp, path="/")
cookie_obj2 = requests.cookies.create_cookie(domain=".posco.net", name="loginproctimeflag", value="", path="/")

with requests.Session() as s:
    s.cookies.set_cookie(cookie_obj1)
    s.cookies.set_cookie(cookie_obj2)
    s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; TCO_20201030171244; rv:11.0) like Gecko"
    s.headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
    s.headers["Accept-Encoding"] = "gzip, deflate, br"
    s.headers["Accept-Language"] = "ko,ko-KR;q=0.9,en;q=0.8"
    s.headers["Cache-Control"] = "max-age=0"
    s.headers["Content-Length"] = "56"
    s.headers["Content-Type"] = "application/x-www-form-urlencoded"
    s.headers["Host"]: "swpsso.posco.net"
    s.headers["Origin"]: "http://swpsso.posco.net"
    s.headers["Referer"]: "http://swpsso.posco.net/"
    s.headers["Sec-Fetch-Dest"]: "document"
    s.headers["Sec-Fetch-Mode"]: "navigate"
    s.headers["Sec-Fetch-Site"]: "cross-site"
    s.headers["Sec-Fetch-User"]: "?1"
    s.headers["Upgrade-Insecure-Requests"]: "1"
    
    first_page_req = s.get(first_url)
    login_req = s.post(login_url, data=login_data)
    print(login_req.text)
    print(s.headers)
    print(s.cookies)