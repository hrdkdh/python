import ssl
import pandas as pd
from urllib.request import urlopen
from bs4 import BeautifulSoup

stock_code = "005490"
start_ymd = 20201015
end_ymd = 20201022
host = "https://finance.naver.com"
df = pd.DataFrame(columns=["ymd", "time", "price"])

for ymd in range(start_ymd, end_ymd): #end_ymd 넣을 것
    print(str(ymd) + " : 수집중...")
    url = host+"/item/sise_time.nhn?code="+stock_code+"&thistime="+str(ymd)+"161036&page=1"

    context = ssl._create_unverified_context()
    soup = BeautifulSoup(urlopen(url, context=context).read(), "html.parser")

    page_url = host+"/item/sise_time.nhn?code=005490&thistime="+str(ymd)+"161036&page="
    last_page_no = int(soup.findAll("a")[-1].attrs["href"].split("=")[-1])

    for page_no in range(1, last_page_no): #2대신 last_page_no 넣을 것
        # print(page_url+str(page_no))
        this_soup = BeautifulSoup(urlopen(page_url+str(page_no), context=context).read(), "html.parser")
        this_tr = this_soup.findAll("table")[-2].findAll("tr")
        for i, tr in enumerate(this_tr):
            if i>1 and len(tr.findAll("td"))>2 and tr.findAll("td")[0] is not None:
                df = df.append({"ymd" : ymd, "time" : tr.findAll("td")[0].findAll("span")[0].text.replace(":", ""), "price" : tr.findAll("td")[1].findAll("span")[0].text.replace(",","")}, ignore_index=True)

df.to_html("result.html")
print(df)