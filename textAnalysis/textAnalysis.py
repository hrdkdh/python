from openpyxl import load_workbook
from konlpy.tag import Kkma
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

#PART1 : 데이터 로드
wb=load_workbook("data.xlsx", data_only=True)
ws=wb["설문 응답결과"]
cells=ws["F2:I73"]
data=""
for row in cells:
	for cell in row:
		if cell.value!=None:
			data=data+" "+str(cell.value)

kkma=Kkma()

print("단어 형태소 필터링...")
dataPos=kkma.pos(data)
dataArr=[]
for wordPos in dataPos:
	word=wordPos[0]
	pos=wordPos[1]
	if pos=="NNG" and word!="토":
		if word=="선배":
			dataArr.append("멘토/선배")
		else:
			dataArr.append(word)
	elif pos=="NNG" and word=="토":
		dataArr.append("멘토/선배")
	elif pos=="VA" and word=="좋":
		dataArr.append("좋았다")

print("빈도 계산중...")
counter=Counter(dataArr).most_common()
keywords={}
resultsFile=open("results.txt", "w", encoding="utf-8")

for keyword in counter:
	word=keyword[0]
	freq=keyword[1]
	if len(word)>1 and freq>2: #한 글자 이상 단어 + 빈도수가 2 이상인 것만 추출
		keywords[word]=freq
		thisText=word+" : "+str(freq)+"건\n"
		resultsFile.write(thisText)

resultsFile.close()

print("워드클라우드 생성중...")
font_path="NanumBarunGothicBold.ttf"
wordcloud=WordCloud(
	font_path=font_path,
	width=800,
	height=800,
	background_color="white"
)
wordcloud=wordcloud.generate_from_frequencies(keywords)
array=wordcloud.to_array()

fig=plt.figure(figsize=(10, 10))
plt.axis("off")
plt.imshow(array, interpolation="bilinear")

fig.savefig("wordcloud.png")
plt.show()

print("워드클라우드 생성완료")
exit(1)