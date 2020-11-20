from openpyxl import load_workbook
from konlpy.tag import Kkma
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

#PART1 : 데이터 로드
wb=load_workbook("sample_data.xlsx", data_only=True)
ws=wb["VOC결과"]
cells=ws["C2:C415"]
data=""
for row in cells:
	for cell in row:
		if cell.value!=None:
			data=data+" "+str(cell.value)

kkma=Kkma()

print("단어 형태소 분석 중... 기다려 주세요")
dataPos=kkma.pos(data)
dataArr=[]
for wordPos in dataPos:
	word=wordPos[0]
	pos=wordPos[1]
	if pos=="NNG" : #명사만 필터링함. 동사도 포함하려면 or pos=="VA" 추가할 것
		dataArr.append(word)

print("단어별 발생빈도 계산중...")
counter=Counter(dataArr).most_common()
keywords={}
resultsFile=open("results.txt", "w", encoding="utf-8")

percent = 0
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