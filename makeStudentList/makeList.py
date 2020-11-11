import pandas as pd
from pptx.util import Pt
from pptx import Presentation


def makePPT(download_path, pic_image_resized_path, cha_name):
    print("교육생 정보가 담긴 데이터를 엑셀에서 복사한 후 엔터키를 눌러주세요.")
    print("조, 출력순서, 성명, 휴대폰, 나이, 대학명, 학부전공, 졸업, 거주지, 숙소 정보가 복사되어야 합니다.")
    _ = input("(복사완료 후 엔터키 입력)")
    print("교육생 명단을 PPT로 작성하는 중...")
    df = pd.read_clipboard()
    if len(df) > 0 and "성명" in df.columns:
        if len(df) <= 30:
            post_file_name = "30"
        elif len(df) <= 36:
            post_file_name = "36"
        elif len(df) <= 42:
            post_file_name = "42"
        prs = Presentation("./ppt_master/master_"+post_file_name+".pptx")
        slide = prs.slides[0]
        for student in df.iloc:
            # print(student["성명"]+" "+str(student["조"])+"-"+str(student["출력순서"]))
            for shape in slide.shapes:
                if shape.has_text_frame:
                    this_paragraph = shape.text_frame.paragraphs[0]
                    if str(student["조"]).strip() == this_paragraph.text.strip()[2:3] and str(student["출력순서"]).strip() == this_paragraph.text.strip()[4:5]:
                        this_label = this_paragraph.text.strip()[0:2]
                        if this_label == "사진":
                            files = os.listdir("./"+pic_image_resized_path)
                            for f in files:
                                if student["휴대폰"] in f:
                                    slide.shapes.add_picture("./"+pic_image_resized_path+f, shape.left, shape.top, shape.width, shape.height)
                                    break
                            this_paragraph.text = ""

                elif shape.has_table:
                    for i in range(0, 7):
                        cell = shape.table.rows[i].cells[0]
                        this_table_paragraph = cell.text_frame.paragraphs[0]
                        this_label = this_table_paragraph.text.strip()[0:2]
                        if str(student["조"]).strip() == this_table_paragraph.text.strip()[2:3] and str(student["출력순서"]).strip() == this_table_paragraph.text.strip()[4:5]:
                            if this_label == "이름":
                                this_table_paragraph.text = student["성명"]
                                this_table_paragraph.font.bold = True
                            elif this_label == "나이":
                                this_table_paragraph.text = str(student["나이"])
                            elif this_label == "대학":
                                this_table_paragraph.text = student["대학명"]
                            elif this_label == "전공":
                                this_table_paragraph.text = student["학부전공"]
                            elif this_label == "지역":
                                this_table_paragraph.text = student["거주지"]
                            elif this_label == "전화":
                                this_table_paragraph.text = student["휴대폰"]
                            elif this_label == "숙소":
                                this_table_paragraph.text = student["숙소"]
                            this_table_paragraph.font.size = Pt(8)
                        
        prs.save(download_path+"교육생 명부_"+cha_name+".pptx")
        print("교육생 명단 작성완료")
    else:
        print("교육생 정보가 클립보드로 복사되지 않았습니다.")
        makePPT(download_path, pic_image_resized_path, cha_name)
