# dataAnalysis

### 소니 크롤링 코드


import urllib.request as req
from bs4 import BeautifulSoup
import os
import openpyxl
import datetime
from openpyxl.drawing.image import Image

# 기존 엑셀파일 삭제
if os.path.exists("./소니세일_크롤링.xlsx"):
 os.remove("./소니세일_크롤링.xlsx")

# 이미지 저장할 폴더 생성
if not os.path.exists("./소니세일이미지"):
 os.mkdir("./소니세일이미지")


# 엑셀 파일 생성
if not os.path.exists("./소니세일_크롤링.xlsx"):
 openpyxl.Workbook().save("./소니세일_크롤링.xlsx")

# 엑셀 파일 불러오기
book = openpyxl.load_workbook("./소니세일_크롤링.xlsx")
# 쓸데 없는 시트 지우기
if "Sheet" in book.sheetnames:
 book.remove(book["Sheet"])
sheet = book.create_sheet()
now = datetime.datetime.now()
sheet.title = f"{now.year}년 {now.month}월 {now.day}일 {now.hour}시 {now.minute}분 {now.second}초"
# 열 너비 조절
sheet.column_dimensions["A"].width = 10
sheet.column_dimensions["B"].width = 100
sheet.column_dimensions["c"].width = 30


page_num = 1
row_num = 1
while True:
 code = req.urlopen(f"https://store.playstation.com/ko-kr/category/383e6eb8-9ec6-4a85-a466-15e60b72d3af/{page_num}")
 soup = BeautifulSoup(code, "html.parser")
 title = soup.select("span.psw-t-body.psw-c-t-1.psw-t-truncate-2.psw-m-b-2")
 price = soup.select("span.psw-m-r-3")
 img = soup.select("span.psw-media-frame.psw-fill-x.psw-image.psw-media.psw-media-interactive.psw-aspect-1-1 > img")


 for i in range(len(title)):
   img_file_name = f"./소니세일이미지/{i+1}.png"
   req.urlretrieve(img[i].attrs["src"], img_file_name)
   print(f"{title[i].text} - {price[i].text}")
   img_for_excel = Image(img_file_name)
   sheet.add_image(img_for_excel, f"A{i+1}")
   sheet.cell(row=row_num, column=2).value = title[i].text
   sheet.cell(row=row_num, column=3).value = price[i].text
   sheet.row_dimensions[i+1].height = 50
   book.save("./소니세일_크롤링.xlsx")
   row_num += 1
 page_num += 1
