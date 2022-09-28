import wordcloud
import xlrd
import imageio
rb = xlrd.open_workbook(r'./RCR.xls')
sheet1 = rb.sheet_by_name('sheet1')
list1 = sheet1.col(5)[1:]
list2 = []
for i in list1:
    if i.value == '':
        pass
    else: 
        str1 = i.value.replace(' ','_')
        list2.append(str1)

string = ' '.join(list2)
mk = imageio.imread('./beijin1.jpg')
x = wordcloud.WordCloud(background_color="white",mask=mk)
x.generate(string)
x.to_file('./pywordcloud1.png')