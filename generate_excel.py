import xlwt
import os
from handle_data import Get_click_persons

class GenerateXLSX(object):

    def __init__(self):
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet('mysheet')
        self.worksheet.col(0).width = 4444

    def backcloud(self,i):
        style = xlwt.XFStyle() # Create the Pattern
        pattern = xlwt.Pattern() # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = i # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        style.pattern = pattern # Add Pattern to Style\
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        style.alignment = alignment
        font = xlwt.Font()
        font.bold = True
        font.name = '华文楷体'
        font.height = 14 * 20
        style.font = font
        return style


    def write_three(self,q_date,book_key,row_index,col_index):
        # 作品曝光量
        result = self.get_results()
        book_result = result[book_key]
        print('*'*80)
        print(book_result)

        numbers = []
        numbers2 = []
        one_date = book_result[q_date[0]]
        second_date = book_result[q_date[1]]
        for item in one_date:
            for itm in item.values():
                numbers.append(itm)

        for item in second_date:
            for itm in item.values():
                numbers2.append(itm)
        
        for i in range(0,len(numbers)):
            try:

                result = (float(str(numbers[i]).replace(',','')) - float(str(numbers2[i]).replace(',',''))) / float(str(numbers2[i]).replace(',',''))
            except Exception as e:
                result = 0
                print('error message is :' + str(e))
            
            if result == 0:
                result = 0
            else:
                result = result * 100
                result = float('%.2f'%result)
                result = str(result) + '%'
            print(result)
            self.worksheet.write(row_index + i,col_index ,numbers[i],self.custom_font())
            self.worksheet.write(row_index + i,col_index + 1,numbers2[i],self.custom_font())
            self.worksheet.write(row_index + i,col_index + 2,result,self.custom_font())
            if book_key == '作品曝光量':
                rate = i + 3
                rate1 = 'E' + str(rate) + '/' + 'B' + str(rate)
                rate2 = 'F' + str(rate) + '/' + 'C' + str(rate)
                rate3 = '(Q' + str(rate) + '-' + 'R' + str(rate) +')/R' +str(rate)
                self.worksheet.write(row_index + i,16,xlwt.Formula(rate1),self.custom_font())
                self.worksheet.write(row_index + i,17,xlwt.Formula(rate2),self.custom_font())
                self.worksheet.write(row_index + i,18,xlwt.Formula(rate3),self.custom_font())
            if book_key == '章节阅读用户数':
                rate = i + 3
                rate1 = 'K' + str(rate) + '/' + 'R' + str(rate)
                rate2 = 'L' + str(rate) + '/' + 'I' + str(rate)
                rate3 = '(T' + str(rate) + '-' + 'U' + str(rate) +')/U' +str(rate)
                self.worksheet.write(row_index + i,19,xlwt.Formula(rate1),self.custom_font())
                self.worksheet.write(row_index + i,20,xlwt.Formula(rate2),self.custom_font())
                self.worksheet.write(row_index + i,21,xlwt.Formula(rate3),self.custom_font())
            if book_key == '作品收藏用户数':
                rate = i + 3
                rate1 = 'N' + str(rate) + '/' + 'K' + str(rate)
                rate2 = 'O' + str(rate) + '/' + 'L' + str(rate)
                rate3 = '(W' + str(rate) + '-' + 'X' + str(rate) +')/X' +str(rate)
                self.worksheet.write(row_index + i,22,xlwt.Formula(rate1),self.custom_font())
                self.worksheet.write(row_index + i,23,xlwt.Formula(rate2),self.custom_font())
                self.worksheet.write(row_index + i,24,xlwt.Formula(rate3),self.custom_font())

    def custom_font(self):
        font = xlwt.Font()
        font.height= 12*20
        font.name = '等线'
        style = xlwt.XFStyle()
        style.font = font
        return style

    def write_xlsx(self):
        self.worksheet.write_merge(0,0,1,3, '作品曝光量', self.backcloud(5))
        self.worksheet.write_merge(0,0,4,6, '点击用户数', self.backcloud(2))
        self.worksheet.write_merge(0,0,7,9,'书首页用户数',self.backcloud(14))
        self.worksheet.write_merge(0,0,10,12,'章节阅读用户数',self.backcloud(13))
        self.worksheet.write_merge(0,0,13,15,'作品收藏用户数',self.backcloud(14))
        self.worksheet.write_merge(0,0,16,18,'点击转换率',self.backcloud(5))
        self.worksheet.write_merge(0,0,19,21,'阅读转换率',self.backcloud(3))
        self.worksheet.write_merge(0,0,22,24,'收藏转换率',self.backcloud(5))

        # self.worksheet.write(11,0,'1. 曝光和点击转化低：上推后最高转化率4.68%，最低转化率仅1.53%.曝光量并没有转化为点击量,反应首页推荐模块的作品并不是用户感兴趣的内容；')
        self.worksheet.write(1,0,'作品名',self.custom_font())

        # date = ['4-26(日)','4-25(六)']
        date = Get_click_persons().run()['章节阅读用户数']
        q_date = []
        for item in date:
            q_date.append(item)
        
        for i in range(1,25,3):
            self.worksheet.write(1,i,q_date[0],self.custom_font())
            self.worksheet.write(1,i+1,q_date[1],self.custom_font())
            self.worksheet.write(1,i+2,'增长率',self.custom_font())

        result = self.get_results()
        print(result)

        zpbgl = result['作品曝光量']
        book_names = []
        one_date = zpbgl[q_date[0]]
        for item in one_date:
            for itm in item.keys():
                book_names.append(itm)
        for i in range(0,len(book_names)):
            self.worksheet.write(2+i,0,book_names[i])

        
        self.write_three(q_date,'作品曝光量',2,1)
        self.write_three(q_date,'作品点击用户数',2,4)
        self.write_three(q_date,'书首页用户数',2,7)
        self.write_three(q_date,'章节阅读用户数',2,10)
        self.write_three(q_date,'作品收藏用户数',2,13)
        self.workbook.save('dddddddddddddddddddddddddddd.xls')

    def get_results(self):
        results = Get_click_persons()
        click_persons_for_bookexposure = results.run()
        return click_persons_for_bookexposure

    def run(self):
        self.write_xlsx()

if __name__ == "__main__":
    main = GenerateXLSX()
    main.run()

# workbook.save('Excel_Workbook.xls')


    