import pandas as pd
import numpy as np
import os

class Get_click_persons(object):

    def __init__(self):
        self.books_name = []
        self.query_date = []

    def enter_user(self):
        if len(self.books_name) > 0 and len(self.query_date) > 0:
            return 
        # books_total = input('你要查询几本书：（请输入整数）\n')
        self.books_name.append('日期')
        
        # for i in range(int(books_total)):
        #     name = input('book name:\n')
        #     self.books_name.append(name)
            
        # for i in range(2):
        #     value = input('你要查询的日期:\n')
        #     self.query_date.append(value)

        self.query_date.append('4-26')
        self.query_date.append('4-25')
        self.books_name.append('偶像归来:狼来了')
        self.books_name.append('四海第一娇')
        self.books_name.append('青你2.导师太会了')


    def list_files(self):
        read_path = []
        all_files = os.listdir()
        for item in all_files:
            if 'csv' in item:
                read_path.append(item)
        return read_path

    def get_results(self,filepath):
        self.enter_user()
        book_name2 = self.books_name
        query_date2 = self.query_date

        data = pd.read_csv(filepath)
        # print(data)
        dat = data.loc[0:,book_name2]
        # print(dat)
        handle_data = {}
        book_number = []
        
        for q_date in query_date2:
            book_number = []
            for i in range(0,len(dat)):
                item = dat[i:i+1]
                # print(item)
                date_details = item.loc[[i],book_name2].values[0]
                # print(date_details)
                date = date_details[0]
                if q_date in date:
                    for k in range(1,len(book_name2)):
                        book_number.append({book_name2[k]:date_details[k]})
                        handle_data[date] = book_number
                        # print(q_date)
                        # print(books_name[k])
                        # print(date_details[k])
                        
                    # print(date_details)
        return handle_data

    def run(self):
        files = self.list_files()
        all_results = {}
        for path in files:
            filename = path.split('_')[0]
            # print(filename)
            # if filename == '作品收藏用户数':
                # print('note')
            # print(path)
            result = self.get_results(path)
            # print(files)
            # print(result)
            all_results[filename] = result
        all_results['date'] = self.query_date
        # print(all_results)
        return all_results

if __name__ == "__main__":
    test = Get_click_persons()
    test.run()

