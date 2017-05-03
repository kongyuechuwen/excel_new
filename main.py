# -*- coding: utf-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')


import xlrd

data = xlrd.open_workbook('1.xls')
table = data.sheets()[0]  # 通过索引顺序获取
nrows = table.nrows
ncols = table.ncols

data1 = xlrd.open_workbook('name.xls')
table1 = data1.sheets()[0]  # 通过索引顺序获取
nrows1 = table1.nrows
ncols1 = table1.ncols

data2 = xlrd.open_workbook('tax.xls')
table2 = data2.sheets()[0]  # 通过索引顺序获取
nrows2 = table2.nrows
ncols2 = table2.ncols
# print table.row_values(i)


# print nrows1
# print nrows2

myList = [([''] * (nrows2+1)) for i in range(nrows1)]


for i in range(nrows1):
    myList[i][0] = table1.row_values(i)[0]
    # print myList[i][0]

# print table.row_values(1)[1]

print table.row_values(6)

for m in range(nrows):
    for n in range(nrows1):
        if table.row_values(m)[1] == myList[n][0] :
            for k in range(nrows2):
                if table.row_values(m)[2] ==  table2.row_values(k)[0]:
                    # print type(myList[n][k + 1])
                    # if str(myList[n][k + 1]) == None:
                    #     myList[n][k + 1] = table.row_values(m)[3]
                    # else:
                    if myList[n][k + 1] == '':
                        myList[n][k + 1] = 0
                    myList[n][k + 1] += table.row_values(m)[3]
                    # print myList[n][k + 1]
                    # print "===" + myList[n][k + 1]
                    # a = float(myList[n][k + 1])
                    # a += table.row_values(m)[3]
                    # myList[n][k + 1] = a
                    # float(myList[n][k + 1]) += table.row_values(m)[3]
                    # print myList[n][k + 1]

# print myList[2]

# for item in myList:
#     print item

with open("Output.txt", "wb") as text_file:
    for item in myList:
        for li in item:
            text_file.write(str(li).encode('gbk')+",")
            # print li,str(li)
        text_file.write('\n')

    	# print item
        # text_file.write(str(item[0])+
        #                 ','+str(item[1])+
        #                 ',' + str(item[2]) +
        #                 ',' + str(item[3]) +
        #                 ',' + str(item[4]) +
        #                 ',' + str(item[5]) +
        #                 ',' + str(item[6]) +
        #                 ',' + str(item[7]) +
        #                 ',' + str(item[8]) +
        #                 ',' + str(item[9] )+
        #                 ',' + str(item[10]) +
        #                 ',' + str(item[11]) +
        #                 ',' + str(item[12]) +
        #                 '\n')



# # for i in range(nrows2):
#     print i
#     temp_list=[]
#     temp_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
#     result[i]=temp_list
#     temp_list =[]
# print result[1]

#
# print table1.row_values(1)
# table1.row_values(1)[1] = "156"
#
# print table1.row_values(1)
# k=1
# for i in range(nrows1):
#     for j in range(nrows2):
#         for k in range(nrows):
#             if table1.row_values(i)[0] == table.row_values(k)[1] and \
#                             table2.row_values(j)[0] == table.row_values(k)[1]:
#
#
#
#                 temp_list = [table1.row_values(i)[0], ]
#                 k = k+1



