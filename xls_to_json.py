import json
import sys

#import xlrd

import xlrd

workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_name(sys.argv[2])
fileName = sys.argv[1].split(".")
vartype = fileName[0]
counter =0;
data = []
keys = [v.value for v in worksheet.row(0)]
for row_number in range(worksheet.nrows):
    if row_number == 0:
        continue
    row_data = {}
    for col_number, cell in enumerate(worksheet.row(row_number)):
        row_data[keys[col_number]] = cell.value
        print("row_data: ",row_data);
    counter=counter+ 1
    create ={ "create" :    { "_index" : fileName[0], "_type" :vartype[:-1] , "_id" : counter } }
    
    print("****************************");
    print(data.append(create))
    data.append(row_data)
    print("after Apped", data)
    with open(sys.argv[3], 'w') as json_file:
        print("\n ******************split*****************")
        for i in data:
            print(i)
            json_file.write(json.dumps(i))
            json_file.write("\n")
    




