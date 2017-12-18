import json
import sys
import datetime
import xlrd



workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_name(sys.argv[2])
fileName = sys.argv[1].split(".")
vartype = fileName[0]
counter =0;
data = []


a=None
keys = [v.value for v in worksheet.row(0)]
for row_number in range(worksheet.nrows):
    if row_number == 0:
        continue
    row_data = {}
    for col_number, cell in enumerate(worksheet.row(row_number)):
        a =row_data.keys()
        def check():
            for i in keys[col_number].split(" "):
                if(("Time" in i) or ("Date" in i) or  ("Duration" in i) or ("Duration (HH:MM)" in i)or ("From Time" in i) or ("End Time" in i)):
                    if((keys[col_number]=="Date") or (keys[col_number] == "Duration") or ( keys[col_number] =="Duration (HH:MM)" ) or (keys[col_number] == "From Time") or (keys[col_number] == "End Time")  ):
                        if((keys[col_number] =="Duration (HH:MM)") or (keys[col_number]=="Duration") or (keys[col_number] == "From Time") or (keys[col_number] == "End Time")): 
                            time_val = str(xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode)).split(" ")
                            row_data[keys[col_number]]=time_val[1]
                        elif(keys[col_number]=="time"):    
                            row_data[keys[col_number]] = str(xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode))
                            new_col_array={'date': 1542, 'time':0000};
                            interval = str(xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode)).split(" ")
                            new_col_array['date'] = interval[0]
                            new_col_array['time'] = interval[1]
                            count=0;
                            for i in new_col_array.keys():
                                if(i == "time"): row_data[i]=interval[1]
                                else:   row_data[i]=interval[0]                        
                                count=count + 1
                            count = 0    
                    else:       
                        row_data[keys[col_number]] = str(xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode))
                    return True
                    
        if(not(check())):
            row_data[keys[col_number]] = cell.value
    counter=counter+ 1
    create ={ "create" :    { "_index" : fileName[0], "_type" :vartype[:-1] , "_id" : counter } }
    
    
    data.append(create)
    data.append(row_data)
    print(data)
    with open(sys.argv[3], 'w') as json_file:
        print("\n ******************split*****************")
        for i in data:
            print("on json: ",i)
            print("\n")
            json_file.write(json.dumps(i))
            json_file.write("\n")
    




