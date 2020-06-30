# -*- coding: utf-8 -*-
"""
Created on Fri Jan  3 10:02:48 2020

@author: Long Bui
email longbui189@gmail.com
"""

import pandas as pd
#reading file
file = 'd:/don cole/python/hehe/benhsoi.xls'
xls = pd.ExcelFile(file)
xls_sheets=[]
for sheet in xls.sheet_names:
    xls_sheets.append(xls.parse(sheet))
    measle = pd.concat(xls_sheets)
#sort by province
measle_sort = measle.sort_values(['prov'], ascending = [True])

#doc file ma tinh, huyen, xa & gan cho index
file = 'd:/don cole/python/hehe/index.xls'
xls = pd.ExcelFile(file)
xls_sheets=[]
for sheet in xls.sheet_names:
    xls_sheets.append(xls.parse(sheet))
    index = pd.concat(xls_sheets)
#sap xep theo Tinh
index_province = index.sort_values(['prov'], ascending = [True])

#-------------------------------------
#Function campare name1 & name2;index = 0: normal comparision; 1: province; 2: district; 3: commune
def compare (name1="", name2="",index=0):  
    
    #uppercase
    name1 = name1.upper()
    name2 = name2.upper()
    #remove space
    name1=name1.replace(' ','',len(name1))
    name2=name2.replace(' ','',len(name2))
    #remove accent
    char_mark     ="Ả,Á,À,Ạ,Ã,Ă,Ẳ,Ắ,Ằ,Ặ,Ẵ,Â,Ẩ,Ấ,Ầ,Ậ,Ẫ,Ẻ,É,È,Ẹ,Ẽ,Ê,Ể,Ế,Ề,Ệ,Ễ,Ủ,Ú,Ù,Ụ,Ũ,Ư,Ử,Ứ,Ừ,Ự,Ữ,Ỉ,Í,Ì,Ị,Ĩ,Ỏ,Ó,Ò,Ọ,Õ,Ơ,Ở,Ớ,Ờ,Ợ,Ỡ,Ô,Ổ,Ố,Ồ,Ộ,Ỗ,Ỷ,Ý,Ỳ,Ỵ,Ỹ"
    char_no_mark  ="A,A,A,A,A,A,A,A,A,A,A,A,A,A,A,A,A,E,E,E,E,E,E,E,E,E,E,E,U,U,U,U,U,U,U,U,U,U,U,I,I,I,I,I,O,O,O,O,O,O,O,O,O,O,O,O,O,O,O,O,O,Y,Y,Y,Y,Y"
    trans=name1.maketrans(char_mark,char_no_mark)
    name1 = name1.translate(trans)
    trans=name2.maketrans(char_mark,char_no_mark)
    name2 = name2.translate(trans)
    # remove prefix 
    if index == 1: #province
        name1 = name1.replace('TINH','')
        name1 = name1.replace('THANHPHO','')
        name2 = name2.replace('TINH','')
        name2 = name2.replace('THANHPHO','')
    elif index == 2: #district
        name1 = name1.replace('HUYEN','')
        name1 = name1.replace('QUAN','')
        name1 = name1.replace('THIXA','')
        name1 = name1.replace('THANHPHO','')
        name2 = name2.replace('HUYEN','')
        name2 = name2.replace('QUAN','')
        name2 = name2.replace('THIXA','')
        name2 = name2.replace('THANHPHO','')
    elif index == 3: #commune
        name1 = name1.replace('XA','')
        name1 = name1.replace('PHUONG','')
        name1 = name1.replace('THITRAN','')
        name2 = name2.replace('XA','')
        name2 = name2.replace('PHUONG','')
        name2 = name2.replace('THITRAN','')
    return (name1 == name2)
         
#find province id, if not, return -1
def find_province(prov=""):
    result = []
    total_rows=index_province.shape[0]
    #print(total_rows)
    for i in range(total_rows):
        if compare(index_province.iat[i,5],prov,1):
            result.append(index_province.iat[i,4])
            result.append(i)
            break
    return result
#-------------------------------------
#-------------------------------------
#find district, if not, return -1
def find_district(dist="",row =0):
    result =[]
    total_rows = index_province.shape[0]
    for i in range(row,total_rows):
        if compare(index_province.iat[i,3],dist,2):
            result.append(index_province.iat[i,2])
            result.append(i)
            break
    return result
#-------------------------------------
#-------------------------------------
#find commune, if not, return -1
def find_commune(comm="",row =0):
    result =[]
    total_rows = index_province.shape[0]
    for i in range(row,total_rows):
        if compare(index_province.iat[i,1],comm,3):
            result.append(index_province.iat[i,0])
            result.append(i)
            break
    return result
#-------------------------------------
# Repair 5 big cities
def standard_city(city=""):
    HCM = ["TP. HỒ CHÍ MINH", "TP HỒ CHÍ MINH",'TP. HCM', 'TP HCM', "TP.H.C.M","TP. H.C.M",'TP H.C.M','THÀNH PHỐ H.C.M', 'THÀNH PHỐ HCM', 'THÀNH PHỐ HỒ CHÍ MINH', 'HỒ CHÍ MINH', 'H.C.M', 'HCM', 'H C M']
    HN  = ['TP. HÀ NỘI', 'TP HÀ NỘI','TP. HN', 'TP HN', 'TP. H.N', 'TP. H.N','THÀNH PHỐ H.N','THÀNH PHỐ HN', 'THÀNH PHỐ HÀ NỘI', 'HÀ NỘI', 'H.N','H. N', 'HN', 'H N']
    DN  = ['TP. ĐÀ NẴNG', 'TP ĐÀ NẴNG','TP. ĐN','TP. DN', 'TP ĐN', 'TP DN', 'TP.Đ.N','TP. Đ.N','TP. D.N', 'TP. Đ.N','TP. D.N','THÀNH PHỐ ĐN', 'THÀNH PHỐ ĐÀ NẴNG', 'ĐÀ NẴNG', 'Đ.N', 'D.N','ĐN', 'DN']
    HP  = ['TP. HẢI PHÒNG', 'TP HẢI PHÒNG','TP. HP', 'TP HP', 'TP. H.P', 'TP.H.P','THÀNH PHỐ H.P', 'THÀNH PHỐ HP', 'THÀNH PHỐ HẢI PHÒNG', 'HẢI PHÒNG', 'H.P', 'HP']
    CT  = ['TP. CẦN THƠ', 'TP CẦN THƠ','TP. CT', 'TP CT', 'TP. C.T', 'TP. C.T','THÀNH PHỐ C.T','THÀNH PHỐ CT', 'THÀNH PHỐ CẦN THƠ', 'CẦN THƠ', 'C.T', 'CT']
    if (city.upper() in HCM):
        return "Thành phố Hồ Chí Minh"
        
    elif (city.upper() in HN):
        return "Thành phố Hà Nội"
       
    elif (city.upper() in DN):
        return "Thành phố Đà Nẵng"
        
    elif (city.upper() in HP):
        return "Thành phố Hải Phòng"
         
    elif (city.upper() in CT):
        return "Thành phố Cần Thơ"
    
    else:
        return ""
#---------------------------------------
#repair province
def standard_province (prov=""):
   TTH = ["THỪA THIÊN HUẾ", "TT-HUẾ","TT- HUẾ","TT - HUẾ"]
   BRVT = ["BÀ RỊA VŨNG TÀU", "B.RỊA - V.TÀU","BÀ RỊA- V.TÀU","BÀ RỊA-V.TÀU"]
   if (prov.upper() in TTH):
       return "Tỉnh Thừa Thiên Huế"
   elif (prov.upper() in BRVT):
       return "Tỉnh Bà Rịa - Vũng Tàu"
   else:
       return "Tỉnh " + prov
#---------------------------------------      
# replace ID to file
def replace_id ():
    result =[]
    count_not_id=0
    total_rows= measle_sort.shape[0]
    #total_rows=100
    print("total rows: ", total_rows)
    row = row_prov=0
    while (row < total_rows):
        #find duplicate province
        province = str(measle_sort.iat[row,3])
        row_prov = row+1
        while (row_prov<total_rows) and (measle_sort.iat[row_prov-1,3] == measle_sort.iat[row_prov,3]) :
            row_prov=row_prov+1
    
        city = standard_city(province)
        if city=="":
            province = standard_province(province)
        else:
            province = city            
        #find province ID
        id_tinh = find_province(province)
        if (id_tinh  == []):
            #print("--- ", province, " incorrect ID \n")
            count_not_id = count_not_id+1
            if province not in result:
                result.append(province)
        else:
            for i in range (row,row_prov):
                #thay the tinh/tp bang id
                measle_sort.iat[i,3]=id_tinh[0]
                #find and replace district id
                id_huyen = find_district(measle_sort.iat[i,4],id_tinh[1]) 
                if (id_huyen  == []):
                   # print("--- ", measle_sort.iat[i,4] , "  thuộc ", id_tinh[0], " không tìm đúng ID\n")
                   count_not_id = count_not_id+1
                   if measle_sort.iat[i,4] not in result:
                       result.append(measle_sort.iat[i,4])
                else: 
                    measle_sort.iat[i,4]=id_huyen[0]
                    #find and replace commune id
                    id_xa = find_commune(measle_sort.iat[i,5],id_huyen[1])
                    if (id_xa  == []):
                        #print("--- ", measle_sort.iat[i,5], " thuộc: ", id_huyen[0], " và ", id_tinh[0], " không tìm đúng ID\n")
                        count_not_id = count_not_id+1
                        if measle_sort.iat[i,5] not in result:
                            result.append(measle_sort.iat[i,5])
                    else:
                        measle_sort.iat[i,5]=id_xa[0]
         #next row   
        row = row_prov            
    result.append(count_not_id)
    return result
#--------------------------------------
#--------------------------------------

result=[]
result = replace_id()
print("Finished replacing \n")
print("Place where can not find ID\n", result)
#save to kq.xls
file = ''
with pd.ExcelWriter(file) as writer:
    measle_sort.to_excel(writer,sheet_name='ID')



