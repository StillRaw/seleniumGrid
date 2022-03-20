import pandas as pd
import glob
import os


def excelMerge():
    # specifying the path to csv files
    path = "C:\\Users\\Onur\\Desktop\\iskur\\test"
    
    # csv files in the path
    global excl_file
    #excl_file = glob.glob('/*.xlsx')
    excl_file = glob.glob(path + "\\*.xlsx")
    #print(excl_list)
    
    # list of excel files we want to merge.
    # pd.read_excel(file_path) reads the excel
    # data into pandas dataframe.
    global excl_list
    
    excl_list = []
    
    for file in excl_file:
        excl_list.append(pd.read_excel(file))

    # create a new dataframe to store the 
    # merged excel file.
    global excl_merged
    excl_merged = pd.DataFrame()
    
    for excl_file in excl_list:
        
        # appends the data into the excl_merged 
        # dataframe.
        excl_merged = pd.concat([excl_merged, excl_file])
    
    #Adjusting
    excl_merged.sort_values(by=['Rating'], inplace=True, ascending=False)#Sortin values acoording to Rating column.
    

def miniAnalysis():#MINI ANALYSIS
    #Year Analysis
    global count1,count2,count3,US
    year = excl_merged["Year"]
    count1 = 0
    count2 = 0
    count3 = 0
    
    for i in year:
        i=int(i)

        if i < 1990:
            count1 += 1
        elif i >= 1990 and i < 2000:
            count2 += 1
        else:
            count3 +=1
    
    #Origin Analysis
    origin=excl_merged["Origin"]
    US=0
    for i in origin:
        if "United States" in i:
            US += 1


def createExcel():# exports the dataframe into excel file with specified name.
    file_name='IMDB250.xlsx'
    sheet_name = 'Final'
    writer = pd.ExcelWriter("C:\\Users\\Onur\\Desktop\\iskur\\test\\"+file_name, engine='xlsxwriter')
    excl_merged.to_excel(writer,sheet_name=sheet_name, index=False)

    workbook  = writer.book
    worksheet = writer.sheets[sheet_name] 
    worksheet.write(len(excl_merged)+4, 0, 'Mini Analysis', workbook.add_format({'bold': True}))
    worksheet.write(len(excl_merged)+5, 0, 'There are ' + str(count1) + ' movies before 1990.', workbook.add_format({'bold': False}))
    worksheet.write(len(excl_merged)+6, 0, 'There are ' + str(count2) + ' movies in 1990s.', workbook.add_format({'bold': False}))
    worksheet.write(len(excl_merged)+7, 0, 'There are ' + str(count3) + ' movies after 2000.', workbook.add_format({'bold': False}))
    worksheet.write(len(excl_merged)+8, 0, 'There are ' + str(US) + ' movies that US coorparete or product itself.', workbook.add_format({'bold': False}))
    
    writer.save()

def deleteParts():
    file = glob.glob('C:\\Users\\Onur\\Desktop\\iskur\\test\\*.xlsx')
    print(file)
    for f in file:
        if f=="C:\\Users\\Onur\\Desktop\\iskur\\test\\IMDB250.xlsx":
            continue
        try:
            os.unlink(f)
        except OSError as e:
            print("Error: %s : %s" % (f, e.strerror))

excelMerge()
miniAnalysis()
createExcel()
#deleteParts()