# %%
import pandas as pd
import numpy as np
from os import listdir
bigDF = pd.DataFrame()
inputFolder = "D:/Covid_data_New/Yr2023/Feb/Feb27/Excels/"
listName = 0
filesInDir = listdir(inputFolder)
listExcel = []
for file in filesInDir:
    if "xlsx" in file:
        listExcel.append(file)
#print(filesInDir)
for excelFile in listExcel:
    listName+=1
    annotationFile = pd.read_excel(inputFolder+"/"+excelFile,header=None,sheet_name="Results")
    metaData = annotationFile[:20]

    dataForInfo = []
    j=0
    for i in metaData[0]:
        if i == "File Name":
            dataForInfo.append(metaData[1][j])
        if i == "Instrument Type":
            dataForInfo.append(metaData[1][j])
        if i == "Block Type":
            dataForInfo.append(metaData[1][j])
        if i == "Passive Reference":
            dataForInfo.append(metaData[1][j])      
        j+=1

    

    completeData = annotationFile[20:]
    #print(completeData)
    completeData1 = pd.DataFrame(completeData.values[1:],columns= completeData.iloc[0])

    sampleList = []
    targetDayes = []
    for i in completeData1.index:
        samples = completeData1.loc[i,"Sample"]
        targetDayes.append(completeData1.loc[i,"Reporter"])
        sampleList.append(samples)

    dyes = np.unique(targetDayes)
    totalSamples = len(sampleList)/len(dyes)
    dataForInfo.append(totalSamples)

    
    dataForInfo.append(dyes)
    dataFrame = pd.DataFrame([dataForInfo],[listName],["File Name","Instrument Type","Block Type","Passive Reference","Total Samples","Reporter"])
    bigDF = pd.concat([dataFrame,bigDF])
    

bigDF = bigDF[::-1]
outFolder = "D:/Covid_data_New/Yr2023/Feb/Feb27/"
excelWriter = pd.ExcelWriter(outFolder+"/aa.xlsx")
bigDF.to_excel(excel_writer=excelWriter,sheet_name= "Data",index=False)
excelWriter.close()


