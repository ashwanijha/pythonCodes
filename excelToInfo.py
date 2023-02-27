# %%
import numpy as np
import pandas as pd

def getInformation(metaData,completeData,listName):
    dataForInfo = []
    dataFrame = pd.DataFrame()
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
    return dataFrame    


# %%
from os import listdir
outFolder = "D:/Covid_data_New/Yr2023/Feb/Feb27/"
bigDF = pd.DataFrame()
inputFolder = "D:/Covid_data_New/Yr2023/Feb/Feb27/Excels/"
listName=0
filesInDir = []
with open(outFolder + "/data.txt") as inFile:
    for lines in inFile:
         filesInDir.append(lines.strip())


listExcel = []
for file in filesInDir:
    if "xlsx" in file:
            listName+=1
            annotationFile = pd.read_excel(inputFolder+"/"+file,sheet_name="Results",header=None)
            #print(file+"\t"+s)
            #outFile.write(file+"\t"+str(annotationFile.loc[1,"Sample"])+"\n")
            metaData = annotationFile[:20]
            completeData = annotationFile[20:]
            bigDF = pd.concat([getInformation(metaData,completeData,listName),bigDF])
bigDF = bigDF[::-1]
excelWriter = pd.ExcelWriter(outFolder+"/matrix_data.xlsx")
bigDF.to_excel(excel_writer=excelWriter,sheet_name= "Data",index=False)
excelWriter.close()


