# %%
import pandas as pd
from os import listdir
import openpyxl as xl

# %%

def readExcelFiles(exportedFile):
    export1 = exportedFile[26:]
    metaDf = exportedFile[:26]
    export = pd.DataFrame(export1.values[1:], columns=export1.iloc[0]) 
    return metaDf, export

# %%
def getmetaDataInfo(metaDf):
    metaData = metaDf[1]
    headMeta = metaDf[0]
    listAddHeader = []
    runName = str(metaData[0]).split("\\")
    listAddElemane=[]
    listAddElemane.append(runName[-1])
    listAddHeader.append(headMeta[0])
    #print(runName[-1])
    for i in range(1,len(metaDf[1])-1):
        listAddHeader.append(headMeta[i])
        if pd.isnull(metaData[i]):
            listAddElemane.append("")
        else:
            listAddElemane.append(metaData[i])
    return listAddElemane, listAddHeader


# %%
def implementLogic(export,listAddElemane,listAddHeader):
    targetThresholds={"VIC":38,"ABY":37,"FAM":38,"ALEXA 647":38,"JUN":33}
    targetCalls = dict()
    data = dict()
    for i in export.index:
        wellPosition = export.loc[i,"Well"]
        sampleName = export.loc[i,"Sample"]
        targets = export.loc[i,"Target"]
        cqValue = export.loc[i,"Cq"]
        reporter = export.loc[i,"Reporter"]
        cols = wellPosition+"\t"+sampleName
        cqValues = targets+"\t"+str(cqValue)
        data.setdefault(cols,[]).append(cqValues)
        if cqValue == "UNDETERMINED":
            cqValue = 43
        if targetThresholds[reporter] >= float(cqValue):
            targetCalls.setdefault(cols,[]).append("Positive")
        else:
            targetCalls.setdefault(cols,[]).append("Negative")

    df = pd.DataFrame()
    for i in data:
        dd = []
        s1 = i.split("\t")
        dd.append(s1[0])
        dd.append(s1[1])
        for j in data[i]:
            s = j.split("\t")
            dd.append(s[1])
        for calls in targetCalls[i]:
            dd.append(calls)
        matchlit = listAddElemane+dd
        targetList = ["Well Position","Sample Name","FluA","FluB","Cov","RSVAB","RNASEP","FluA","FluB","Cov","RSVAB","RNASEP"]
        headers = listAddHeader + targetList
        df1 = pd.DataFrame([matchlit],[s1[0]],columns=headers)
        #print(df1)
        df = pd.concat([df1,df])

    df = df[::-1]
    #print(df)
    return df


    

# %%
dataFrameNew = pd.DataFrame()
inputFolder = "D:/Covid_data_New/Yr2023/Feb/Feb16/"
filesInDirectory = listdir(inputFolder)
for files in filesInDirectory:
    if "xlsx" in files:
        try:
            exportedFile = pd.read_excel(inputFolder+"/"+files,sheet_name = "Primary_result",header=None, engine="openpyxl")
            metaDF, export = readExcelFiles(exportedFile)
            listAddElemane, listAddHeader = getmetaDataInfo(metaDF)
            datFrame = implementLogic(export,listAddElemane,listAddHeader)
            dataFrameNew = pd.concat([datFrame,dataFrameNew])
        except:
            print("No such sheet in file, skipping " + files)
        
writer = pd.ExcelWriter("D:/Covid_data_New/Yr2023/Feb/Feb16/aa.xlsx")
dataFrameNew.to_excel(writer,sheet_name="Test",engine='xlsxwriter',index=False)
writer.close()


