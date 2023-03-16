# %%
import pandas as pd
from os import listdir
from statistics import mean, stdev

target=["ABY","ALEXA 647","FAM","JUN","VIC"]
datF = pd.DataFrame()
inFolder = "D:/Covid_data_New/Yr2023/March/16March/formean"
allFiles = listdir(inFolder)

dic = dict()
for files in allFiles:
    
    print(files)
    try:
        file = pd.read_excel(inFolder+"/"+files,sheet_name="Primary_result",skiprows=26)
        file = file.replace(to_replace="UNDETERMINED",value=41)
        for i in file.index:
            samples = file.loc[i,"Sample"]
            #samples = samples.replace(to_rtp)
            if samples == "PC_PC" or samples == "NC_NC":
                continue
            if pd.isnull(samples):
                continue
            s = str(samples)
            s = s.replace(" ","")
            sam = s.split("_")
            reporter = file.loc[i,"Reporter"]
            cq = file.loc[i,"Cq"]
            sampleRep = str(sam[0])+"\t"+str(reporter)
            dic.setdefault(sampleRep,[]).append(cq)
    except:
        print("no file")
for tgt in target:
    for key in dic:
        s = key.split("\t")
        if tgt == s[1]:

            listL = [s[0],s[1],"%.3f" % mean(dic.get(key)),"%.3f" % stdev(dic.get(key))]
            df = pd.DataFrame([listL],[s[1]],["Sample","Dye","Mean","SD"])
            datF = pd.concat([df,datF])
#        print(key)
#print(datF)
datF = datF[::-1]
writer = pd.ExcelWriter(inFolder+"/Combined_MLAnnotationData_AllDetails_20230301_AJ.xlsx")
datF.to_excel(writer,sheet_name="data",index=False)
writer.close()

# %%



datF = pd.DataFrame()
inFolder = "D:/Covid_data_New/Yr2023/March/15March/formean"
allFiles = listdir(inFolder)

for files in allFiles:
    
    print(files)
    analysis(inFolder,files)
   
datF = datF[::-1]
writer = pd.ExcelWriter(inFolder+"/Combined_MLAnnotationData_AllDetails_20230301_AJ.xlsx")
datF.to_excel(writer,sheet_name="data",index=False)
writer.close()




