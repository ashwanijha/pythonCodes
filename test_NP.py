# %%
import pandas as pd
infile = "D:/Annotations_automation/ons/"
x= pd.read_excel(infile+"/24MAR2023_SFSH_RRO_ST3_D1_OP1_QS5Dx_INST1_Copy_20230725_112659.xlsx",
                 sheet_name="Results",header=None)

# %%
metaData = x[:23]
completeData = x[24:]
completeData1 = pd.DataFrame(completeData.values[1:],columns= completeData.iloc[0])

mergedData = dict()
colNames = completeData1.columns
otherHeader = [colNames[1],colNames[2],"FLUA_CqValue","FLUB_CqValue", "COV19_CqValue",
               "RSVAB_CqValue",	"RNASEP_CqValue"]
#print(otherHeader)
for l in completeData1.index:
    wellPos = completeData1.loc[l,'Well Position']
    sampleName = completeData1.loc[l,'Sample']
    valueCq = completeData1.loc[l,'Cq']
    valueMerged = str(wellPos)+"\t"+str(sampleName)
    mergedData.setdefault(valueMerged,[]).append(valueCq)

metaD = []
header = []
j=0
for i in metaData[0]:
    metaD.append(metaData[1][j])
    header.append(metaData[0][j])
    j+=1

header.extend(otherHeader)

print(header)
dataFrame = pd.DataFrame()

data = []
bigDF = pd.DataFrame()
for da in mergedData:
    dataForInfo = []
    
    tst = da.split("\t")
    dataForInfo.extend(metaD)
    dataForInfo.append(tst[0])
    dataForInfo.append(tst[1])
    dataForInfo.extend(mergedData[da])
    datF = pd.DataFrame([dataForInfo],columns=[header])
    bigDF = pd.concat([datF,bigDF])
    j+=1

    #bigDF = pd.concat([getInformation(metaData,completeData1,listName),bigDF])
   # print(da,mergedData[da])
#print(datF)    
#print(j)


dataFrame = pd.DataFrame([dataForInfo],[j])
#print(bigDF)
bigDF = bigDF[::-1]
excelWriter = pd.ExcelWriter(infile+"/test.xlsx")
bigDF.to_excel(excel_writer=excelWriter,sheet_name= "Data",index=False)
excelWriter.close()



