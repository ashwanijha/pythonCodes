# %%
import pandas as pd
import warnings

inputFolder = "D:/Covid_data_New/Yr2023/5May/5May/Res/" # Outfolder of Hari's Script
inputFolderExcels = "D:/Covid_data_New/Yr2023/5May/5May/Excels"  # DA exported excels

def wellResultFile(inputFiles,inputFolderExcels):
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        excelFiles = pd.read_excel(inputFiles, sheet_name="Amplification Data",header=None,engine="openpyxl")
        edsFile = excelFiles[:1]
        edsFileName = edsFile[1][0].split("\\")
        csvFile = edsFileName[-1].split(".eds")
        ofile = open(inputFolderExcels+"/"+csvFile[0]+".csv","w")
        ofile.write("Well,Target\n")
        completeData = excelFiles[23:]
        
        completeData1 = pd.DataFrame(completeData.values[1:],columns= completeData.iloc[0])
        wellPositon = {}
        for i in completeData1.index:
            cycleNumber = completeData1.loc[i,"Cycle Number"]
            dRN = float(completeData1.loc[i,"dRn"])
            if cycleNumber == 1:
                if dRN < -20000:
                    wellTarget = completeData1.loc[i,"Well Position"]+"\t"+completeData1.loc[i,"Target"]
                    wellPositon[wellTarget] = dRN

        for i in wellPositon:
            a = i.split("\t")
            ofile.write(a[0]+","+a[1]+"\n")

        ofile.close()



# %%
import os

files = os.listdir(inputFolderExcels)
for file in files:
    if "xlsx" in file:
        excelFile = inputFolderExcels+"/"+file
        wellResultFile(excelFile,inputFolder)

# %%
annotationFiles = os.listdir(inputFolder)
annotationFileName = ""
for aFile in annotationFiles:
    if "AllDetails" in aFile:
        annotationFileName = aFile

outAnnotationFileName = annotationFileName[0:annotationFileName.rindex("_")]+"_AJ.xlsx"


annotationFile = pd.read_excel(inputFolder+"/"+annotationFileName)


# %%
for i in annotationFile.index:
    #splitPattern = "_20230329"
    splitPattern = ".eds"
    string = annotationFile.loc[i,"FileName"].split(splitPattern)
    wellAnnotation = annotationFile.loc[i,"Well_Position"]
    targetAnnotation = annotationFile.loc[i,"Target"]
    csvFile = string[0]+".csv"
    try:
        csvFileReader = pd.read_csv(inputFolder+"/"+csvFile)
        match = 0
        for j in csvFileReader.index:
            wellCSV = csvFileReader.loc[j,"Well"]
            targetCSV = csvFileReader.loc[j,"Target"]
            if wellAnnotation == wellCSV and targetAnnotation == targetCSV:
                match +=1
                annotationFile.at[i,"Clean|ProblemSevere|ProblemLite"] = "ProblemLite"
                annotationFile.at[i,"Artifact (Bubble|BaselineIssue|WaterFall|SystemError|Smiley|Creeper|CrossTalk|RoosterTail|NoisyBaseline|Abnormal)"] = "NoisyBaseline"
        if match == 0:
            annotationFile.at[i,"Clean|ProblemSevere|ProblemLite"] = "Clean"
    except:
        annotationFile.at[i,"Clean|ProblemSevere|ProblemLite"] = "Clean"

    if annotationFile.loc[i,"AmpType (Clear|Strong|Weak|NoAmp)"] == "noAmp":
        annotationFile.at[i,"AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Non_Amp"
    else:
        annotationFile.at[i,"AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Amp"
    
    if annotationFile.loc[i,"AmpStatus (Amp|NoAmp|Incon) -DA specific"] == "Inconclusive":
        annotationFile.at[i,"AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Amp"
    
    
    annotationFile.at[i,"Annotator"] = "Ashwani"
        
writer = pd.ExcelWriter(inputFolder+"/"+outAnnotationFileName)
annotationFile.to_excel(writer,sheet_name="data",index=False)
writer.close()



