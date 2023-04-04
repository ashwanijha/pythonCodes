# %%
import pandas as pd
inputFolder = "D:/Covid_data_New/Yr2023/March/20March/Res/"
annotationFile = pd.read_excel(inputFolder+"/Combined_MLAnnotationData_AllDetails_20230320_1227PM.xlsx")


# %%
for i in annotationFile.index:
    splitPattern = "_20230126"
    if "_20230320" in annotationFile.loc[i,"ExcelFileName"]: 
        splitPattern = "_20230320"
    string = annotationFile.loc[i,"ExcelFileName"].split(splitPattern)
    wellAnnotation = annotationFile.loc[i,"Well_Position"]
    targetAnnotation = annotationFile.loc[i,"Target"]
    csvFile = string[0]+".csv"
    #print(csvFile)
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

    #
    if annotationFile.loc[i,"AmpType (Clear|Strong|Weak|NoAmp)"] == "noAmp":
        annotationFile.at[i,"AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Non_Amp"
    else:
        annotationFile.at[i,"AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Amp"
    
    annotationFile.at[i,"Annotator"] = "Puneet"
        
writer = pd.ExcelWriter(inputFolder+"/Combined_MLAnnotationData_AllDetails_20230307_PK.xlsx")
annotationFile.to_excel(writer,sheet_name="data",index=False)
writer.close()



