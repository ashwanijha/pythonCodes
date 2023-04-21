# %%
import pandas as pd
inputFolder = "D:/Covid_data_New/Yr2023/4April/14April/PKK/"
annotationFile = pd.read_excel(inputFolder+"/Combined_MLAnnotationData_AllDetails_20230414_1136AM.xlsx")


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
    
    annotationFile.at[i,"Annotator"] = "Ashwani"
        
writer = pd.ExcelWriter(inputFolder+"/Combined_MLAnnotationData_AllDetails_20230414_AJ.xlsx")
annotationFile.to_excel(writer,sheet_name="data",index=False)
writer.close()



