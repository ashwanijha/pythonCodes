ma# %%
import pandas as pd
inputFolder = "D:/Covid_data_New/Yr2023/Feb/Feb21/Res"
annotationFile = pd.read_excel(inputFolder+"/Combined_MLAnnotationData_AllDetails_forMyPython.xlsx")


# %%
for i in annotationFile.index:
    string = annotationFile.loc[i,"FileName"].split(".eds")
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
    except:
        annotationFile.at[i,"Clean|ProblemSevere|ProblemLite"] = "Clean"

    if annotationFile.loc[i,"AmpType (Clear|Strong|Weak|NoAmp)"] == "NoAmp":
        annotationFile.at["AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Non_Amp"
    else:
        annotationFile.at["AmpStatus (Amp|Non_Amp|Unknown) -Annotator"] = "Amp"
    
    annotationFile.at["Annotator"] = "Ashwani"
        
writer = pd.ExcelWriter(inputFolder+"/testing.xlsx")
annotationFile.to_excel(writer,sheet_name="Sheet new",index=False)
writer.close()



