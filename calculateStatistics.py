# %%
import pandas as pd
from statistics import mean, stdev
import re

ofile = pd.ExcelWriter("D:/Covid_data_New/Yr2023/Feb/Feb7/aa.xlsx")
mypath = "D:/Covid_data_New/Yr2023/Feb/Feb7/10OCT2022_StarFish_Fresh_Frozen_Cov_19_KA_QS5-0046_Analyzed_20230207 113218.xlsx"
diomniExport = pd.read_excel(mypath,sheet_name = "Results",skiprows= 23)
diomniExport = diomniExport.replace(to_replace = "Undetermined", value = 41)

data = dict()
for i in diomniExport.index:
    sample = diomniExport.loc[i,"Sample"]
    if pd.isna(sample):
        continue

    #m = re.search(r'\d+$', sample)
    if re.search(r'\d+$', sample):
        sample = sample[:-1]
    
    reporter = diomniExport.loc[i,"Reporter"]
    cq = diomniExport.loc[i,"Cq"]
    merge = str(sample)+"\t"+str(reporter)
    if sample == "PC" or sample == "NC":
        continue

    if cq != 41:
        data.setdefault(merge,[]).append(cq)
    else:
        data.setdefault(merge,[]).append(0)
dataF = pd.DataFrame()
for d in data:
    s = d.split("\t")
    listL = ["%.3f" % mean(data[d])]
    dd = [s[0],s[1]]  
    df = pd.DataFrame(listL,[s[1]],[s[0]])
    dataF = pd.concat([df,dataF])
#print(dataF)
dataF.to_excel(ofile,sheet_name="Test Sheet")
ofile.close()




