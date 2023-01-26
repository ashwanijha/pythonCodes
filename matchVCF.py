# %%
import sys
import re

score = []
with open("match.txt") as snpFile:
    for line1 in snpFile:
        score.append(line1.rstrip())
#print (score)


with open("test.vcf") as file:
    for line in file:
        if(re.match("#",line)):
            pass
        else:
            splitLine=line.rstrip().split("\t")
            if(splitLine[2] in score):
                print (line.rstrip()+"\t"+splitLine[2])



# %%
import pandas as pd
dataframe = pd.read_excel("D:/Covid_data_New/Yr2023/Jan/Jan3/Excel/2022-11-22_TaqPath C19_Flu_RSV Select_LOD_FMT_LOT-2_PLATE-2_QS7_0019_20230103 102319.xlsx",sheet_name="Amplification Data"
,skiprows=23)



# %%
#print(dataframe['Well Position']+dataframe['Target'])
#print(dataframe['dRn'])

cond = 1.5
while cond < 27:
    print (cond)
    cond +=1

# use of range() to define a range of values
values = range(4)

# iterate from i = 0 to i = 3
for i in values:
    print(i)



