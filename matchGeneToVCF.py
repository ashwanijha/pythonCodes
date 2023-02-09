# %% [markdown]
# Code to match gene name to vcf. Genes are in excel and vcf is a text file.

# %%
import pandas as pd
import re
ensembleMartExportFile = pd.read_excel("D:/Script/mart_export.xlsx")
geneList = []
for i in ensembleMartExportFile.index:
    if pd.isna(ensembleMartExportFile.loc[i,'Gene name']):
        continue
    geneList.append(ensembleMartExportFile.loc[i,'Gene name'])

geneNames = set(geneList)

with open("D:/Script/tnFreebayes.filt.snpEff.reportable.vcf") as vcfFile:
    for vcfContent in vcfFile:
        if re.match("#",vcfContent):
            continue
        infoColumn = vcfContent.split("\t")
        geneInVCF = infoColumn[7].split("|")
        if geneInVCF[3] in geneList:
            print(geneInVCF[3])





