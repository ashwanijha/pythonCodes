# %% [markdown]
# Perl map functions

# %%
def solution(A,B):
    lengthA = len(A)
    hashA = {}
    hashB = {}
    j=0
    k=0
    sumA = 0
    sumB = 0
    for i in A:
        sumA+=i
        hashA[j]=sumA
        j+=1
    for i in B:
        sumB+=i
        hashB[k]=sumB
        k+=1
    allValues = []
    for i in hashA:
        print(i)
        if hashA[i] == hashB[i]:
            allValues.append(i)
        
    fair = min(allValues)
    print(hashA,hashB,fair)

A = [3, 2, 6]
B = [4, 1, 6]
solution(A,B)


