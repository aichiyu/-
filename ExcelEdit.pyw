from openpyxl import load_workbook
import codecs
wb = load_workbook('data.xlsx')
sheet = wb["Sheet1"]
#recov=open('C:\\Users\\Joe\\AppData\\LocalLow\\SogouPY.users\\00000002\\PhraseEdit.txt', 'w+', encoding="utf-16")
with open("PhraseEdit1.txt", 'w+', encoding="utf-16") as txt:
    for i in sheet.rows:
        if i[0].value=='Type，分类用，可以不填':continue
        for j in range(5,len(i)):
            if i[1].value == None or i[j].value == None or i[j].value.startwith('#'): continue
            shift = int(i[4].value) if i[4].value!=None else 0
            name=[str(i[1].value)]+list(str(i[2].value).split(';')) if i[2].value!=None else [str(i[1].value)]
            for k in name:
                txt.write(k + ',' + str(j - 4 + shift) + '=' + str(i[j].value) + '\n')
    #for line in txt:
        #recov.write(line)
#recov.close()
print('Finished')
