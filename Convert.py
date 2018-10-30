from ExcellReader import *
from pprint import pprint
import json
import sys
from time import sleep

def multiConvert(filename,newFileName,nameOptions,indexOptions):
    try:
        if not(len(nameOptions) == len(indexOptions)):
            raise ValueError("Len Error")
    except:
        print(sys.exc_info()[1])
        return
    for i in range(len(nameOptions)):
        if i > 0:
            filename = newFileName
        convert(filename, newFileName, nameOptions[i], indexOptions[i])
        sleep(0.05)

def multiRecovery(filename,indexOptions):
    for i in range(len(indexOptions)):
        recovery(filename,indexOptions[i])



def recovery(fileName,Index):
    ##Geriye Dondurme
    ws = excellReader(fileName)
    row, col = ws.getExcellSize()
    FROM = ws.createIndex(Index, 1)
    TO = ws.createIndex(Index, row)
    # Excell File Okundu
    ws.readExcellArray(FROM, TO)
    # Sutunlardan Unique Elemanlar Cekildi
    header = ws.getUniqueColumnElements()[0]
    ws.styledArray()
    #Open JSON File
    with open(header + '.json', 'r') as f:
        UpdatedNames = json.load(f)

    print("Read!")
    pprint(UpdatedNames, indent=4)

    #Replace
    r, c = ws.getExcellArraySize()
    ExcellArray = ws.getExcellArray()
    for i in range(1, r):
        data = ExcellArray[i][0]
        if data == None:
            continue
        if data in UpdatedNames[header]['newNames']:
            index = UpdatedNames[header]['newNames'].index(data)
            #print("index", index)
            #print("Data :", data)
            #print("OldName:", UpdatedNames[header]['oldNames'][index])
            #print("NewName:", UpdatedNames[header]['newNames'][index])
            #print("Array :", ExcellArray[i][0])
            #print('*' * 15)
            #print("[{:.>5}-{:.>5}]{:_>15} > {:_>15}\r".format(i,r,data, UpdatedNames[header]['newNames'][index]))
            print("[{:.>5}-{:.>5}]{:_>15} > {:_>15}\r".format(i,r-1,data, UpdatedNames[header]['oldNames'][index]))
            ws.setExcellArray([UpdatedNames[header]['oldNames'][index]], i)
        ws.saveExcellArray(fileName)
        ws.close()

def convert(fileName,newFileName,dataName,Index):
    ws = excellReader(fileName)
    row, col = ws.getExcellSize()
    FROM = ws.createIndex(Index, 1)
    TO = ws.createIndex(Index, row)
    # Excell File Okundu
    ws.readExcellArray(FROM, TO)
    # Sutunlardan Unique Elemanlar Cekildi
    header = ws.getUniqueColumnElements()[0]
    # Guncellenen İsimler Hakkında JSON File Oluşturuldu
    # Guncelleme Indis İsimleri Olarak Test Verildi
    UpdatedNames = ws.getRenamedJSONFile(dataName)
    pprint(UpdatedNames, indent=4)
    # JSON file Test İndex İsimleri İle Yazıldı.
    ws.saveRenamedJSONFile(dataName)
    ws.styledArray()
    # Arama Gerçekleştirildi

    r, c = ws.getExcellArraySize()
    ExcellArray = ws.getExcellArray()
    for i in range(1, r):
        data = ExcellArray[i][0]
        if data == None:
            continue
        if data in UpdatedNames[header]['oldNames']:
            index = UpdatedNames[header]['oldNames'].index(data)
            #print("index", index)
            #print("Data :", data)
            #print("OldName:", UpdatedNames[header]['oldNames'][index])
            #print("NewName:", UpdatedNames[header]['newNames'][index])
            #print("Array :", ExcellArray[i][0])
            #print('*' * 15)
            print("[{:.>5}-{:.>5}]{:_>15} > {:_>15}\r".format(i,r-1,data, UpdatedNames[header]['newNames'][index]))
            ws.setExcellArray([UpdatedNames[header]['newNames'][index]], i)
        ws.saveExcellArray(newFileName)
        ws.close()
