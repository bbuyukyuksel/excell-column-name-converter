from Convert import *
from time import sleep
if __name__ == '__main__':


    print('''
    multiConv       -> mc
    multiRecovery   -> mr
    ''')

    try:
        select = input("Select Option: ").lower()
        if not select in ['mc','mr']:
            assert ValueError("An Error Occured")

        if (select in "mc"):
            fileName =      input("Existing File Name  (excellFile)\n->") + '.xlsx'
            newFileName =   input("New File Name    (newExcellFile)\n->") + '.xlsx'
            print('\n', '*' * 20, '\n')
            newDataNames =  input("New Data Names   (Ex : 'spam,eggs,foo')\n->")
            indexNames =    input("Index Names      (Ex : 'A,B,C')\n->").upper()

            newDataNames = tuple(newDataNames.split(','))
            indexNames = tuple(indexNames.split(','))

            multiConvert(fileName,newFileName,newDataNames,indexNames)
        elif select in 'mr':
            fileName = input("Existing File Name  :\n->") + '.xlsx'
            indexNames = input("Index Names      (Ex : 'A,B,C')\n->").upper()
            indexNames = tuple(indexNames.split(','))
            multiRecovery(fileName, indexNames)
        else:
            "Finished!"
    except:
        print(sys.exc_info()[1])
        exit()


        #multiConvert("test.xlsx","newExcellFile.xlsx",('TEST','DENEME'),('A','B'))
        #multiRecovery("newExcellFile.xlsx",('A','B'))











