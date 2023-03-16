import os
import openpyxl
import socket
class DataReadMaster:
    ROOT_DIR = None
    print(socket.gethostname())
    if socket.gethostname()=="DESKTOP-KMS7763":
        ROOT_DIR = "C:/Users/gagandeep.singh_bits/PythonWorkSpace/MIMORepo/"
    newPath = ROOT_DIR.replace(os.sep, '/')
    Path=newPath
    ExcelFileName = "MasterDataFile"
    locx = (Path + 'TestEnvironment/MasterData/' + ExcelFileName + '.xlsx')
    wbx = openpyxl.load_workbook(locx)

    @classmethod
    def GlobalData(cls,Sheetname,FieldName):
        sheetx = DataReadMaster.wbx[Sheetname]
        for ix in range(1, 200):
            if sheetx.cell(ix, 1).value == None:
                break
            else:
                if sheetx.cell(ix, 1).value == FieldName:
                    return sheetx.cell(ix, 2).value