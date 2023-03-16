import os
import openpyxl

class DataReadMaster:
    # ROOT_DIR = sys.path[1]
    ROOT_DIR="C:/Users/Neeraj/PycharmProjects/MIMO/"
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