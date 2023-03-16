import requests
import wget
import zipfile
import os
import sys
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster

class ChromeCls:
    sys.path.append("/chrome")
    ROOT_DIRChrCls = "C:/Users/Neeraj/PycharmProjects/MIMO"
        #sys.path[1]
    NewChromePathChrCls = ROOT_DIRChrCls.replace(os.sep, '/')
    NewChromePathChrCls = NewChromePathChrCls + "/chrome"
    NewChromePath1ChrCls = NewChromePathChrCls + "/chromedriver.exe"
    print(NewChromePath1ChrCls)

    @classmethod
    def ChromeMeth(cls):
        ROOT_DIRChrCls1 = DataReadMaster.ROOT_DIR
        NewChromePathChrCls1 = ROOT_DIRChrCls1.replace(os.sep, '/')
        NewChromePath1ChrCls1 = NewChromePathChrCls1 + "/chrome"
        url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
        response = requests.get(url)
        version_number = response.text
        download_url = "https://chromedriver.storage.googleapis.com/" + version_number +"/chromedriver_win32.zip"
        latest_driver_zip = wget.download(download_url,'chromedriver.zip')
        with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
            zip_ref.extractall(NewChromePath1ChrCls1)
        os.remove(latest_driver_zip)


def chrome_meth():
    aa="Testing"
    print(aa)
    return None