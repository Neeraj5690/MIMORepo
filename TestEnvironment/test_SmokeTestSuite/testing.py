#-----------To check any error on the screen---------------------
# try:
#     WebDriverWait(driver, SHORT_TIMEOUT
#                   ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
#
#     WebDriverWait(driver, LONG_TIMEOUT
#                   ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
# except TimeoutException:
#     pass
# try:
#     time.sleep(2)
#     bool1 = driver.find_element(By.XPATH, "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1").is_displayed()
#     if bool1 == True:
#         ErrorFound1 = driver.find_element(By.XPATH,
#             "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").text
#         print(ErrorFound1)
#         driver.find_element(By.XPATH,
#             "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[2]/div/button").click()
#         TestResult.append(PageName + " not able to open\n" + ErrorFound1)
#         TestResultStatus.append("Fail")
#         bool1 = False
# except Exception:
#     try:
#         time.sleep(2)
#         bool2 = driver.find_element(By.XPATH,
#             "//div[@class='MessageLayout---message MessageLayout---error']").is_displayed()
#         if bool2 == True:
#             ErrorFound2 = driver.find_element(By.XPATH,
#                 "//div[@class='MessageLayout---message MessageLayout---error']/div/p").text
#             print(ErrorFound2)
#             TestResult.append(PageName + " not able to open\n" + ErrorFound2)
#             TestResultStatus.append("Fail")
#             bool2 = False
#     except Exception:
#         pass
#     pass
#-----------------------------------------------------------------------------------
import os
import sys

# from selenium import webdriver
# from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
# from TestEnvironment.GlobalLoader.Loader import LoaderCls
# print("111111111111"+DataReadMaster.GlobalData("test_AllModulesAdmin","Directory"))
# DataReadMaster.GlobalData("test_AllModulesAdmin","Directory")
# driver = webdriver.chrome(executable_path=DataReadMaster.GlobalData("GlobalData", "ChromePath"))
# LoaderCls.LoaderMeth(driver)
#-----------------------------------------------------------------------------------

# import sys
# print(sys.path[1])
# print(ROOT_DIR)
# ROOT_DIR=os.path.normpath(os.getcwd() + os.sep + os.pardir + os.sep + os.pardir)
# newPath = ROOT_DIR.replace(os.sep, '/')
# print(newPath)
#-----------------------------------------------------------------------------------
from selenium import webdriver
#ChromeCls.ChromeMeth()
#-----------------------------------------------------------------------------------

#def send_email(user, pwd, recipient, subject, body):
# import smtplib
# FROM = "neeraj1wayitsol@gmail.com"
# TO = ['neerajpebma0@gmail.com'] \
# #if isinstance(recipient, list) else [recipient]
# SUBJECT = "subject"
# TEXT = "body"
# # Prepare actual message
# message = """From: %s\nTo: %s\nSubject: %s\n\n%s
# """ % (FROM, ", ".join(TO), SUBJECT, TEXT)
# try:
#     server = smtplib.SMTP("smtp.gmail.com", 587)
#     #server = smtplib.SMTP('192.168.1.12', 587)
#
#     server.ehlo()
#     server.starttls()
#     server.login(FROM, "aaa")
#     server.sendmail(FROM, TO, message)
#     server.close()
#     print("successfully sent the mail")
# except  Exception as ert:
#     print(ert)
#     print ("failed to send mail")
#-----------------------------------------------------------------------------
import ast

# x = '[ "A","B","C" , " D"]'
# y=DataReadMaster.GlobalData("GlobalData", "EmailTo1")
# x = ast.literal_eval(y)
# print(x)
#-------------------------------------------------------------------------

# abcc=SafeToElementActionCls.SafeToElementActionMeth(SafeToVerify,"test_AllModulesAdmin", "SafeToHomePropertyClick")
# print(abcc)
# if abcc == "Yes":
#     print("This is correct")
# else:
#     print("This is Incorrect")
# start = 'of '
# end = ''
# s = '1 – 10 of 11'
# print (s[s.find(start)+len(start):s.rfind(end)])
#---------------------------------------------------------------------------------------
import os

import sys
if "C:/Users/Neeraj/PycharmProjects/MIMO" not in sys.path:
    sys.path.append("C:/Users/Neeraj/PycharmProjects/MIMO")

from TestEnvironment.test_SmokeTestSuite import LatestChrome2
from chrome import latestChrome1
def test():
    print("This is test internal")
    LatestChrome2.chrome_meth()
    latestChrome1.chrome_meth()
if __name__=='__main__':
    test()

#chrome_cls.chrome_meth()

