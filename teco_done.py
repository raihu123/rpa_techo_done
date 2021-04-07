from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import xlrd
import datetime
import os


workbook = xlrd.open_workbook('Milestone_TECO_Done.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')

file = open("credentials.txt", "r")
cred=file.read().splitlines()

passw = cred[1]
user = cred[0]

count=0
fail=0
list=[]

current_dir = os.path.dirname(os.path.abspath(__file__))
driver_location = current_dir + '\\chromedriver.exe'
print(driver_location)

# for i in range(1,worksheet.nrows):

    # try:
driver = webdriver.Chrome(driver_location)
driver.get("http://rocapp.robi.com.bd/ROC/ROCPages/Pages/ROBIROC_Login.aspx")
driver.maximize_window()
driver.implicitly_wait(15)


elem = driver.find_element_by_id('txtUsername')
elem.send_keys(user)
elem = driver.find_element_by_id('txtPassword')
elem.send_keys(passw)

elem = driver.find_element_by_xpath('//*[@id="btnLogin"]').click()

time.sleep(3)

try:

    elem = driver.switch_to.alert.accept()

except:

    driver.refresh()


elem = driver.find_element_by_xpath("""//*[@id="nav"]/li[4]/a""").click()
elem = driver.find_element_by_xpath("""//*[@id="nav"]/li[4]/ul/li[1]""").click()

########################### LOOP ##################

for i in range(1,worksheet.nrows):

    # try:
    wr = worksheet.cell(i, 0).value

    print('Attempting Entry:', wr)

    elem = driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_txtWRNO"]""")
    elem.clear()
    elem.send_keys(wr)

    # elem = driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_pnlSearch"]/table/tbody/tr[5]/td[6]""").click()

    elem = driver.find_element_by_xpath('//*[@value="  Search  "]').click()

    name = '//*[text()="' + wr + '"]'
    # print(name)

    elem = driver.find_element_by_xpath(name).click()

    # elem =driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_gdvWFList"]/tbody/tr[2]/td[1]""").click()

    # text_found = re.search(r'TECO', src)
    # print(text_found)

    tt = driver.find_element_by_link_text("TECO")
    tt.click()

    ed = worksheet.cell(i, 1).value

    if ed == 'N/A':
        elem = Select(driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_drpMILESTONE_STATUS"]"""))
        elem.select_by_value("Not Applicable")

        elem = driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_txtAddMilestoneComments"]""")
        elem.send_keys("N/A")





    else:

        driver.switch_to.frame("frmMilestoneAction")
        elem = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtTecoDoneDate')
        elem.clear()

        dt = datetime.datetime(*xlrd.xldate_as_tuple(ed, workbook.datemode))

        a = dt.strftime("%d")
        b = dt.strftime("%b")
        c = dt.strftime("%y")

        str = [a, b, c]

        string = '-'.join(str)

        elem.send_keys(string)
        elem.send_keys(Keys.RETURN)

        ############################## SAVE BUTTON for date #######################
        elem = driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_btnSave"]""").is_enabled()

        if elem:
            elem = driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_btnSave"]""").click()

            time.sleep(4)
            elem = driver.switch_to.alert
            print(elem.text)
            list.append(wr)
            list.append('[Considered as passed]')
            list.append(elem.text)
            elem.accept()
        # elem = driver.switch_to.alert
        # elem.dismiss()
        # print(elem.text)


        ###########################################################################

        driver.switch_to.parent_frame()
        elem = Select(driver.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_drpMILESTONE_STATUS"]"""))
        elem.select_by_value("Done")

        #################### SAVE and SUBMIT Button##########################
    # elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSaveMilestone"]').click()

    elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSaveMilestone"]').is_enabled()
    # print(elem)
    if elem:
        elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSaveMilestone"]').click()
        time.sleep(6)
        print('Successful Entry.')
        count += 1
        try:
            driver.switch_to.alert.accept()
        except:
            pass

    else:
        print('failed')
        list.append(wr)
        fail += 1

        driver.refresh()
        driver.get("http://rocapp.robi.com.bd/ROC/Pages/WorkFlow/WFMaster.aspx")
        driver.refresh()


    # except:
    #     print('failed')
    #     list.append(wr)
    #     fail += 1
    #
    #     # driver.refresh()
    #     driver.get("http://rocapp.robi.com.bd/ROC/Pages/WorkFlow/WFMaster.aspx")
    #     driver.refresh()




        # driver.get("http://rocapp.robi.com.bd/ROC/Pages/WorkFlow/WFMaster.aspx")
    # except:
    #     print('failed')
    #     list.append(wr)
    #
    #     driver.close()
    #     fail+=1
driver.close()

print('Entries done ', count)
print('Failed ', fail)
print(list)
print('creating log')
with open("log.txt", "w") as text_file:
    text_file.write("\nEntries: %s" % count)
    text_file.write("\nFail: %s" % fail)
    for item in list:
        text_file.write("\n%s\n" % item)
print('log done')














