from selenium import webdriver
import pandas as pd

chromedriver = "C:\\webdrivers\\chromedriver.exe"
#creating an instance of Chrome Webdriver
driver = webdriver.Chrome(chromedriver)


'''
A method to login with the user login credentials 
'''
def uniAssist_login():
    #enter username and password of the authorized user below
    driver.get("https://ww2.uni-assist.de/portal/index.php?go=sea")
    driver.find_element_by_id("login").send_keys("SOME_USERNAME")
    driver.find_element_by_id("pass").send_keys("SOME_PASSWORD")
    #since "Anmelden" button has no ID, the xpath was used to find the button element
    driver.find_element_by_xpath("//*[@id=\"content\"]/form/div/input[3]").click()


'''
A method to read the data from an excel file and extract the first column which stores the applicant number
Note: installed "xlrd" package
Note: header=None because we need to consider the column header as a value
'''
def uniAssist_getApplicantNumberFromExcel():
    name = ['BWID']
    # df = pd.read_excel('C:\\Users\\Ashish\\Downloads\\uni-assist_applicants.xlsx', sheet_name='Mathe (until 6-14)',
    #                    header=None,
    #                    encoding='utf-8', names=name)
    df = pd.read_excel('C:\\Users\\Ashish\\Downloads\\uni-assist_applicants.xlsx', sheet_name='Tabelle1',
                       header=None,
                       encoding='utf-8', names=name)
    # df = pd.read_excel('C:\\Users\\mages02\\Downloads\\uni-assist_applicants.xlsx', sheet_name='Tabelle1',
    #                     header=None,
    #                     encoding='utf-8', names=name)
    for item in df.index:
        uniAssist_applicantProfile(str(df['BWID'][item]))
        uniAssist_downloadFile(str(df['BWID'][item]))


'''
A method to navigate to the URL page (where the Applicant files are stored)after logging in
'''
def uniAssist_applicantProfile(applicant_number):
    driver.get("https://ww2.uni-assist.de/portal/index.php?go=bewfiles&bew={}".format(applicant_number))


'''
A method to navigate to the pdf file to be downloaded and 
'''
def uniAssist_downloadFile(applicant_number):
    #driver.find_element_by_xpath("//a[contains(@href, '&bewid=2024265') and contains(@href, 'index.php?go=pdf&do=join&id=')]").click()
    driver.find_element_by_xpath(
        "//a[contains(@href, '&bewid={}') and contains(@href, 'index.php?go=pdf&do=join&id=')]".format(applicant_number)).click()
    #driver.find_element_by_xpath(
    #    "//a[contains(@href, '&bewid=2024265') and contains(@href, 'index.php?go=pdf&do=joinonlinefiles')]").click()
    driver.find_element_by_xpath(
        "//a[contains(@href, '&bewid={}') and contains(@href, 'index.php?go=pdf&do=joinonlinefiles')]".format(applicant_number)).click()


'''
A method to log out of the uni-assist portal after the files have been downloaded
'''
def uniAssist_logout():
    driver.find_element_by_xpath("//a[contains(@href, 'index.php?go=logout')]").click()


'''
method calls
'''
uniAssist_login()
uniAssist_getApplicantNumberFromExcel()
uniAssist_logout()



