from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as Ec
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from datetime import date
import excel
# import sleep

# Login Details 
username="youremail@gmail.com"
password="your_password"
MeetLink="meet_link"
# username=input("Enter Your Email Id: ")           #these are optional 
# password=input("Enter Your Password: ")
# MeetLink=input("Enter Your Meet Link: ")

option = Options()
option.add_argument("use-fake-ui-for-media-stream")
option.add_argument("disable-notifications")

driver = webdriver.Chrome(executable_path="C:\Drivers\chromedriver.exe",options=option)
# driver = webdriver.Chrome(executable_path="C:\Drivers\chromedriver.exe")
wait = WebDriverWait(driver,60)


# login to google
driver.get("https://www.google.co.in/")
driver.maximize_window()
# driver.minimize_window()
sign = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[1]/div/div/div/div[2]/a")))
sign.click()
email = wait.until(Ec.element_to_be_clickable((By.ID,"identifierId")))
email.send_keys(username)
# print(sign.text)
next = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/div/button/span")))
next.click()

driver.implicitly_wait(4)
passw = wait.until(Ec.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div[1]/div[1]/div/div/div/div/div[1]/div/div[1]/input")))
passw.send_keys(password)

next = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/div/button/span")))
next.click()

# now enter in meet 
try:
    driver.implicitly_wait(10)
    driver.get(MeetLink)
    driver.implicitly_wait(4)
    driver.refresh()
    driver.implicitly_wait(6)
except:
    driver.implicitly_wait(10)
    driver.get(MeetLink)
    driver.implicitly_wait(4)
    driver.refresh()
    driver.implicitly_wait(6)
micoff = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/c-wiz/div/div/div[9]/div[3]/div/div/div[4]/div/div/div[1]/div[1]/div/div[4]/div[1]/div/div/div")))
micoff.click()
cameraOff = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/c-wiz/div/div/div[9]/div[3]/div/div/div[4]/div/div/div[1]/div[1]/div/div[4]/div[2]/div/div")))
cameraOff.click()
join = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/c-wiz/div/div/div[9]/div[3]/div/div/div[4]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/span/span"))).click()
driver.implicitly_wait(2)


# now get number of peoples in meet 
wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/c-wiz/div[1]/div/div[9]/div[3]/div[10]/div[3]/div[2]/div/div/div[2]/span/button/i[1]"))).click()
driver.implicitly_wait(10)
peoples = driver.find_elements_by_class_name("KV1GEc")
total=len(peoples)
print(len(peoples))


# get names of peoples in meet 
names=set()
for i in range (1,total+1):
    temp = wait.until(Ec.element_to_be_clickable((By.XPATH,"/html/body/div[1]/c-wiz/div[1]/div/div[9]/div[3]/div[4]/div[2]/div[2]/div[2]/div[3]/div/div["+str(i)+"]/div[1]/div/div/span[1]")))
    name=temp.text
    names.add(name)

driver.quit()
# close driver 

#now excel work start
today = date.today()
path = 'attendence.xlsx'
rows=excel.getRowCount(path,"Sheet1")
cols=excel.getColCount(path,"Sheet1")
for k in range (1,cols+1):
    if excel.readData(path,"Sheet1",1,k)==None or excel.readData(path,"Sheet1",1,k)=="" or excel.readData(path,"Sheet1",1,k)==" ":
        cols=k-1
        break
for k in range (1,rows):
    if excel.readData(path,"Sheet1",k,1)==None or excel.readData(path,"Sheet1",k,1)=="" or excel.readData(path,"Sheet1",k,1)==" ":
        rows=k-1
        break

print("today= ",today)
excel.writeData(path,"Sheet1",1,cols+1,today)
for i in range (2,rows+1):
    print(i)
    excelName=excel.readData(path,"Sheet1",i,1)
    if(excelName in names):
        excel.writeData(path,"Sheet1",i,cols+1,"P")
    else:
        excel.writeData(path,"Sheet1",i,cols+1,"A")
print("DONE","rows=",rows," cols=",cols,"writeCol=",cols+1)