from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from getpass import getpass

# credentials retrieval
usr = input('Username: ')
pwd = getpass('Password: ')
dfa = input("Clave 2FA: ")

# driver setup
s = Service("..\Chrome\chromedriver.exe")

# chrome options to run on headless mode and no logging on the terminal
chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
#chrome_options.add_argument("--headless")
chrome_options.add_argument("--log-level=3")

# instantiate the webdriver
driver = webdriver.Chrome(service=s, options=chrome_options)
driver.implicitly_wait(40)

#driver = webdriver.Chrome('..\Chrome\chromedriver.exe')
driver.get("https://virtual.upb.edu.co/d2l/login?noRedirect=1")

# # login
#username = driver.find_element_by_id("userName")
#password = driver.find_element_by_id("password")

username = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "userName"))
        )
password = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "password"))
        )

username.send_keys(usr)
#password.send_keys(pwd)
password.send_keys("2/2*2Hijosmemsam")
password.send_keys(Keys.RETURN)
sleep(1)

#2FA
driver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")

l2fa = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "z_d")))
l2fa.send_keys(dfa)
l2fa.send_keys(Keys.RETURN)

# read the URLs using pandas
data = pd.read_csv('test.csv')
data = data[['Enlace curso']]
course_info = []
for index, row in data.iterrows():
    url = row['Enlace curso']
    print(url)
    id = url[-5:]
    if id.isdigit():
        new_url = 'https://virtual.upb.edu.co/d2l/lp/manageCourses/course_offering_info_viewedit.d2l?ou' + id
    else:
        id = url[-4:]
        new_url = 'https://virtual.upb.edu.co/d2l/lp/manageCourses/course_offering_info_viewedit.d2l?ou' + id

    try:
        driver.get(new_url)
        short_name = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "z_l")))
        # short_name = driver.find_element_by_id("z_l")
        short_name_value = short_name.get_attribute('value')
        course_info.append((short_name_value, short_name_value[-5:]))
    except Exception as e:
        print(e)

df = pd.DataFrame(course_info, columns=['Nombre', 'NRC'])
df.to_csv('shortnames.csv', index=False)
driver.close()
