from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import glob
import pandas as pd


path = r'C:\Users\BDG-BSALAMANCA\Downloads\consultaSisben\consultaSisben' # use your path
all_files = glob.glob(path + "/*.csv")
lista = {}
for filename in all_files:
    df = pd.read_csv(filename, index_col=None, header=None)
    cedulas = df[0]
    for i in cedulas.index:

        driver = webdriver.Chrome("C:\chromedriver.exe")
        driver.get("https://reportes.sisben.gov.co/dnp_sisbenconsulta")

        time.sleep(3)
        driver.refresh()
        sel = driver.find_element("xpath", "//*/form/div/div/div/div[1]/div/select")

        print("The input Element is: ", sel)
        sel.send_keys("Cédula de Ciudadanía")


        box = driver.find_element("xpath", "//*/form/div/div/div/div[2]/div/input")
        print("The input Element is: ", box)

        box.send_keys(str(cedulas[i]))
        time.sleep(2)
        box.send_keys(Keys.RETURN)
        time.sleep(2)

        driver.save_screenshot(f"./screenshots/{cedulas[i]}.png")

        time.sleep(3)
        try:

            sel2 = driver.find_element("xpath", "//*/div/div/div[2]/div/div[2]/div[3]/div/p")
            lista[cedulas[i]]=sel2.text


        except:
            pass
        driver.close()

df = pd.DataFrame([[key, lista[key]] for key in lista.keys()], columns=['Cedula', 'Calificacion'])
df.to_excel('SisbenCalificacion.xlsx')



