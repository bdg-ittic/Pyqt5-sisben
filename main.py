"""
Created on 22 SEP 2022
@author: BryandSalamanca (bryand.salamanca@bdguidance.com)

"""
import sys
# Importar modulo Qt
from PyQt5 import QtCore, QtWidgets, QtGui
from ui_main import Ui_MainWindow 
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog
from PyQt5.uic import loadUi
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import glob
import os
import pandas as pd


# Importar el código del modulo compilado UI
# Improtar Modulos externos, pyqtgraph, numpy..setCursor


class Principal(QtWidgets.QMainWindow):
    
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self) 
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.a=int()
 
    @QtCore.pyqtSlot()
    def on_pushButton_clicked(self):
        self.fname=QFileDialog.getOpenFileName(self,'Open file', 'D:\codefirst.io\PyQt5 tutorials\Browse Files', 'CSV (*.csv)')
        myFont=QtGui.QFont()
        myFont.setPointSize(12)
        self.ui.label.setText("Archivo abierto: "+str(self.fname))
        fnamestr = str(self.fname)
        todo = len(fnamestr)
        self.rute = fnamestr[2:todo-17]
        self.rutes = fnamestr[2:todo-21]
        self.ui.label.setWordWrap(True)
        self.a=1
        
    @QtCore.pyqtSlot()   
    def on_pushButton_2_clicked(self):
        if self.a == 0:
            self.ui.label.setText("Ingrese archivo .CSV")
            myFont=QtGui.QFont()
            myFont.setBold(True)
            myFont.setPointSize(15)
            self.ui.label.setFont(myFont)
        if self.a == 1:
            self.ui.label.setText("Iniciando proceso, porfavor espere")
            myFont=QtGui.QFont()
            myFont.setBold(True)
            myFont.setPointSize(15)
            self.ui.label.setFont(myFont)
            path = r''+str(self.rute) # use your path
            all_files = glob.glob(""+self.rute)
            lista = {}
            os.makedirs(""+self.rutes+"/screenshots")
            for filename in all_files:
                df = pd.read_csv(filename, index_col=None, header=None)
                cedulas = df[0]
                for i in cedulas.index:
                    driver = webdriver.Chrome("1x/chromedriver.exe")
                    driver.get("https://reportes.sisben.gov.co/dnp_sisbenconsulta")

                    time.sleep(3)
                    driver.refresh()
                    sel = driver.find_element("xpath", "//*/form/div/div/div/div[1]/div/select")

                    sel.send_keys("Cédula de Ciudadanía")


                    box = driver.find_element("xpath", "//*/form/div/div/div/div[2]/div/input")

                    box.send_keys(str(cedulas[i]))
                    time.sleep(2)
                    box.send_keys(Keys.RETURN)
                    time.sleep(2)
                    
                    driver.save_screenshot(self.rutes+f"./screenshots/{cedulas[i]}.png")
                    
                    time.sleep(3)
                    try:

                        sel2 = driver.find_element("xpath", "//*/div/div/div[2]/div/div[2]/div[3]/div/p")
                        lista[cedulas[i]]=sel2.text


                    except:
                        pass
                    driver.close()
            
            time.sleep(3)
            df = pd.DataFrame([[key, lista[key]] for key in lista.keys()], columns=['Cedula', 'Calificacion'])
            df.to_excel(self.rutes+"/SisbenCalificacion.xlsx")
            self.a=0
            time.sleep(3)
        
def main():
    """Corre la App"""
    # Nuevamente, esto es estándar, será igual en cada
    # aplicación que escribas
    app = QtWidgets.QApplication(sys.argv) 
    # Se crea una instancia de la clase
    ventana = Principal()
    # Se muestra el elemento en pantalla
    ventana.show()
    # Se ejecuta y expera a que termine la aplicación
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
