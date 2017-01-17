#!/usr/bin/env python
#coding=utf-8

# Tämä skripti on massapostitusohjelman pääskripti. 
# Skripti sisältää graafisen käyttöliittymän luonnin ja myös ohjelman toiminnallisuuden
# Ohjelma itsessään käyttää Googlen avointa gmail-SMTP palvelua sisäänkirjautumiseen ja viestien lähetykseen.
# Ohjelma käyttää Excel-taulukoita hakeakseen osoitteita joille haluat lähettää sähköpostia.
# Käyttäjä syöttää oikeat solut ja lisää viestin, jonka jälkeen ohjelma hoitaa lähettämisen.

__author__ = "Antti Pakkanen"
__copyright__ = "Copyright 2016, Rapuposti"
__credits__ = ["Antti Pakkanen", "Hannu Santti", "Lauri Ståhlberg"]
__license__ = "GPL"
__version__ = "Alpha 0.9"
__date__ = "24.11.2016"
__maintainer__ = "Antti Pakkanen"
__email__ = "antti.pakkanen@edu.turkuamk.fi"
__status__ = "Production"

import sys
import os
import smtplib
from PyQt4 import QtGui, QtCore
from openpyxl import load_workbook

# Luodaan pohja PyQt -moduulille
app = QtGui.QApplication(sys.argv)

# Kirjautumisikkuna
class Login(QtGui.QDialog):

    # Luodaan näkymä virheilmoitukselle
    err = QtGui.QWidget()

    # Korjaa sanity check nr.4
    try:
        smtpObj = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    except:
        QtGui.QMessageBox.critical(err,"Error", "Verkkoyhteys ei onnistunut. Tarkista yhteys ja suorita ohjelma uudelleen.")
        sys.exit()


    gmailAcc = ""

    def __init__(self, parent=None):
        super(Login, self).__init__(parent)

        # Asetetaan ikkunan sijanti(50,50) ja koko(600, 500)
        # Asetetaan ikkunan sijanti(50,50) ja koko(600, 500)
        self.setGeometry(50, 50, 300, 300)
        # Ikkunan title ja kuvake
        self.setWindowTitle("Rapuposti")
        self.setWindowIcon(QtGui.QIcon('blizzard.PNG'))
        # Luodaan infokenttä
        self.info = QtGui.QTextBrowser()
        infotext = "Hei, tervetuloa käyttämään Rapupostia! \n\nOhjelmalla saa lähetettyä massasähköpostia kuten uutislehtiä. Ohjelmaan ladataan omatekoinen excel-taulukko, joka sisältää listan sähköpostiosoitteita. Lopuksi ohjelma luo kopion käytetystä taulukosta ja merkitsee kenelle sähköpostit lähetettiin. \n\n HUOM! Ohjelma tukee vain gmail-tunnuksia!"
        infotext = infotext.decode('utf-8')
        self.info.setText(infotext)
        # Luodaan tekstikentät salasanalle ja sähköpostille
        self.textName = QtGui.QLineEdit(self)
        self.textPass = QtGui.QLineEdit(self)
        # Piiloitetaan salasanakentän kirjaimet
        self.textPass.setEchoMode(QtGui.QLineEdit.Password)
        # Luodaan login-nappi
        self.buttonLogin = QtGui.QPushButton('Login', self)
        self.buttonLogin.clicked.connect(self.email_connect)

        # #### LABELS ######
        spt = "Sähköposti (GMail)"
        sat = "Salasana"

        passlabel = QtGui.QLabel(self)
        passlabel.setText(sat.decode('utf-8'))
        acclabel = QtGui.QLabel(self)
        acclabel.setText(spt.decode('utf-8'))
        # ##################

        # Lisätään luomamme objektit login-ikkunaan
        layout = QtGui.QVBoxLayout(self)
        layout.addWidget(self.info)
        layout.addWidget(acclabel)
        layout.addWidget(self.textName)
        layout.addWidget(passlabel)
        layout.addWidget(self.textPass)
        layout.addWidget(self.buttonLogin)

    # /__init__

    # Sähköpostitiliin yhdistys

    def email_connect(self):
        gP = self.textPass.text()
        gA = self.textName.text()
        # säädetään tekstikentät tukemaan ääkkösiä
        gP = unicode(gP)
        gA = unicode(gA)
        gmailPass = gP.encode('ISO-8859-1')
        Login.gmailAcc = gA.encode('ISO-8859-1')
        # Yritetään yhdistää palvelimeen ja kirjautua käyttäjän antamilla tunnuksilla
        try:
            self.smtpObj.ehlo()
            self.smtpObj.login(Login.gmailAcc, gmailPass)
        except smtplib.SMTPAuthenticationError:
            # Pop-up mikäli tunnistetiedoissa on virhe
            QtGui.QMessageBox.question(self, 'Error', "Yhteys ei onnistunut. Tarkista tunnistetiedot.", QtGui.QMessageBox.Ok)
        else:
            QtGui.QMessageBox.question(self, 'Yhteys muodostettu', "Tervetuloa!", QtGui.QMessageBox.Ok)
            # Annetaan koppi pääikkunalle
            self.accept()

    # /email_connect

# /Login

# Pääikkuna
# noinspection PyUnresolvedReferences
class Window(QtGui.QMainWindow):
    # __init__ suoritetaan aina kun uusi ikkuna avautuu
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)
        # Asetetaan ikkunan sijanti(50,50) ja koko(600, 500)
        self.setGeometry(50,50, 600, 500)
        # Ikkunan title ja kuvake
        self.setWindowTitle("Rapuposti")
        self.setWindowIcon(QtGui.QIcon('blizzard.PNG'))

        # ----- Solukentät. 2 on loppuva solu ------ #

        # Sähköposti
        # Alkava solu
        self.emailCell = QtGui.QLineEdit("", self)
        self.emailCell.setGeometry(15, 50, 30, 30)
        self.emailCellnr = QtGui.QSpinBox(self)
        self.emailCellnr.setGeometry(45, 50, 40, 30)
        self.emailCellnr.setMaximum(10000)
        # Loppuva solu
        self.emailCell2 = QtGui.QLineEdit("", self)
        self.emailCell2.setGeometry(90, 50, 30, 30)
        self.emailCellnr2 = QtGui.QSpinBox(self)
        self.emailCellnr2.setGeometry(120, 50, 40, 30)
        self.emailCellnr2.setMaximum(10000)

        # Nimi
        # Alkava solu
        self.nameCell = QtGui.QLineEdit("", self)
        self.nameCell.setGeometry(15, 110, 30, 30)
        self.nameCellnr = QtGui.QSpinBox(self)
        self.nameCellnr.setGeometry(45, 110, 40, 30)
        self.nameCellnr.setMaximum(10000)
        # Loppuva solu
        self.nameCell2 = QtGui.QLineEdit("", self)
        self.nameCell2.setGeometry(90, 110, 30, 30)
        self.nameCellnr2 = QtGui.QSpinBox(self)
        self.nameCellnr2.setGeometry(120, 110, 40, 30)
        self.nameCellnr2.setMaximum(10000)

        # Validaatio
        # Alkava solu
        self.validCell = QtGui.QLineEdit("", self)
        self.validCell.setGeometry(15, 170, 30, 30)
        self.validCellnr = QtGui.QSpinBox(self)
        self.validCellnr.setGeometry(45, 170, 40, 30)
        self.validCellnr.setMaximum(10000)
        # Loppuva solu
        self.validCell2 = QtGui.QLineEdit("", self)
        self.validCell2.setGeometry(90, 170, 30, 30)
        self.validCellnr2 = QtGui.QSpinBox(self)
        self.validCellnr2.setGeometry(120, 170, 40, 30)
        self.validCellnr2.setMaximum(10000)

        # ------- SOLUT END ----------- #

        # Viesti- & aihekenttä
        self.subjectBox = QtGui.QLineEdit("Aihe", self)
        self.subjectBox.setGeometry(QtCore.QRect(170, 50, 300, 30))
        self.messageBox = QtGui.QTextEdit("Viesti", self)
        self.messageBox.setGeometry(QtCore.QRect(170, 100, 400, 300))

        # Tiedostojen selaus
        fileBroswer = QtGui.QAction("&Avaa Excel-tiedosto", self)
        fileBroswer.setStatusTip('Open File')
        fileBroswer.triggered.connect(self.file_open)

        # QUIT TOIMINTO
        extractAction = QtGui.QAction("&Quit", self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip('Exit Application')
        extractAction.triggered.connect(self.close_app)

        # # menubaari# #
        self.statusBar()
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(extractAction)
        fileMenu.addAction(fileBroswer)
        # ########### #

        # Logo
        self.pic = QtGui.QLabel(self)
        self.pic.setGeometry(50, 230, 400, 100)
        self.pic.setPixmap(QtGui.QPixmap(os.getcwd() + "/rapupostilogo1.png"))

        self.home()
    # /__init__

# ################ HOME ###################

    def home(self):
        # #### NAPIT #######
        ebtn = QtGui.QPushButton("EXIT", self)
        ebtn.clicked.connect(self.close_app)
        ebtn.move(10, 450)
        self.gobtn = QtGui.QPushButton("GO", self)
        self.gobtn.move(490, 450)
        self.gobtn.clicked.connect(self.doEverything)
        self.progress = QtGui.QProgressBar(self)
        self.progress.setGeometry(200, 450, 250, 20)
        # ###################

        # #### LABELS ######
        ec = "Sähköpostisolut"
        nc = "Nimisolut"
        vc = "Validaatiosolut"

        elabel = QtGui.QLabel(self)
        elabel.setText(ec.decode('utf-8'))
        elabel.move(20, 27)
        elabel = QtGui.QLabel(self)
        elabel.setText(nc.decode('utf-8'))
        elabel.move(20, 80)
        elabel = QtGui.QLabel(self)
        elabel.setText(vc.decode('utf-8'))
        elabel.move(20, 140)
        # ##################

        self.show()

    # /home

# ########################################################

    def file_open(self):
        # Avaa filebrowserin Windowsissa ja tallentaa tiedostopolun "exel" -muuttujaan
        self.exel = QtGui.QFileDialog.getOpenFileName(self, 'Open File')
        # Muutetaan tiedostopolku absoluuttisesksi esim. 'C:/documents/tiedosto.xlsx  -> C:\documents\tiedosto.xlsx'
        self.pathTofile = os.path.abspath(self.exel)
    # /file_open

# ########## OHJELMAN SYDÄN (TÄHÄN TULEE TAULUKON TARKISTELU SEKÄ S-POSTIN LÄHETYS)#################
    def doEverything(self):

        # Tarkistetaan onko excel-tiedosto valittu
        try:
            self.exel
        except AttributeError:
            # Pop-up mikäli tiedostoa ei ole valittu
            QtGui.QMessageBox.question(self, 'Error', "Excel-tiedostoa ei ole valittu. Valitse tiedosto file-valikosta.",
                                                QtGui.QMessageBox.Ok)

        # Haetaan käyttäjän syöttämät arvot tekstikentistä ja käytetään ne str-luokan läpi,
        # jotta openpyxl pystyy lukemaan niitä
        # Spinboxien arvot saadaan poimittua valueFromText-moduulilla integer -muodossa
        # s1 = sähköposti, s2 = nimi, s3 = validation, a = alkava solu, b = loppuva solu, NR = rivinumero
        # Sähköposti
        s1a = self.emailCell.text()
        s1a = str(s1a)
        s1b = self.emailCell.text()
        s1b = str(s1b)
        NR1a = self.emailCellnr.text()
        NR1aint = self.emailCellnr.valueFromText(NR1a)
       # NR1astr = str(NR1a) <-- potentiaalinen muuttuja lisäominaisuuksille
        NR1b = self.emailCellnr2.text()
        NR1bint = self.emailCellnr2.valueFromText(NR1b)
        NR1bstr = str(NR1bint)

        # Nimi
        s2a = self.nameCell.text()
        s2a = str(s2a)
        s2b = self.nameCell.text()
        s2b = str(s2b)
        NR2a = self.nameCellnr.text()
        NR2aint = self.nameCellnr.valueFromText(NR2a)
        # NR2astr = str(NR2a) <-- potentiaalinen muuttuja lisäominaisuuksille
        NR2b = self.nameCellnr2.text()
        NR2bint = self.nameCellnr2.valueFromText(NR2b)
        NR2bstr = str(NR2bint)

        # Validaatio
        s3a = self.validCell.text()
        s3a = str(s3a)
        s3b = self.validCell2.text()
        s3b = str(s3b)
        NR3a = self.validCellnr.text()
        NR3aint = self.validCellnr.valueFromText(NR3a)
        # NR3astr = str(NR3aint) <-- potentiaalinen muuttuja lisäominaisuuksille
        NR3b = self.validCellnr2.text()
        NR3bint = self.validCellnr2.valueFromText(NR3b)
        NR3bstr = str(NR3bint)

        # Haetaan käyttäjän antama absoluuttinen tiedostopolku ja avataan se openpyxl-moduulilla
        wb = load_workbook(self.pathTofile)
        sheet = wb.get_sheet_by_name('Sheet1')

        # Loppuvat solut
        # Korjaa sanity check nr. 6
        try:
            endCellva = sheet[s3b + NR3bstr].value
            endCellna = sheet[s2b + NR2bstr].value
            endCellem = sheet[s1b + NR1bstr].value
        except IndexError:
            QtGui.QMessageBox.critical(self, "Error", "Yksi tai useampi loppuva solu ei ole valittu. Valitse solut ja koita uudelleen.")

        sub = self.subjectBox.text()
        mes = self.messageBox.toPlainText()
        subUn = unicode(sub)
        mesUn = unicode(mes)
        subEnc = subUn.encode('ISO-8859-1')
        mesEnc = mesUn.encode('ISO-8859-1')

        # Progress bar näyttää että lähetys on käynnissä
        self.completed = 0

        # muuttuja incrementaatiolle
        i = 0

        for everyRow in range(NR1aint - 1, NR1bint, 1):
            # alustetaan alkavat solut openpyxliä varten
            emailCellnr = NR1aint + i
            nameCellnr = NR2aint + i
            validCellnr = NR3aint + i
            # incrementaatio
            i += 1
            emailCellnrStr = str(emailCellnr)
            nameCellnrStr = str(nameCellnr)
            validCellnrStr = str(validCellnr)

            # Korjaa sanity check nr. 6
            try:
                emCell = sheet[s1a + emailCellnrStr].value
                naCell = sheet[s2a + nameCellnrStr].value
                vaCell = sheet[s3a + validCellnrStr].value
            except IndexError:
                QtGui.QMessageBox.critical(self, "Error", "Yksi tai useampi alkava solu ei ole valittu. Valitse solut ja koita uudelleen.")

            # Jos solu ei palauttanut arvoa tehdään muutujista tyhjiä stringejä
            if not emCell and not vaCell:
                emCell = ""
                vaCell = ""
                naCell = ""

            # Mikäli sähköpostiriviltä löytyy @-merkki eikä käyttäjälle ole vielä lähetetty spostia
            if '@' in emCell and vaCell != 'e-mail sent':
                Login.smtpObj.sendmail(Login.gmailAcc, emCell, 'Subject:' + subEnc + '\n' + mesEnc)
                sheet[s3a + validCellnrStr] = 'e-mail sent'

                # Progress bar käynnistyy
                while self.completed < 28:
                    self.completed += 0.0001
                    self.progress.setValue(self.completed)

                if emCell == endCellem and naCell == endCellna and vaCell == endCellva:
                    # Luodaan uusi excel-tiedosto, jossa validaatiokentät ovat muuttuneet osoittamaan kenelle posti on lähetetty
                    dest_filename = 'rapuposti_sheet.xlsx'
                    wb.save(filename=dest_filename)

        # Progress bar menee loppuun
        while self.completed < 100:
            self.completed += 0.0001
            self.progress.setValue(self.completed)

    # /do_everything

# #############################################

    # Exit toiminto. Kysyy käyttäjältä onko tämä varma, jospa kyseessä onkin misclick.
    def close_app(self):
            choice = QtGui.QMessageBox.question(self, 'Rapuposti', "Oletko Varma?", QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
            if choice == QtGui.QMessageBox.Yes:
                Login.smtpObj.quit()
                sys.exit()
            else:
                pass
    # /close_app

# /Window

# Luodaan Käyttöliittymä
def mainApp():

    login = Login()

    # Avataan pääikkuna
    if login.exec_() == QtGui.QDialog.Accepted:
        GUI = Window()
        sys.exit(app.exec_())

# /mainApp

# Avataan käyttöliittymä
mainApp()
