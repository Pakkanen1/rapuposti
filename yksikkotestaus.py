# coding=utf-8

import sys
import lobsteriposti2
from PyQt4.QtTest import QTest
from PyQt4.QtCore import Qt
from PyQt4 import QtGui
import unittest

app = QtGui.QApplication(sys.argv)
lol = lobsteriposti2.Login()
lwi = lobsteriposti2.Window()

class Lobsterunittest(unittest.TestCase):

    # Unit test 1, tarkistaa ovatko täytettävät tekstikentät tyhjiä ja toimiiko login-napin painallus
    def test_login(self):
        self.assertEqual(lol.textName.text(), "")
        self.assertEqual(lol.textPass.text(), "")
        logbutton = lol.buttonLogin
        QTest.mouseClick(logbutton, Qt.LeftButton)

    # /test_login

    # Unit test 2, testaa go-napin painalluksen kun tiedostoa ei ole valittu
    def test_go(self):
        gobutton = lwi.gobtn
        QTest.mouseClick(gobutton, Qt.LeftButton)

    # /test_go

    # Unit test 3, tarkistaa, kolumnikentät ovat tyhjiä oletuksena
    def test_cells(self):
        self.assertEqual(lwi.emailCell.text(), "")
        self.assertEqual(lwi.emailCell2.text(), "")
        self.assertEqual(lwi.nameCell.text(), "")
        self.assertEqual(lwi.nameCell2.text(), "")
        self.assertEqual(lwi.validCell.text(), "")
        self.assertEqual(lwi.validCell2.text(), "")

    # /test_cells

    # Unit test 4, testaa spinboxien nuolten toiminnan (solunumeron valinta)
    def test_cellspinbox(self):
        # Sähköpostisolut
        enr1 = lwi.emailCellnr
        enr2 = lwi.emailCellnr2
        QTest.mouseClick(enr1, Qt.LeftButton)
        QTest.mouseClick(enr2, Qt.LeftButton)
        # Nimisolut
        nnr1 = lwi.nameCellnr
        nnr2 = lwi.nameCellnr2
        QTest.mouseClick(nnr1, Qt.LeftButton)
        QTest.mouseClick(nnr2, Qt.LeftButton)
        # Validaatiosolut
        vnr1 = lwi.validCellnr
        vnr2 = lwi.validCellnr2
        QTest.mouseClick(vnr1, Qt.LeftButton)
        QTest.mouseClick(vnr2, Qt.LeftButton)

    # /test_cellspinbox

    # Unit test 5, testataan onko progressbar oletuksena yhtään täyttynyt
    def test_statusbar(self):
        self.assertEqual(lwi.progress.value(), -1)

    # /test-statusbar

unittest.main(exit=False)
