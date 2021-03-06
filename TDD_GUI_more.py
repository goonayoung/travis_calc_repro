import sys
import gui_demo
from unittest import TestCase, main

from PyQt5.QtWidgets import (QApplication, QTableWidgetItem)
app= QApplication( sys.argv )

class TestCase3(TestCase):
    def setup(self):
        pass
    
    def test_Save_clicked(self):
        a = gui_demo.MyApp()
        
        inputTable1 = 'Input Source'
        inputTable2 = 'Signal Description'
        inputTable3 = 'Type'
        inputTable4 = '1234'
        
        a.table.setItem(0, 0, QTableWidgetItem(inputTable1))
        a.table.setItem(0, 1, QTableWidgetItem(inputTable2))
        a.table.setItem(0, 2, QTableWidgetItem(inputTable3))
        a.table.setItem(0, 3, QTableWidgetItem(inputTable4))
        
        a.setBtn.click()
        a.searchBtn.click()
        
        a.moreBtn.click()
        self.assertEqual(a.outputTable.rowCount(), 2)
        
        a.moreBtn.click()
        self.assertEqual(a.outputTable.rowCount(), 3)

        
if __name__ == "__main__":
    main()   