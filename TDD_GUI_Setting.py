import sys
import SeeReal_KAI_
import unittest

from PyQt5.QtWidgets import QApplication
app= QApplication( sys.argv )

class TestCase1(unittest.TestCase):
    def setup(self):
        pass
    
    def test_Set_clicked(self):
        a = SeeReal_KAI_.MyApp()
        
        SearchR1 = '3.2.12'
        SearchR2 = '3.2.13'
        qle3 = '2'
        
        a.groupbox1.setChecked(True)
        a.groupbox2.setChecked(True)
        
        a.cb1.setCurrentText(SearchR1)
        a.cb2.setCurrentText(SearchR2)
        a.qle3.setText(qle3)
        
        a.setBtn.click()
         
        self.assertEqual(a.cb1.currentText()[0] +'.' + a.cb1.currentText()[2] + '.' + a.cb1.currentText()[4:], SearchR1)
        self.assertEqual(a.cb2.currentText()[0] +'.' + a.cb2.currentText()[2] + '.' + a.cb2.currentText()[4:], SearchR2)
        self.assertEqual(a.outputnum, int(qle3))
        

        
if __name__ == "__main__":
    unittest.main()   