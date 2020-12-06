from unittest import TestCase, main
import gui_demo



class TestCase1(TestCase):
    def setup():
        pass
    
    def test_checkSimilarity(self):
        b=gui_demo().listToVector(['data is goo','none was bad'])
        c=gui_demo().vectorToDense(b)
        d=gui_demo().checkSimilarity(c)
        
        self.assertEqual(d,0.0)
        
        
if __name__ == '__main__':
    main()