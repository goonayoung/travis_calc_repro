from unittest import TestCase, main
import gui_demo



class TestCase1(TestCase):
    def setup():
        pass
    
    def test_checkSimilarity(self):
        test = gui_demo()
        b=test.listToVector(['data is goo','none was bad'])
        c=test.vectorToDense(b)
        d=test.checkSimilarity(c)
        
        self.assertEqual(d,0.0)
        
        
if __name__ == '__main__':
    main()