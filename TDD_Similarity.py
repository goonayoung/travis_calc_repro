from unittest import TestCase, main
import SeeReal_KAI_



class TestCase1(TestCase):
    def setup():
        pass
    
    def test_checkSimilarity(self):
        b=SeeReal_KAI_.Similarity().listToVector(['data is goo','none was bad'])
        c=SeeReal_KAI_.Similarity().vectorToDense(b)
        d=SeeReal_KAI_.Similarity().checkSimilarity(c)
        
        self.assertEqual(d,0.0)
        
        
if __name__ == '__main__':
    main()