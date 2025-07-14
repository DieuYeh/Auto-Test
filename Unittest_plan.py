import unittest
class MathTest(unittest.TestCase):
    # 這個類別應該在你的測試文件中定義

if __name__ == '__main__':
    suite = unittest.TestSuite()
    suite.addTest(MathTest('setUpClass'))
    suite.addTest(MathTest('setUp'))
    suite.addTest(MathTest('test_case01_WelcomePage'))
    suite.addTest(MathTest('test_case02_username'))
    suite.addTest(MathTest('test_case03_password'))
    suite.addTest(MathTest('test_case04_login_failed'))
    suite.addTest(MathTest('test_case05_login_default'))
    suite.addTest(MathTest('test_case06_login'))
    suite.addTest(MathTest('test_case07_LiveView'))
    suite.addTest(MathTest('test_case08_Microphone'))
    suite.addTest(MathTest('test_case09_ButtonTips'))
    suite.addTest(MathTest('test_case10_Snapshot'))
    suite.addTest(MathTest('test_case11_Fullscreen'))
    suite.addTest(MathTest('test_case12_Zoom'))
    suite.addTest(MathTest('tearDownClass'))
    runner = unittest.TextTestRunner()
    runner.run(suite)
