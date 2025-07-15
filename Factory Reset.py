import unittest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager # type: ignore
from selenium.webdriver.common.action_chains import ActionChains
import time
import cv2
import numpy as np
import HTMLTestRunner # type: ignore
import os
import configparser

class MyTestCase(unittest.TestCase):

    @classmethod
    def setUpClass(cls):

        chrome_options = Options()
        chrome_options.add_argument('--log-level=3')  # SSL訊息和警告都不顯示

        # 在網站自動授予權限，可以用以下参数
        chrome_options.add_experimental_option(
        "prefs", {
            "profile.default_content_setting_values.media_stream_mic": 1,  # 允許麥克風
            "profile.default_content_setting_values.media_stream_camera": 1,  # 允許攝影機
            "profile.default_content_setting_values.geolocation": 1,  # 允許地理位置
            "profile.default_content_setting_values.notifications": 1,  # 允許通知
            "download.default_directory": "D:\\downloads"  # 更新為你的下載路徑
            }
        )
        
        # 使用該設定開啟chrome
        service = Service(ChromeDriverManager().install())
        cls.driver = webdriver.Chrome()
        cls.driver.implicitly_wait(10)
        cls.driver.maximize_window()

        # 讀取配置文件
        config = configparser.ConfigParser()
        config.read(r'D:/selenium project/config.ini')

        URL = config['URL_Config']['URL']

        #開啟特定網址網頁
        cls.driver.get(URL)
        
        # 讀取配置文件
        config = configparser.ConfigParser()
        config.read(r'D:/selenium project/config.ini')
        username = config['Login_Config']['username']
        password = config['Login_Config']['password']

        # 輸入帳號密碼
        Username_button = cls.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys(username)
        Password_button = cls.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys(password)
        LoginIn_button = cls.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()
        time.sleep(10)

        #點擊Administrator按鈕進入Management頁面
        Administrator_button = cls.driver.find_element(By.XPATH, "//span[@data-lang='span_A026']")
        Administrator_button.click()
        time.sleep(3)
        #點擊Factory按鈕進行重置
        Factory_Reset_button = cls.driver.find_element(By.XPATH, "//span[@data-lang='span_B400']")
        Factory_Reset_button.click()
        time.sleep(2)
        #對彈窗訊息點擊OK按鈕確認進行重置
        Factory_Reset_button = cls.driver.find_element(By.ID, "button_OK")
        Factory_Reset_button.click()
        time.sleep(120)
        #重新輸入帳號密碼
        Username_button = cls.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys(username)
        Password_button = cls.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys(password)
        LoginIn_button = cls.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()
        time.sleep(10)
           
        
    def setUp(self):

        time.sleep(2)

    #Case 01:檢查亮度是否為50%
    def test_case01_Check_Brightness(self):

        #定義圖片的基本儲存路徑
        base_path = 'D:/SeleniumProject/Resolution/'
        #如果路徑不存在，則創建資料夾
        if not os.path.exists(base_path):
            os.makedirs(base_path)
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Brightness元素
        slider_Brightness = self.driver.find_element(By.ID, "input_Brightness")
        slider_Brightness_style = slider_Brightness.get_attribute('value')
        print(slider_Brightness_style)
        if slider_Brightness_style=="50%":
            print("Default button works")
        else:
            self.fail("Default button does not work(Brightness)")


    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
       

if __name__ == '__main__':
    unittest.main(testRunner=HTMLTestRunner.HTMLTestRunner(output='D:/SeleniumProject/test_reports'))
