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


class FactoryReset(unittest.TestCase):

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
        
    def setUp(self):

        time.sleep(2)

    
    #Case 07:檢查色溫是否為5000k
    def test_case07_Check_Color_Temperature(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
        checkbox = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance .slider")
        if checkbox.is_selected():
            checkbox.click()  # 若不是OFF，點一下勾選它
        # 找到色溫滑桿
        color_temp_slider = self.driver.find_element(By.ID, "slider_colorTemperature")
        # 取得當前的 value 屬性
        current_temp = color_temp_slider.get_attribute("value")
        # 驗證是否為 5000K
        if current_temp == "5000":
            print("Color temperature is 5000K")
        else:
            self.fail(f"Color temperature is {current_temp}K, not 5000K")
        #等待兩秒後切換回為ON，不能馬上切換，否則會失敗
        time.sleep(2)
        checkbox1 = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance .slider")
        if not checkbox1.is_selected():
            checkbox1.click()  # 若不是 ON，點一下勾選它


            
       
    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
       

if __name__ == '__main__':
    unittest.main(testRunner=HTMLTestRunner.HTMLTestRunner(output='D:/SeleniumProject/test_reports'))
