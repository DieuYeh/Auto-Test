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

        '''
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
        '''
        
    def setUp(self):

        time.sleep(2)

    #Case 01:檢查亮度是否為50%
    def test_case001_Check_Brightness(self):
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Brightness元素
        slider_Brightness = self.driver.find_element(By.ID, "input_Brightness")
        slider_Brightness_style = slider_Brightness.get_attribute('value')
        print(slider_Brightness_style)
        if slider_Brightness_style=="50%":
            print("factory button works, change brightness to 50%")
        else:
            self.fail("factory reset does not work(Brightness)")
    
    #Case 02:檢查contast是否為50%
    def test_case002_Check_Contrast(self):
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Contrast元素
        slider_Contrast = self.driver.find_element(By.ID, "input_Contrast")
        slider_Contrast_style = slider_Contrast.get_attribute('value')
        print(slider_Contrast_style)
        if slider_Contrast_style=="50%":
            print("factory button works, change contrast to 50%")
        else:
            self.fail("factory reset button does not work(Contrast)")
    
    #Case 03:檢查Saturation是否為50%
    def test_case003_Check_Saturation(self):
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Saturation元素
        slider_Saturation = self.driver.find_element(By.ID, "input_Saturation")
        slider_Saturation_style = slider_Saturation.get_attribute('value')
        print(slider_Saturation_style)
        if slider_Saturation_style=="50%":
            print("factory button works, change saturation to 50%")
        else:
            self.fail("factory reset button does not work(Saturation)")
    
    #Case 04:檢查Sharpness是否為50%
    def test_case004_Check_Sharpness(self):
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Sharpness元素
        slider_Sharpness = self.driver.find_element(By.ID, "input_Saturation")
        slider_Sharpness_style = slider_Sharpness.get_attribute('value')
        print(slider_Sharpness_style)
        if slider_Sharpness_style=="50%":
            print("factory button works, change sharpness to 50%")
        else:
            self.fail("factory reset button does not work(Sharpness)")
    
    #Case 05:檢查Gamma是否為50%
    def test_case005_Check_Gamma(self):
        
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)

        # 定位到slider_Gamma元素
        slider_Gamma = self.driver.find_element(By.ID, "input_Gamma")
        slider_Gamma_style = slider_Gamma.get_attribute('value')
        print(slider_Gamma_style)
        if slider_Gamma_style=="50%":
            print("factory button works, change Gamma to 50%")
        else:
            self.fail("factory reset button does not work1(Gamma1)")
    
    #Case 06:檢查auto_wb_mode是否為on
    def test_case006_Check_Auto_wb_Mode(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
        #判斷是否為on
        checkbox = self.driver.find_element(By.ID, "WhiteBalanceAuto")
        if checkbox.is_selected():
            print("WhiteBalanceAuto is ON")
        else:
            self.fail("WhiteBalanceAuto is OFF")
    
    #Case 07:檢查色溫是否為5000k
    def test_case007_Check_Color_Temperature(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
         #input的checkbox才是真正的狀態控制，slidery則是可點擊，因此判斷checkbox點擊slider
        checkbox = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance input[type='checkbox']")
        slider = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance .slider")
        if checkbox.is_selected():
            slider.click()  # 若不是OFF，點一下勾選它
        # 找到色溫滑桿
        color_temp_slider = self.driver.find_element(By.ID, "slider_colorTemperature")
        # 取得當前的 value 屬性
        current_temp = color_temp_slider.get_attribute("value")
        # 驗證是否為 5000K
        if current_temp == "5000":
            print("Color temperature is 5000K")
        else:
            self.fail(f"Color temperature is {current_temp}K, not 5000K")
        #等待三秒後切換回為ON，不能馬上切換，否則會失敗
        time.sleep(3)
        checkbox = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance input[type='checkbox']")
        slider = self.driver.find_element(By.CSS_SELECTOR, "#div_WhiteBalance .slider")
        if not checkbox.is_selected():
            slider.click()  # 若不是 ON，點一下勾選它

    #Case 08:檢查LDC是否為off
    def test_case008_Check_LDC(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
        #判斷是否為off
        checkbox = self.driver.find_element(By.ID, "LDC")
        if checkbox.is_selected():
            self.fail("LDC is ON")
        else:
            print("LDC is OFF")
    
    #Case 09:檢查Flip是否為off
    def test_case09_Check_RotateViewFlip(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
        #點擊Image config按鈕進入config頁面
        Image_button = self.driver.find_element(By.ID, "a_ImageConfigs")
        Image_button.click()
        time.sleep(2)
        flip_div = self.driver.find_element(By.ID, "select_ImagePara_Flip_div")
        status = flip_div.get_attribute("data-text")
        if status == "Off":
            print("選項目前是OFF")
        else:
            self.fail("flip is not off, it's " + status)

    #Case 10:檢查VideoOrientation是否為0°
    def test_case010_Check_VideoOrientation(self):
        #點擊Image按鈕進入image頁面
        Image_button = self.driver.find_element(By.ID, "a_Image")
        Image_button.click()
        time.sleep(2)
        #點擊Image config按鈕進入config頁面
        Image_button = self.driver.find_element(By.ID, "a_ImageConfigs")
        Image_button.click()
        time.sleep(2)
        flip_div = self.driver.find_element(By.ID, "select_ImagePara_VideoOrientation_div")
        status = flip_div.get_attribute("data-text")
        if status == "0°":
            print("the option is 0°")
        else:
            self.fail("Video Orientation is not 0°, it's " + status)


    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
       

if __name__ == '__main__':
    unittest.main(testRunner=HTMLTestRunner.HTMLTestRunner(output='D:/SeleniumProject/test_reports'))
