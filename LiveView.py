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

class Liveview(unittest.TestCase):

    @classmethod
    def setUpClass(cls):

        chrome_options = Options()
        chrome_options.add_argument('--log-level=3')  #  SSL訊息和警告都不顯示

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
        
        '''
        # 使用該設定開啟chrome
        service = Service(ChromeDriverManager().install())
        cls.driver = webdriver.Chrome(service=service, options=chrome_options)
        cls.driver.implicitly_wait(10)
        cls.driver.maximize_window()
        '''
        cls.driver = webdriver.Chrome()
        cls.driver.implicitly_wait(10)
        cls.driver.maximize_window()
        

        # 讀取配置文件
        config = configparser.ConfigParser()
        config.read(r'D:/selenium project/config.ini')
        URL = config['URL_Config']['URL']

        #開啟特定網址網頁
        cls.driver.get(URL)

        # 通過id查找元素:詳細資訊
        advance_button = cls.driver.find_element(By.ID, "details-button")
        advance_button.click()
        # 通過id查找元素:連結
        link_button = cls.driver.find_element(By.ID, "proceed-link")
        link_button.click()    
        
    def setUp(self):

        time.sleep(2)

    #Case 1:登入檢查
    def test_case01_WelcomePage(self):

        #確認login頁面welcome文字是否出現，若沒出現則判定fail
        try:
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "div_WelcomeMain"))
        )
        except TimeoutException:
            self.fail("cannot find welcome page")
        
        # 點擊Sign In Button
        SignIn_button = self.driver.find_element(By.ID, "button_Main_SignIn")
        SignIn_button.click()
    
    #Case 2:username欄位檢查
    def test_case02_username(self):

        #確認login頁面的username欄位是否出現，若沒出現則判定fail
        try:
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "div_SignIn_Username"))
        )
        except TimeoutException:
            self.fail("cannot find username column")

    #Case 3:password欄位檢查
    def test_case03_password(self):
        #確認login頁面的password欄位是否出現，若沒出現則判定fail
        try:
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "div_SignIn_Password"))
        )
        except TimeoutException:
            self.fail("cannot find password column")
    
    '''
    #Case 4:使用錯誤的帳密登入超過3次
    def test_case04_login_failed(self):
        Username_button = self.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys("admin1")
        Password_button = self.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys("Ability.2")
        LoginIn_button = self.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()
        time.sleep(2)
        LoginIn_button.click()
        time.sleep(2)
        LoginIn_button.click()
        time.sleep(2)
        #確認時間倒數是否出現，若沒出現則判定fail
        try:
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "span_CountDownTime"))
        )
        except TimeoutException:
            self.fail("CountDownTime was not triggered")
  
    #Case 5:使用預設帳密並登入後，並修改帳密
    def test_case05_login_default(self):
        Username_button = self.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys("admin")
        Password_button = self.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys("AbAi.Camera1")
        LoginIn_button = self.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()
        time.sleep(2)
        Username_button = self.driver.find_element(By.ID, "div_NewUser_Username")
        Username_button.send_keys("admin1")
        Password_button = self.driver.find_element(By.ID, "div_NewUser_Password")
        Password_button.send_keys("Ability.1")
        Confirm_Password_button = self.driver.find_element(By.ID, "div_NewUser_ConfirmPassword")
        Confirm_Password_button.send_keys("Ability.1")
        Ok_button = self.driver.find_element(By.ID, "button_NewUser_SignIn")
        Ok_button.click()
    
    #Case 6:使用帳密並登入後登出
    def test_case06_login(self):
        Username_button = self.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys("admin1")
        Password_button = self.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys("Ability.1")
        LoginIn_button = self.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()
        time.sleep(3)
        try:
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "a_Logout"))
            )
            Logout_button = self.driver.find_element(By.ID, "a_Logout")
            Logout_button.click()
            WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "button_OK"))
            )
            button_OK = self.driver.find_element(By.ID, "button_OK")
            button_OK.click()
            time.sleep(3)
        except TimeoutException:
            self.fail("登入失敗或者無法正常登出")
        
        # 點擊Sign In Button
        SignIn_button = self.driver.find_element(By.ID, "button_Main_SignIn")
        SignIn_button.click()
        '''
    
    #播放Live View後截圖，判定Live View正不正常
    def test_case07_LiveView(self):

        # 讀取配置文件
        config = configparser.ConfigParser()
        config.read(r'D:/selenium project/config.ini')
        username = config['Login_Config']['username']
        password = config['Login_Config']['password']

        # 輸入帳號密碼
        Username_button = self.driver.find_element(By.ID, "div_SignIn_Username")
        Username_button.send_keys(username)
        Password_button = self.driver.find_element(By.ID, "div_SignIn_Password")
        Password_button.send_keys(password)
        LoginIn_button = self.driver.find_element(By.ID, "button_SignIn_OK")
        LoginIn_button.click()

        #等待網頁全部讀取完畢才能點擊播放按鈕，目前無法使用selenium的各種等待，原因待查
        time.sleep(10)

        # 使用顯示等待元素可見
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "button_play")),"Login failed"
            )
        #點擊播放按鈕
        play_button=self.driver.find_element(By.ID, "button_play")
        play_button.click()
        #等待下方工具列消失後再截圖，需大於3秒
        time.sleep(4)

        # 讀取瀏覽器DPI縮放比例
        device_pixel_ratio = self.driver.execute_script("return window.devicePixelRatio")

        #定位截圖目標
        target_element = self.driver.find_element(By.ID,"canvas")

        # 擷取目標的位置和尺寸
        target_location = target_element.location
        target_size = target_element.size

        # 擷取整個螢幕截圖
        screenshot = self.driver.get_screenshot_as_png()
        screenshot = cv2.imdecode(np.frombuffer(screenshot, np.uint8), -1)

        # 裁剪出目標
        x = int(target_location['x']*device_pixel_ratio)
        y = int(target_location['y']*device_pixel_ratio)
        width = int(target_size['width']*device_pixel_ratio)
        height = int(target_size['height']*device_pixel_ratio)
        target_area = screenshot[y:y+height, x:x+width]

        # 保存螢幕截圖
        cv2.imwrite('D:/SeleniumProject/LiveView_screenshot.png', target_area)
        
        # 讀取圖片
        image = cv2.imread('D:/SeleniumProject/LiveView_screenshot.png', cv2.IMREAD_GRAYSCALE)

        # 檢查圖片是否全黑
        if (image == 0).all():
            print("video is black")
            self.fail("video is black")  # 如果圖片全黑，則測試案例失敗
        elif (image == 255).all():
            print("video is white")
            self.fail("video is white")  # 如果圖片全白，則測試案例失敗
        else:
            print("video is normal")

    #模擬滑鼠移動後點擊麥克風圖示
    def test_case08_Microphone(self):

            # 找到懸停的目標元素
            hover_element = self.driver.find_element(By.ID, "canvas")  
            # 執行懸停操作
            ActionChains(self.driver).move_to_element(hover_element).perform()
            time.sleep(1)
            # 點擊麥克風
            Microphone_button = self.driver.find_element(By.ID, "input_Mic")
            Microphone_button.click()
            # 確認右上角的麥克風圖示有出現，否則fail
            try:
                WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "LVSpeaker"))
            )
            except TimeoutException:
                self.fail("Can not enable microphone")
    
    
    #模擬滑鼠停在button上時，是否有正確顯示button tips
    def test_case09_ButtonTips(self):
            
            # 找到懸停的麥克風目標元素
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "span_Main_Microphone"))
            )
            hover_element = self.driver.find_element(By.ID, "span_Main_Microphone")
            #獲取懸停前的style屬性
            original_style = hover_element.get_attribute('style')
            print("Original Style:", original_style)
            # 執行懸停操作
            ActionChains(self.driver).move_to_element(hover_element).perform()
            #獲取懸停後的style屬性
            hover_style = hover_element.get_attribute('style')
            print("Hover Style:", hover_style)
            if original_style != hover_style:
                print("Microphone style changes are verified on hover.")
            else:
                print("Microphone style no changes in style on hover.")
                self.fail("No microphone button tips.")  

            #停3秒等待下方工具列消失
            time.sleep(3)
             # 找到懸停的Live View目標元素
            hover_element = self.driver.find_element(By.ID, "canvas")  
            # 執行懸停操作，確保下方工具列可再次顯示
            ActionChains(self.driver).move_to_element(hover_element).perform()
            time.sleep(1)
            #找到懸停的音量目標元素
            hover_element_volume_input = self.driver.find_element(By.ID, "input_Volume")
            hover_element_volume_input.click()
            #獲取懸停前的volume style屬性
            hover_element_volume = self.driver.find_element(By.ID, "span_Main_Volume")
            original_style_volume = hover_element_volume.get_attribute('style')
            print("Original Style:", original_style_volume)
            # 執行懸停操作
            ActionChains(self.driver).move_to_element(hover_element_volume).perform()
            #獲取懸停後的volume style屬性
            hover_style_volume = hover_element_volume.get_attribute('style')
            print("Hover Style:", hover_style_volume)
            if original_style_volume != hover_style_volume:
                print("Volume style changes are verified on hover.")
            else:
                print("Volume style no changes in style on hover.")
                self.fail("No volume button tips.")
            
            #停5秒等待下方工具列消失
            time.sleep(3)
            
    #點擊snapshot按鈕
    def test_case10_Snapshot(self):
        
        download_folder = "D:/downloads"
        files_before = set(os.listdir(download_folder))  #下載前的文件列表
        # 顯性等待Live View              
        WebDriverWait(self.driver, 10).until(
        EC.visibility_of_element_located((By.ID, "canvas")),"Cannot find live view"
            )
        # 找到懸停的Live View目標元素
        hover_element_live = self.driver.find_element(By.ID, "canvas")  
         # 執行懸停操作，確保下方工具列可再次顯示
        ActionChains(self.driver).move_to_element(hover_element_live).perform()

        # 點擊snapshot按鈕
        Microphone_button = self.driver.find_element(By.ID, "input_SnapShot")
        Microphone_button.click()
        time.sleep(2)
        # 點擊確認按鈕
        WebDriverWait(self.driver, 10).until(
        EC.visibility_of_element_located((By.ID, "button_BottomConfirmDialog")),"Cannot find confirm button"
            )
        ConfirmButton = self.driver.find_element(By.ID, "button_BottomConfirmDialog")
        ConfirmButton.click()
        time.sleep(5)

        #確認資料夾內是否有產生新檔案
        files_after = set(os.listdir(download_folder))  # 下載後的文件列表
        new_files = files_after - files_before  # 新下載的文件
        self.assertTrue(new_files, "No new file downloaded.")
        print("snapshot name is :", new_files)

    #點擊full screen按鈕
    def test_case11_Fullscreen(self):

         # 顯性等待Live View              
        WebDriverWait(self.driver, 10).until(
        EC.visibility_of_element_located((By.ID, "canvas")),"Cannot find live view"
            )
        # 找到懸停的Live View目標元素
        hover_element_live = self.driver.find_element(By.ID, "canvas")  
        # 執行懸停操作，確保下方工具列可再次顯示
        ActionChains(self.driver).move_to_element(hover_element_live).perform()
        #先找到Live View畫面的原始style屬性
        Original_LiveView = self.driver.find_element(By.ID, "canvas")
        Original_LiveView_style = Original_LiveView.get_attribute('style')
        print(Original_LiveView_style)
        # 點擊full screen
        Fullscreen_button = self.driver.find_element(By.ID, "input_FullScreen")
        Fullscreen_button.click()
        time.sleep(3)
        #找到Live View畫面點擊full screen後的style屬性
        Later_LiveView = self.driver.find_element(By.ID, "canvas")
        Later_LiveView_style = Later_LiveView.get_attribute('style')
        print(Later_LiveView_style)
        #比較點擊按鈕後的屬性
        if Original_LiveView_style != Later_LiveView_style:
                print("Live View style changes are verified on fullscreen button.")
        else:
                print("Live View style no changes in style on fullscreen button.")
                self.fail("Full Screen button does not work.")
        # 找到懸停的Live View目標元素
        hover_element_live = self.driver.find_element(By.ID, "canvas")  
        # 執行懸停操作，確保下方工具列可再次顯示
        ActionChains(self.driver).move_to_element(hover_element_live).perform()
        # 再次點擊full screen回到原始UI頁面
        Fullscreen_button = self.driver.find_element(By.ID, "input_FullScreen")
        Fullscreen_button.click()
        time.sleep(3)

    #點擊Zoom in/out按鈕
    def test_case12_Zoom(self):

        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "input_ZoomIn")),"cannot find Zoom button"
         )   
        # 點擊Zoom in
        ZoomIn_button = self.driver.find_element(By.ID, "input_ZoomIn")
        ZoomIn_button.click()
        time.sleep(1)

        # 等待確認按鈕存在並點擊
        try:
            confirm_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "button_BottomConfirmDialog"))
            )
            confirm_button.click()
        except:
            pass # 如果不存在就忽略

        # 讀取配置文件
        config = configparser.ConfigParser()
        config.read(r'D:/selenium project/config.ini')
        zoom_in = config['Zoom_Config']['zoom_in']

        # 取得目前的zoom倍率
        span_element = self.driver.find_element(By.ID,"span_ZoomValue")
        zoom_value = span_element.text
        print(zoom_value)
        #連續點擊Zoom in，依照機種不同，Zoom value的上限會不同
        while zoom_value < zoom_in:
            ZoomIn_button.click()
            span_element = self.driver.find_element(By.ID,"span_ZoomValue")
            zoom_value = span_element.text
            time.sleep(1)
        #連續點擊Zoom out
        ZoomOut_button = self.driver.find_element(By.ID, "input_ZoomOut")
        while zoom_value != "1.0x":
            ZoomOut_button.click()
            span_element = self.driver.find_element(By.ID,"span_ZoomValue")
            zoom_value = span_element.text
            time.sleep(1)
        time.sleep(2)
        
                       
                       
    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
       

if __name__ == '__main__':
    unittest.main(testRunner=HTMLTestRunner.HTMLTestRunner(output='D:/SeleniumProject/test_reports'))
