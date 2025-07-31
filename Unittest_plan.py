import unittest
import HTMLTestRunner # type: ignore
import os
from datetime import datetime

from FactoryReset import FactoryReset
from LiveView import Liveview

if __name__ == '__main__':
    # 確保測試報告目錄存在
    report_dir = 'D:/SeleniumProject/test_reports'
    os.makedirs(report_dir, exist_ok=True)

    print(f"\n--- Running tests for FactoryReset.py ---")
    suite = unittest.TestSuite()
    class_name = 'FactoryReset'
    selected_cases_list = ['test_case01_Check_Brightness']
    cases_description = '包含測試案例: ' + ', '.join(selected_cases_list)
    suite.addTest(FactoryReset('test_case01_Check_Brightness'))
    os.makedirs('D:/SeleniumProject/test_reports', exist_ok=True)
    existing_files = set(os.listdir('D:/SeleniumProject/test_reports'))
    runner = HTMLTestRunner.HTMLTestRunner(
        output='D:/SeleniumProject/test_reports',
        title=f'Test Report for {class_name} (FactoryReset)',
        description=f'Test results for {class_name} class in FactoryReset module. {cases_description}'
    )
    runner.run(suite)
    new_files = set(os.listdir('D:/SeleniumProject/test_reports')) - existing_files
    new_html_report_path = None
    for f_name in new_files:
        if f_name.endswith('.html'):
            new_html_report_path = os.path.join('D:/SeleniumProject/test_reports', f_name)
            break
    
    if new_html_report_path:
        current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
        desired_report_name = f'{class_name}_{current_date}.html'
        desired_report_full_path = os.path.join('D:/SeleniumProject/test_reports', desired_report_name)
        try:
            os.rename(new_html_report_path, desired_report_full_path)
            print(f"Test report for {class_name} saved and renamed to {desired_report_full_path}")
        except Exception as e:
            print(f"Error renaming report for {class_name}: {e}")
            print(f"Original report path: {new_html_report_path}")
    else:
        print(f"Could not find newly generated HTML report for {class_name} in D:/SeleniumProject/test_reports")

    print(f"\n--- Running tests for LiveView.py ---")
    suite = unittest.TestSuite()
    class_name = 'Liveview'
    selected_cases_list = ['test_case01_WelcomePage']
    cases_description = '包含測試案例: ' + ', '.join(selected_cases_list)
    suite.addTest(Liveview('test_case01_WelcomePage'))
    os.makedirs('D:/SeleniumProject/test_reports', exist_ok=True)
    existing_files = set(os.listdir('D:/SeleniumProject/test_reports'))
    runner = HTMLTestRunner.HTMLTestRunner(
        output='D:/SeleniumProject/test_reports',
        title=f'Test Report for {class_name} (LiveView)',
        description=f'Test results for {class_name} class in LiveView module. {cases_description}'
    )
    runner.run(suite)
    new_files = set(os.listdir('D:/SeleniumProject/test_reports')) - existing_files
    new_html_report_path = None
    for f_name in new_files:
        if f_name.endswith('.html'):
            new_html_report_path = os.path.join('D:/SeleniumProject/test_reports', f_name)
            break
    
    if new_html_report_path:
        current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
        desired_report_name = f'{class_name}_{current_date}.html'
        desired_report_full_path = os.path.join('D:/SeleniumProject/test_reports', desired_report_name)
        try:
            os.rename(new_html_report_path, desired_report_full_path)
            print(f"Test report for {class_name} saved and renamed to {desired_report_full_path}")
        except Exception as e:
            print(f"Error renaming report for {class_name}: {e}")
            print(f"Original report path: {new_html_report_path}")
    else:
        print(f"Could not find newly generated HTML report for {class_name} in D:/SeleniumProject/test_reports")

    print("\n--- All selected test suites have been executed and reported. ---")
