import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PY 檔案與測試報告工具")
        self.geometry("900x700")

        self.py_file_path = ""
        self.selected_cases = {} # 儲存選中的測試案例 {case_name: Tkinter_BooleanVar}
        self.all_found_cases = [] # 儲存所有找到的測試案例名稱

        # 創建筆記本 (Tabbed Interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # PY 檔案處理 Tab
        self.py_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.py_tab, text="PY 檔案處理")
        self.create_py_tab_widgets(self.py_tab)

        # Excel 報告處理 Tab
        self.excel_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.excel_tab, text="Excel 報告處理")
        self.create_excel_tab_widgets(self.excel_tab)

    def create_py_tab_widgets(self, parent_frame):
        # 檔案選擇區
        file_frame = tk.LabelFrame(parent_frame, text="PY 檔案載入", padx=10, pady=10)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_label = tk.Label(file_frame, text="未選擇檔案")
        self.file_label.pack(side=tk.LEFT, expand=True, fill="x")

        select_file_btn = tk.Button(file_frame, text="載入 PY 檔案", command=self.load_py_file)
        select_file_btn.pack(side=tk.RIGHT)

        # 搜尋設定區
        search_frame = tk.LabelFrame(parent_frame, text="搜尋設定", padx=10, pady=10)
        search_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(search_frame, text="搜尋關鍵字 (預設: def):").pack(side=tk.LEFT)
        self.keyword_entry = tk.Entry(search_frame, width=20)
        self.keyword_entry.insert(0, "def") # 設定預設值
        self.keyword_entry.pack(side=tk.LEFT, padx=5)

        search_btn = tk.Button(search_frame, text="搜尋", command=self.search_py_file)
        search_btn.pack(side=tk.LEFT)

        # 結果顯示區 (Treeview)
        result_frame = tk.LabelFrame(parent_frame, text="搜尋結果", padx=10, pady=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(result_frame, columns=("No.", "Case name", "Select"), show="headings")
        self.tree.heading("No.", text="No.", anchor=tk.CENTER)
        self.tree.heading("Case name", text="Case Name", anchor=tk.W)
        self.tree.heading("Select", text="選擇", anchor=tk.CENTER)

        self.tree.column("No.", width=50, anchor=tk.CENTER)
        self.tree.column("Case name", width=300, anchor=tk.W)
        self.tree.column("Select", width=70, anchor=tk.CENTER)

        self.tree.pack(side=tk.LEFT, fill="both", expand=True)

        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.tree.config(yscrollcommand=scrollbar.set)

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click) # 綁定點擊事件

        # 選項按鈕區
        option_frame = tk.Frame(parent_frame, padx=10, pady=5)
        option_frame.pack(fill="x", padx=10, pady=5)

        select_all_btn = tk.Button(option_frame, text="全選", command=self.select_all_cases)
        select_all_btn.pack(side=tk.LEFT, padx=5)

        deselect_all_btn = tk.Button(option_frame, text="全取消", command=self.deselect_all_cases)
        deselect_all_btn.pack(side=tk.LEFT, padx=5)

        self.selected_count_label = tk.Label(option_frame, text="已勾選: 0")
        self.selected_count_label.pack(side=tk.LEFT, padx=10)

        export_btn = tk.Button(option_frame, text="匯出 Unittest Plan", command=self.export_unittest_plan)
        export_btn.pack(side=tk.RIGHT)

    def create_excel_tab_widgets(self, parent_frame):
        # Excel 設定區
        excel_settings_frame = tk.LabelFrame(parent_frame, text="Excel 設定", padx=10, pady=10)
        excel_settings_frame.pack(fill="x", padx=10, pady=5)

        # 讀取行數
        tk.Label(excel_settings_frame, text="讀取行數 (列A, 行B):").grid(row=0, column=0, sticky="w", pady=2)
        self.read_col_entry = tk.Entry(excel_settings_frame, width=5)
        self.read_col_entry.grid(row=0, column=1, padx=2, pady=2)
        self.read_row_entry = tk.Entry(excel_settings_frame, width=5)
        self.read_row_entry.grid(row=0, column=2, padx=2, pady=2)
        self.read_col_entry.insert(0, "A") # 預設 A 列
        self.read_row_entry.insert(0, "1") # 預設 1 行

        # 寫入行數
        tk.Label(excel_settings_frame, text="寫入行數 (列A, 行B):").grid(row=1, column=0, sticky="w", pady=2)
        self.write_col_entry = tk.Entry(excel_settings_frame, width=5)
        self.write_col_entry.grid(row=1, column=1, padx=2, pady=2)
        self.write_row_entry = tk.Entry(excel_settings_frame, width=5)
        self.write_row_entry.grid(row=1, column=2, padx=2, pady=2)
        self.write_col_entry.insert(0, "B") # 預設 B 列
        self.write_row_entry.insert(0, "1") # 預設 1 行

        # Excel 操作按鈕
        excel_buttons_frame = tk.Frame(parent_frame, padx=10, pady=5)
        excel_buttons_frame.pack(fill="x", padx=10, pady=5)

        load_testplan_btn = tk.Button(excel_buttons_frame, text="載入 Testplan (Excel)", command=self.load_testplan)
        load_testplan_btn.pack(side=tk.LEFT, padx=5)

        write_results_btn = tk.Button(excel_buttons_frame, text="寫入結果 (HTML -> Excel)", command=self.write_results_to_excel)
        write_results_btn.pack(side=tk.LEFT, padx=5)

    # --- PY 檔案處理相關方法 ---
    def load_py_file(self):
        """載入 PY 檔案並更新顯示路徑"""
        file_path = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
        if file_path:
            self.py_file_path = file_path
            self.file_label.config(text=f"已選擇: {os.path.basename(file_path)}")
            messagebox.showinfo("成功", "PY 檔案載入成功！")
            self.search_py_file() # 載入後自動搜尋

    def search_py_file(self):
        """讀取 PY 檔案內容，搜尋關鍵字並更新 Treeview"""
        if not self.py_file_path:
            messagebox.showwarning("警告", "請先載入 PY 檔案！")
            return

        keyword = self.keyword_entry.get().strip()
        if not keyword:
            messagebox.showwarning("警告", "請輸入搜尋關鍵字！")
            return

        self.all_found_cases = []
        self.selected_cases.clear()
        self.tree.delete(*self.tree.get_children()) # 清空 Treeview

        try:
            with open(self.py_file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 使用正則表達式尋找 def XXXX(cls) 模式並擷取 XXXX
            # 這個正則表達式會匹配 'def ' 後面跟著一個或多個非空白字元，直到遇到 '('
            # 並確保後面可能有 (cls) 或其他參數
            pattern = re.compile(rf'{re.escape(keyword)}\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\([^)]*\):')
            matches = pattern.finditer(content)

            no = 1
            for match in matches:
                case_name = match.group(1) # 擷取 XXXX
                # 檢查是否已存在，避免重複加入 (如果檔案中有重複定義)
                if case_name not in [item[1] for item in self.all_found_cases]:
                    self.all_found_cases.append(case_name)
                    var = tk.BooleanVar(value=False) # 預設為不選中
                    self.selected_cases[case_name] = var
                    self.tree.insert("", "end", iid=case_name, values=(no, case_name, "☐"), tags=("checkbox",))
                    no += 1
            self.update_selected_count()
            messagebox.showinfo("搜尋完成", f"找到 {len(self.all_found_cases)} 個符合 '{keyword}' 的項目。")

        except Exception as e:
            messagebox.showerror("錯誤", f"讀取或搜尋檔案時發生錯誤: {e}")

    def on_tree_click(self, event):
        """處理 Treeview 點擊事件，切換選擇狀態"""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        column = self.tree.identify_column(event.x)
        if column == "#3": # 選擇欄位
            case_name = self.tree.item(item_id, "values")[1] # 取得 Case name
            if case_name in self.selected_cases:
                current_var = self.selected_cases[case_name]
                current_var.set(not current_var.get()) # 切換布林值
                self.update_checkbox_display(item_id, current_var.get())
                self.update_selected_count()

    def update_checkbox_display(self, item_id, is_selected):
        """更新 Treeview 中的複選框顯示"""
        current_values = list(self.tree.item(item_id, "values"))
        current_values[2] = "☑" if is_selected else "☐"
        self.tree.item(item_id, values=current_values)

    def update_selected_count(self):
        """更新已勾選項目數量顯示"""
        count = sum(1 for var in self.selected_cases.values() if var.get())
        self.selected_count_label.config(text=f"已勾選: {count}")

    def select_all_cases(self):
        """全選所有搜尋到的測試案例"""
        for case_name, var in self.selected_cases.items():
            var.set(True)
            self.update_checkbox_display(case_name, True)
        self.update_selected_count()

    def deselect_all_cases(self):
        """全取消選所有搜尋到的測試案例"""
        for case_name, var in self.selected_cases.items():
            var.set(False)
            self.update_checkbox_display(case_name, False)
        self.update_selected_count()

    def export_unittest_plan(self):
        """匯出 Unittest Plan 到使用者指定的檔案"""
        selected_cases_list = [
            name for name, var in self.selected_cases.items() if var.get()
        ]

        if not selected_cases_list:
            messagebox.showwarning("警告", "請至少選擇一個測試案例！")
            return

        # 彈出檔案儲存對話框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
            initialfile="Unittest_plan.py" # 預設檔名
        )

        if not file_path: # 如果使用者取消了儲存對話框
            return

        output_content = [
            "import unittest\n",
            "class MathTest(unittest.TestCase):\n", # 假設這裡有一個 MathTest 類別，你可能需要根據實際情況修改
            "    # 這個類別應該在你的測試文件中定義\n\n",
            "if __name__ == '__main__':\n",
            "    suite = unittest.TestSuite()\n"
        ]

        for case_name in selected_cases_list:
            # 這裡假設你的測試案例都是 MathTest 類別的方法
            # 如果你的實際情況不同，請修改 'MathTest'
            output_content.append(f"    suite.addTest(MathTest('{case_name}'))\n")

        output_content.extend([
            "    runner = unittest.TextTestRunner()\n",
            "    runner.run(suite)\n"
        ])

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.writelines(output_content)
            messagebox.showinfo("匯出成功", f"Unittest Plan 已成功匯出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"匯出檔案時發生錯誤: {e}")

    # --- Excel 報告處理相關方法 ---
    def load_testplan(self):
        """載入 Testplan (Excel)，並儲存到 Result 資料夾"""
        excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if excel_file_path:
            result_dir = "Result"
            os.makedirs(result_dir, exist_ok=True) # 確保 Result 資料夾存在

            dest_path = os.path.join(result_dir, os.path.basename(excel_file_path))
            try:
                shutil.copy(excel_file_path, dest_path)
                messagebox.showinfo("成功", f"Testplan '{os.path.basename(excel_file_path)}' 已成功載入並保存到 Result 資料夾！")
            except Exception as e:
                messagebox.showerror("錯誤", f"複製檔案時發生錯誤: {e}")

    def write_results_to_excel(self):
        """讀取 HTML 報告，解析結果，並寫入 Excel"""
        html_dir = filedialog.askdirectory(title="選擇包含 HTML 報告的資料夾")
        if not html_dir:
            return

        excel_files_in_result = [f for f in os.listdir("Result") if f.endswith((".xlsx", ".xls"))]
        if not excel_files_in_result:
            messagebox.showwarning("警告", "Result 資料夾中沒有找到 Excel Testplan 檔案！請先載入。")
            return

        # 簡單地選擇 Result 資料夾中的第一個 Excel 檔案作為 Testplan
        # 你可以根據需要修改此邏輯，例如讓使用者選擇哪個 Testplan
        testplan_path = os.path.join("Result", excel_files_in_result[0])

        try:
            # 讀取 Excel Testplan
            workbook = load_workbook(testplan_path)
            sheet = workbook.active # 預設使用活動工作表

            # 獲取讀取和寫入的行列資訊
            read_col = self.read_col_entry.get().upper()
            read_row = int(self.read_row_entry.get()) - 1 # openpyxl 索引從 0 開始
            write_col = self.write_col_entry.get().upper()
            write_row = int(self.write_row_entry.get()) - 1

            if not read_col.isalpha() or not write_col.isalpha():
                messagebox.showerror("錯誤", "讀取/寫入欄位必須是字母 (例如 A, B)！")
                return
            if read_row < 0 or write_row < 0:
                messagebox.showerror("錯誤", "讀取/寫入行數必須是正整數！")
                return

            read_col_idx = ord(read_col) - ord('A') # 將列字母轉換為 0-based 索引
            write_col_idx = ord(write_col) - ord('A')

            # 讀取所有 HTML 報告
            html_files = [f for f in os.listdir(html_dir) if f.endswith(".html")]
            if not html_files:
                messagebox.showwarning("警告", "選擇的資料夾中沒有找到 HTML 報告檔案。")
                return

            all_html_results = {}
            for html_file in html_files:
                file_path = os.path.join(html_dir, html_file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    soup = BeautifulSoup(f, 'html.parser')
                    results = self.parse_html_report(soup)
                    for item in results:
                        all_html_results[item['name']] = item['result'] # 儲存所有 HTML 報告的結果

            # 比對並寫入 Excel
            for row_idx, row in enumerate(sheet.iter_rows()):
                if row_idx < read_row: # 跳過設定的讀取起始行之前的行
                    continue

                # 讀取 Testcase name
                testcase_name_cell = sheet.cell(row=row_idx + 1, column=read_col_idx + 1) # Excel cell 是 1-based
                testcase_name = str(testcase_name_cell.value).strip() if testcase_name_cell.value else ""

                if testcase_name in all_html_results:
                    result_to_write = all_html_results[testcase_name]
                    # 寫入結果
                    sheet.cell(row=row_idx + 1, column=write_col_idx + 1, value=result_to_write)
                    print(f"將 {testcase_name} 的結果 '{result_to_write}' 寫入到 {get_column_letter(write_col_idx + 1)}{row_idx + 1}")

            # 保存修改後的 Excel 檔案
            workbook.save(testplan_path)
            messagebox.showinfo("成功", "結果已成功寫入 Excel Testplan 並保存！")

        except Exception as e:
            messagebox.showerror("錯誤", f"寫入結果到 Excel 時發生錯誤: {e}")


    def parse_html_report(self, soup):
        """解析 HTML 報告並擷取測試案例名稱和結果"""
        all_test_results = []
        for tr_tag in soup.find_all('tr', class_='hiddenRow'):
            test_info = {}

            # 1. 尋找測試案例名稱
            name_tag = tr_tag.find('div', class_='testcase').find('a', class_='popup_link')
            if name_tag:
                test_info['name'] = name_tag.get_text(strip=True)
            else:
                test_info['name'] = "N/A (名稱未找到)"

            # 2. 尋找測試結果 (PASS/FAIL)
            result_tag = tr_tag.find('td', align='center')
            if result_tag:
                result_link = result_tag.find('button').find('a', class_='popup_link')
                if result_link:
                    test_info['result'] = result_link.get_text(strip=True)
                else:
                    test_info['result'] = "N/A (結果連結未找到)"
            else:
                test_info['result'] = "N/A (結果欄位未找到)"

            all_test_results.append(test_info)
        return all_test_results

if __name__ == "__main__":
    app = App()
    app.mainloop()