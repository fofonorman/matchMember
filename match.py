import random
import pandas as pd
from typing import List, Tuple, Set
from itertools import combinations
import os
import tkinter as tk
from tkinter import messagebox

class MatchingGUI:
    def __init__(self):
        # 創建主視窗
        self.window = tk.Tk()
        self.window.title("配對系統")
        self.window.geometry("400x300")  # 加大視窗高度以容納文字框
        
        # Excel 檔案名稱輸入區域
        tk.Label(self.window, text="Excel 檔案名稱：").pack(pady=10)
        self.filename_var = tk.StringVar(value="配對系統.xlsx")
        tk.Entry(self.window, textvariable=self.filename_var, width=30).pack()
        
        # 提示文字
        tk.Label(self.window, text="(檔案將存放在桌面，請包含 .xlsx 副檔名)").pack(pady=5)
        
        # 配對按鈕
        self.match_button = tk.Button(self.window, text="配對", command=self.do_matching)
        self.match_button.pack(pady=20)
        
        # 狀態顯示（使用文字框替代標籤）
        self.status_text = tk.Text(self.window, height=3, width=35)
        self.status_text.pack(pady=10)
        self.status_text.config(state='disabled')  # 預設為不可編輯
        
    def update_status(self, message: str, is_error: bool = False):
        """更新狀態文字框的內容"""
        self.status_text.config(state='normal')  # 暫時允許編輯
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, message)
        if is_error:
            self.status_text.config(fg="red")
        else:
            self.status_text.config(fg="green")
        self.status_text.config(state='disabled')  # 恢復為不可編輯
        
    def do_matching(self):
        try:
            filename = self.filename_var.get()
            
            # 檢查檔案名稱
            if not filename.endswith('.xlsx'):
                self.update_status("失敗：檔案名稱必須以 .xlsx 結尾", True)
                return
            
            # 建立配對系統實例
            matcher = MatchingSystem(filename)
            
            # 執行配對
            matches = matcher.match_people()
            
            # 儲存結果
            matcher.save_matching_result(matches)
            
            # 更新狀態
            self.update_status("已完成")
            
            # 顯示配對結果
            result_text = "配對結果：\n" + "\n".join([" - ".join(match) for match in matches])
            messagebox.showinfo("配對完成", result_text)
            
        except Exception as e:
            self.update_status(f"失敗：{str(e)}", True)
    
    def run(self):
        self.window.mainloop()

class MatchingSystem:
    def __init__(self, excel_filename: str):
        # 獲取桌面路徑
        self.desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        # 完整的 Excel 檔案路徑
        self.excel_path = os.path.join(self.desktop_path, excel_filename)
        
        try:
            # 嘗試讀取現有的 Excel 檔案
            self.excel = pd.ExcelFile(self.excel_path)
        except FileNotFoundError:
            # 如果檔案不存在，創建新的 Excel 檔案
            people_df = pd.DataFrame(columns=['姓名'])
            history_df = pd.DataFrame(columns=['配對1', '配對2', '配對3'])
            
            with pd.ExcelWriter(self.excel_path) as writer:
                people_df.to_excel(writer, sheet_name='人員名單', index=False)
                history_df.to_excel(writer, sheet_name='配對歷史', index=False)
            print(f"已在桌面創建新的 Excel 檔案：{excel_filename}")
    
    def get_all_people(self) -> List[str]:
        """獲取所有待配對人員名單"""
        df = pd.read_excel(self.excel_path, sheet_name='人員名單')
        return [name for name in df['姓名'].dropna().tolist()]
        
    def get_matching_history(self) -> Set[Tuple[str, ...]]:
        """獲取歷史配對記錄"""
        history_set = set()
        df = pd.read_excel(self.excel_path, sheet_name='配對歷史')
        
        for _, row in df.iterrows():
            record = [r for r in row.tolist() if isinstance(r, str) and r.strip()]
            if len(record) == 2:
                history_set.add(tuple(sorted([record[0], record[1]])))
            elif len(record) == 3:
                history_set.add(tuple(sorted([record[0], record[1], record[2]])))
        return history_set
    
    def save_matching_result(self, matches: List[Tuple[str, ...]]):
        """保存本次配對結果"""
        try:
            # 先讀取所有需要的資料
            existing_history = pd.read_excel(self.excel_path, sheet_name='配對歷史', engine='openpyxl')
            existing_people = pd.read_excel(self.excel_path, sheet_name='人員名單', engine='openpyxl')
            
            # 準備新的配對結果
            new_matches = []
            for match in matches:
                row = list(match) + [''] * (3 - len(match))
                new_matches.append(row)
            
            # 將新的配對結果轉換為 DataFrame
            new_df = pd.DataFrame(new_matches, columns=['配對1', '配對2', '配對3'])
            
            # 合併現有記錄和新記錄
            history_df = pd.concat([existing_history, new_df], ignore_index=True)
            
            # 移除重複的記錄
            history_df = history_df.drop_duplicates()
            
            # 寫入所有資料
            with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
                # 保存原有的人員名單
                existing_people.to_excel(writer, sheet_name='人員名單', index=False)
                # 保存更新後的配對歷史
                history_df.to_excel(writer, sheet_name='配對歷史', index=False)
            
        except FileNotFoundError:
            # 如果檔案不存在，創建新的檔案
            new_matches = []
            for match in matches:
                row = list(match) + [''] * (3 - len(match))
                new_matches.append(row)
            
            history_df = pd.DataFrame(new_matches, columns=['配對1', '配對2', '配對3'])
            people_df = pd.DataFrame(columns=['姓名'])
            
            with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
                people_df.to_excel(writer, sheet_name='人員名單', index=False)
                history_df.to_excel(writer, sheet_name='配對歷史', index=False)

    def is_valid_pair(self, pair: Tuple[str, ...], history: Set[Tuple[str, ...]]) -> bool:
        """
        檢查配對是否有效
        - 檢查所有可能的2人和3人子組合是否出現在歷史記錄中
        """
        # 將 pair 中的所有可能 2 人組合檢查是否在歷史記錄中
        for combo in combinations(pair, 2):
            sorted_combo = tuple(sorted(combo))
            if sorted_combo in history:
                return False
            
        # 如果是 3 人組合，還需要檢查完整的組合
        if len(pair) == 3:
            sorted_pair = tuple(sorted(pair))
            if sorted_pair in history:
                return False
            
        return True

    def match_people(self) -> List[Tuple[str, ...]]:
        people = self.get_all_people()
        history = self.get_matching_history()
        
        def try_matching(remaining_people: List[str], current_matches: List[Tuple[str, ...]], 
                        max_attempts: int = 10) -> Tuple[bool, List[Tuple[str, ...]]]:
            if not remaining_people:
                return True, current_matches
            
            if len(remaining_people) == 2:
                pair = tuple(sorted(remaining_people))
                if self.is_valid_pair(pair, history):
                    return True, current_matches + [pair]
                return False, current_matches
                
            if len(remaining_people) == 3:
                trio = tuple(sorted(remaining_people))
                if self.is_valid_pair(trio, history):
                    return True, current_matches + [trio]
                return False, current_matches
            
            attempts = 0
            while attempts < max_attempts:
                random.shuffle(remaining_people)
                pair = tuple(sorted(remaining_people[:2]))
                
                if self.is_valid_pair(pair, history):
                    success, new_matches = try_matching(
                        remaining_people[2:],
                        current_matches + [pair]
                    )
                    if success:
                        return True, new_matches
                
                attempts += 1
            
            return False, current_matches

        # 主要配對邏輯
        max_total_attempts = 20
        total_attempts = 0
        
        while total_attempts < max_total_attempts:
            random.shuffle(people)
            success, matches = try_matching(people, [])
            
            if success:
                return matches
                
            total_attempts += 1
        
        raise Exception("無法完成配對，請管理員手動調整")

def main():
    app = MatchingGUI()
    app.run()

if __name__ == "__main__":
    main()
