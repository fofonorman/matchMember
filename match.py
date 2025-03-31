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
            matches, repeated_pairs = matcher.match_people()
            
            # 儲存結果
            matcher.save_matching_result(matches, repeated_pairs)
            
            # 更新狀態
            if repeated_pairs:
                self.update_status(f"已完成，有重複配對，請檢查 excel 檔案")
            else:
                self.update_status("已完成，無重複配對")
            
            # 顯示配對結果
            result_text = "配對結果：\n" + "\n".join([" - ".join(match) for match in matches])
            if repeated_pairs:
                result_text += "\n\n重複配對：\n" + "\n".join([" - ".join(pair) for pair in repeated_pairs])
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
            participants_df = pd.DataFrame(columns=['姓名'])
            
            with pd.ExcelWriter(self.excel_path) as writer:
                people_df.to_excel(writer, sheet_name='人員名單', index=False)
                participants_df.to_excel(writer, sheet_name='參與配對人員', index=False)
            print(f"已在桌面創建新的 Excel 檔案：{excel_filename}")
    
    def get_all_people(self) -> List[str]:
        """獲取所有待配對人員名單"""
        df = pd.read_excel(self.excel_path, sheet_name='人員名單')
        return [name for name in df['姓名'].dropna().tolist()]
        
    def get_matching_history(self) -> Set[Tuple[str, ...]]:
        """從人員名單獲取歷史配對記錄"""
        history_set = set()
        
        try:
            # 讀取人員名單
            df = pd.read_excel(self.excel_path, sheet_name='人員名單')
            
            # 確保有「姓名」欄位
            if '姓名' not in df.columns:
                return history_set
            
            # 獲取所有配對者欄位（除了「姓名」以外的所有欄位）
            partner_columns = [col for col in df.columns if col != '姓名']
            
            # 如果沒有配對者欄位，返回空集合
            if not partner_columns:
                return history_set
            
            # 遍歷每一行（每個人）
            for _, row in df.iterrows():
                person = row['姓名']
                if not isinstance(person, str) or not person.strip():
                    continue
                    
                # 遍歷該人的所有配對者
                for col in partner_columns:
                    partner = row[col]
                    if isinstance(partner, str) and partner.strip():
                        # 將配對加入歷史記錄
                        pair = tuple(sorted([person, partner]))
                        history_set.add(pair)
            
            return history_set
        
        except Exception as e:
            print(f"讀取配對歷史時出錯: {str(e)}")
            return history_set
    
    def save_matching_result(self, matches: List[Tuple[str, ...]], repeated_pairs: List[Tuple[str, ...]] = None):
        """保存本次配對結果，並標記重複配對"""
        if repeated_pairs is None:
            repeated_pairs = []
            
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font
            
            # 設定重複配對的樣式
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_font = Font(color="FF0000", bold=True)
            
            # 讀取現有的 Excel 檔案
            workbook = openpyxl.load_workbook(self.excel_path)
            
            # 創建配對結果字典，方便查詢每個人的配對者
            match_dict = {}
            for match in matches:
                if len(match) == 2:
                    match_dict[match[0]] = [match[1]]
                    match_dict[match[1]] = [match[0]]
                elif len(match) == 3:
                    match_dict[match[0]] = [match[1], match[2]]
                    match_dict[match[1]] = [match[0], match[2]]
                    match_dict[match[2]] = [match[0], match[1]]
            
            # 更新人員名單工作表
            people_sheet = workbook['人員名單']
            
            # 獲取當前日期作為新欄位名稱
            import datetime
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            new_column_base = f"配對者 {today}"
            
            # 在「姓名」欄位的右側插入新欄位
            # 首先找到「姓名」欄位的位置
            name_col_idx = None
            for col_idx, cell in enumerate(people_sheet[1], 1):  # 第一行是標題行
                if cell.value == '姓名':
                    name_col_idx = col_idx
                    break
            
            if name_col_idx is None:
                # 如果找不到「姓名」欄位，假設它在第一列
                name_col_idx = 1
            
            # 找出最大配對者數量，決定要插入幾個欄位
            max_partners = max([len(partners) for partners in match_dict.values()], default=1)
            
            # 插入多個新欄位
            for i in range(max_partners):
                people_sheet.insert_cols(name_col_idx + 1 + i)
                # 設置新欄位的標題
                if max_partners > 1:
                    column_title = f"{new_column_base} {i+1}"
                else:
                    column_title = new_column_base
                people_sheet.cell(row=1, column=name_col_idx + 1 + i, value=column_title)
            
            # 獲取所有現有的姓名
            existing_names = []
            for row_idx, row in enumerate(people_sheet.iter_rows(min_row=2), 2):  # 從第二行開始
                name = row[name_col_idx - 1].value
                if name:
                    existing_names.append((row_idx, name))
            
            # 更新現有人員的配對結果
            for row_idx, name in existing_names:
                if name in match_dict:
                    partners = match_dict[name]
                    for i, partner in enumerate(partners):
                        cell = people_sheet.cell(row=row_idx, column=name_col_idx + 1 + i, value=partner)
                        
                        # 檢查是否為重複配對
                        for pair in repeated_pairs:
                            if name in pair and partner in pair:
                                # 這是重複配對，設定黃底紅字
                                cell.fill = yellow_fill
                                cell.font = red_font
            
            # 添加新人員
            existing_name_set = {name for _, name in existing_names}
            new_row_idx = len(existing_names) + 2  # 從最後一行之後開始
            
            for person, partners in match_dict.items():
                if person not in existing_name_set:
                    # 添加新行
                    people_sheet.cell(row=new_row_idx, column=name_col_idx, value=person)
                    for i, partner in enumerate(partners):
                        cell = people_sheet.cell(row=new_row_idx, column=name_col_idx + 1 + i, value=partner)
                        
                        # 檢查是否為重複配對
                        for pair in repeated_pairs:
                            if person in pair and partner in pair:
                                # 這是重複配對，設定黃底紅字
                                cell.fill = yellow_fill
                                cell.font = red_font
                    
                    new_row_idx += 1
            
            # 保存工作簿
            workbook.save(self.excel_path)
            
        except FileNotFoundError:
            # 如果檔案不存在，創建新的檔案
            import pandas as pd
            
            # 創建配對結果字典
            match_dict = {}
            for match in matches:
                if len(match) == 2:
                    match_dict[match[0]] = [match[1]]
                    match_dict[match[1]] = [match[0]]
                elif len(match) == 3:
                    match_dict[match[0]] = [match[1], match[2]]
                    match_dict[match[1]] = [match[0], match[2]]
                    match_dict[match[2]] = [match[0], match[1]]
            
            # 創建人員名單 DataFrame
            people_list = list(set([person for match in matches for person in match]))
            
            import datetime
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            
            # 找出最大配對者數量
            max_partners = max([len(partners) for partners in match_dict.values()], default=1)
            
            # 創建 DataFrame
            people_data = {'姓名': people_list}
            
            # 添加配對者欄位
            for i in range(max_partners):
                if max_partners > 1:
                    column_name = f"配對者 {today} {i+1}"
                else:
                    column_name = f"配對者 {today}"
                    
                people_data[column_name] = [
                    match_dict.get(person, [])[i] if person in match_dict and i < len(match_dict[person]) else ''
                    for person in people_list
                ]
            
            people_df = pd.DataFrame(people_data)
            
            # 創建參與配對人員 DataFrame
            participants_df = pd.DataFrame(columns=['姓名'])
            
            with pd.ExcelWriter(self.excel_path) as writer:
                people_df.to_excel(writer, sheet_name='人員名單', index=False)
                participants_df.to_excel(writer, sheet_name='參與配對人員', index=False)
            
            # 註意：如果是新建檔案，需要再次打開來設定樣式
            if repeated_pairs:
                # 再次打開檔案來設定樣式
                workbook = openpyxl.load_workbook(self.excel_path)
                people_sheet = workbook['人員名單']
                
                # 找到姓名列的索引
                name_col_idx = None
                for col_idx, cell in enumerate(people_sheet[1], 1):
                    if cell.value == '姓名':
                        name_col_idx = col_idx
                        break
                
                if name_col_idx is None:
                    name_col_idx = 1
                
                # 設定重複配對的樣式
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                red_font = Font(color="FF0000", bold=True)
                
                # 遍歷每一行
                for row_idx in range(2, people_sheet.max_row + 1):
                    person = people_sheet.cell(row=row_idx, column=name_col_idx).value
                    if not person:
                        continue
                    
                    # 遍歷配對者欄位
                    for col_idx in range(name_col_idx + 1, name_col_idx + 1 + max_partners):
                        partner = people_sheet.cell(row=row_idx, column=col_idx).value
                        if not partner:
                            continue
                        
                        # 檢查是否為重複配對
                        for pair in repeated_pairs:
                            if person in pair and partner in pair:
                                # 這是重複配對，設定黃底紅字
                                cell = people_sheet.cell(row=row_idx, column=col_idx)
                                cell.fill = yellow_fill
                                cell.font = red_font
                
                # 保存工作簿
                workbook.save(self.excel_path)

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

    def match_people(self) -> Tuple[List[Tuple[str, ...]], List[Tuple[str, ...]]]:
        """
        配對人員並返回配對結果和重複配對列表
        返回: (matches, repeated_pairs)
        """
        # 從「參與配對人員」分頁獲取本次參與配對的人員
        try:
            participants_df = pd.read_excel(self.excel_path, sheet_name='參與配對人員')
            people = [name for name in participants_df['姓名'].dropna().tolist()]
        except Exception as e:
            # 如果讀取失敗，顯示錯誤訊息
            raise Exception(f"無法讀取參與配對人員: {str(e)}")
        
        if not people:
            raise Exception("參與配對人員名單為空")
        
        history = self.get_matching_history()
        
        # 找出一個配對方案中的重複配對
        def find_repeated_pairs(matches: List[Tuple[str, ...]]) -> List[Tuple[str, ...]]:
            repeated = []
            for match in matches:
                # 檢查所有可能的2人組合
                for combo in combinations(match, 2):
                    sorted_combo = tuple(sorted(combo))
                    if sorted_combo in history and sorted_combo not in repeated:
                        repeated.append(sorted_combo)
            return repeated
        
        # 計算一個配對方案中的重複配對數量
        def count_repeated_pairs(matches: List[Tuple[str, ...]]) -> int:
            return len(find_repeated_pairs(matches))
        
        # 使用回溯法找出所有可能的配對方案
        def find_all_matchings(remaining: List[str], current_matches: List[Tuple[str, ...]]) -> List[List[Tuple[str, ...]]]:
            if not remaining:  # 基本情況：沒有剩餘的人要配對
                return [current_matches]
                
            all_matchings = []
            
            if len(remaining) == 2:  # 只剩兩個人
                pair = tuple(sorted(remaining))
                all_matchings.extend(find_all_matchings([], current_matches + [pair]))
                
            elif len(remaining) == 3:  # 只剩三個人
                trio = tuple(sorted(remaining))
                all_matchings.extend(find_all_matchings([], current_matches + [trio]))
                
            else:  # 至少有4個人，可以選擇2人一組
                # 固定第一個人，嘗試與其他每個人配對
                first_person = remaining[0]
                new_remaining = remaining[1:]
                
                for i in range(len(new_remaining)):
                    second_person = new_remaining[i]
                    pair = tuple(sorted([first_person, second_person]))
                    
                    # 準備下一輪遞迴的剩餘人員名單
                    next_remaining = new_remaining.copy()
                    next_remaining.pop(i)
                    
                    # 遞迴找尋剩餘人員的所有可能配對
                    pair_matchings = find_all_matchings(next_remaining, current_matches + [pair])
                    all_matchings.extend(pair_matchings)
            
            return all_matchings
        
        # 主要配對邏輯
        print("開始尋找所有可能的配對方案...")
        
        # 先嘗試找出沒有重複配對的方案（提早終止條件）
        def try_no_repeats(remaining: List[str], current_matches: List[Tuple[str, ...]]) -> Tuple[bool, List[Tuple[str, ...]]]:
            if not remaining:
                return True, current_matches
            
            if len(remaining) == 2:
                pair = tuple(sorted(remaining))
                if self.is_valid_pair(pair, history):
                    return True, current_matches + [pair]
                return False, current_matches
                
            if len(remaining) == 3:
                trio = tuple(sorted(remaining))
                if self.is_valid_pair(trio, history):
                    return True, current_matches + [trio]
                return False, current_matches
            
            # 試著先匹配沒有歷史記錄的配對
            first_person = remaining[0]
            new_remaining = remaining[1:]
            
            # 隨機打亂以增加找到解的可能性
            random.shuffle(new_remaining)
            
            for i in range(len(new_remaining)):
                second_person = new_remaining[i]
                pair = tuple(sorted([first_person, second_person]))
                
                if self.is_valid_pair(pair, history):
                    next_remaining = new_remaining.copy()
                    next_remaining.pop(i)
                    
                    success, matches = try_no_repeats(next_remaining, current_matches + [pair])
                    if success:
                        return True, matches
            
            return False, current_matches
        
        # 首先嘗試找到一個無重複的方案（這比窮舉要快得多）
        for _ in range(100):  # 多試幾次隨機順序
            random.shuffle(people)
            success, matches = try_no_repeats(people, [])
            if success:
                print("找到了無重複的配對方案！")
                return matches, []  # 無重複配對
        
        # 如果人數超過特定閾值，直接使用次優解方案
        if len(people) > 10:  # 根據實際需求調整閾值
            print("參與人數過多，使用啟發式方法尋找次優解...")
            
            best_solution = None
            best_score = float('inf')
            fallback_attempts = 1000  # 增加嘗試次數以找到更好的解
            
            for _ in range(fallback_attempts):
                all_people = people.copy()
                random.shuffle(all_people)
                matches = []
                
                while len(all_people) >= 2:
                    if len(all_people) == 2:
                        matches.append(tuple(sorted(all_people)))
                        all_people = []
                    elif len(all_people) == 3:
                        matches.append(tuple(sorted(all_people)))
                        all_people = []
                    else:
                        matches.append(tuple(sorted(all_people[:2])))
                        all_people = all_people[2:]
                
                score = count_repeated_pairs(matches)
                
                if score < best_score:
                    best_score = score
                    best_solution = matches
                    
                    if score == 0:  # 找到無重複解，立即返回
                        repeated_pairs = find_repeated_pairs(best_solution)
                        return best_solution, repeated_pairs
            
            print(f"已找到最佳次優解決方案，重複配對數: {best_score}")
            repeated_pairs = find_repeated_pairs(best_solution)
            return best_solution, repeated_pairs
        
        # 對於人數較少的情況，使用窮舉法尋找所有可能的配對方案
        print("開始窮舉所有可能的配對方案...")
        all_possible_matchings = find_all_matchings(people, [])
        
        print(f"共找到 {len(all_possible_matchings)} 種可能的配對方案")
        
        # 找出重複配對最少的方案
        best_matching = None
        min_repeats = float('inf')
        
        for matching in all_possible_matchings:
            repeats = count_repeated_pairs(matching)
            if repeats < min_repeats:
                min_repeats = repeats
                best_matching = matching
                
                # 如果找到完全無重複的方案，立即返回
                if repeats == 0:
                    print("找到了無重複的配對方案！")
                    return best_matching, []  # 無重複配對
        
        if best_matching:
            print(f"已找到最佳配對方案，重複配對數: {min_repeats}")
            repeated_pairs = find_repeated_pairs(best_matching)
            return best_matching, repeated_pairs
        else:
            raise Exception("無法完成配對，請管理員手動調整")

def main():
    app = MatchingGUI()
    app.run()

if __name__ == "__main__":
    main()
