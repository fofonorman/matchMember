import random
import pandas as pd
from typing import List, Tuple, Set
from itertools import combinations
import os

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
            
            # 創建 ExcelWriter 物件
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
            # 過濾掉 NaN 值並轉換為列表
            record = [r for r in row.tolist() if isinstance(r, str) and r.strip()]
            if len(record) == 2:
                history_set.add(tuple(sorted([record[0], record[1]])))
            elif len(record) == 3:
                history_set.add(tuple(sorted([record[0], record[1], record[2]])))

        print(history_set)
        return history_set

    def save_matching_result(self, matches: List[Tuple[str, ...]]):
        """保存本次配對結果"""
        # 讀取現有的配對歷史
        df = pd.read_excel(self.excel_path, sheet_name='配對歷史')
        
        # 準備新的配對結果
        new_matches = []
        for match in matches:
            # 如果是二人配對，第三個位置補空值
            row = list(match) + [''] * (3 - len(match))
            new_matches.append(row)
        
        # 將新的配對結果轉換為 DataFrame
        new_df = pd.DataFrame(new_matches, columns=['配對1', '配對2', '配對3'])
        
        # 合併現有記錄和新記錄
        df = pd.concat([df, new_df], ignore_index=True)
        
        # 寫回 Excel，保留其他工作表
        with pd.ExcelWriter(self.excel_path, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='配對歷史', index=False)

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
            
            # 嘗試配對兩人
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
                self.save_matching_result(matches)
                return matches
                
            total_attempts += 1
        
        raise Exception("無法完成配對，請管理員手動調整")

def main():
    matcher = MatchingSystem("配對系統.xlsx")
    try:
        matches = matcher.match_people()
        print("配對成功！")
        for match in matches:
            print(f"配對組合: {' - '.join(match)}")
    except Exception as e:
        print(f"錯誤: {str(e)}")

if __name__ == "__main__":
    main()
