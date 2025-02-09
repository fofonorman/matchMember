import random
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from typing import List, Tuple, Set
from itertools import combinations

class MatchingSystem:
    def __init__(self, spreadsheet_name: str):
        # 設定 Google Sheets API 認證
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('liquid-almanac-450023-i0-7e8d6798c29d.json', scope)
        client = gspread.authorize(creds)
        
        # 開啟試算表
        self.sheet = client.open_by_key("1sRK4goqZhpxEDjKKN5BkCG59XYeMKWhPAo8PX6bVaO4")  # 替換成你複製的 ID
        self.people_worksheet = self.sheet.worksheet("人員名單")
        self.history_worksheet = self.sheet.worksheet("歷史配對")
        
    def get_all_people(self) -> List[str]:
        """獲取所有待配對人員名單"""
        return self.people_worksheet.col_values(1)[1:]  # 跳過標題行
        print("test")
    def get_matching_history(self) -> Set[Tuple[str, ...]]:
        """獲取歷史配對記錄"""
        history_records = self.history_worksheet.get_all_values()[1:]  # 跳過標題行
        history_set = set()
        for record in history_records:
            if len(record) == 2:
                history_set.add(tuple(sorted([record[0], record[1]])))
            elif len(record) == 3:
                history_set.add(tuple(sorted([record[0], record[1], record[2]])))
        return history_set

    def save_matching_result(self, matches: List[Tuple[str, ...]]):
        """保存本次配對結果"""
        row = len(self.history_worksheet.get_all_values()) + 1
        for match in matches:
            self.history_worksheet.insert_row(list(match), row)
            row += 1
        
    def is_valid_pair(self, pair: Tuple[str, ...], history: Set[Tuple[str, ...]]) -> bool:
        """檢查配對是否有效（未在歷史記錄中）"""
        return tuple(sorted(pair)) not in history

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
    matcher = MatchingSystem("配對系統試算表")
    try:
        matches = matcher.match_people()
        print("配對成功！")
        for match in matches:
            print(f"配對組合: {' - '.join(match)}")
    except Exception as e:
        print(f"錯誤: {str(e)}")

if __name__ == "__main__":
    main()
