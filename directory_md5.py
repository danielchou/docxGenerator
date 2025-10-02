import hashlib
import os
from pathlib import Path

def calculate_directory_md5(directory_path, include_filenames=True, sort_files=True):
    """
    計算整個目錄的單一 MD5 值
    
    Args:
        directory_path: 目錄路徑
        include_filenames: 是否將檔案名稱也納入 MD5 計算（預設 True）
        sort_files: 是否對檔案進行排序以確保一致性（預設 True）
    
    Returns:
        str: 目錄的 MD5 值，如果失敗則返回 None
    """
    directory_path = Path(directory_path)
    
    if not directory_path.exists():
        print(f"錯誤: 目錄 {directory_path} 不存在")
        return None
    
    if not directory_path.is_dir():
        print(f"錯誤: {directory_path} 不是一個目錄")
        return None
    
    hash_md5 = hashlib.md5()
    file_count = 0
    
    try:
        # 收集所有檔案
        all_files = []
        for file_path in directory_path.rglob('*'):
            if file_path.is_file():
                all_files.append(file_path)
        
        # 排序檔案以確保一致性
        if sort_files:
            all_files.sort()
        
        print(f"正在計算目錄 MD5: {directory_path}")
        print(f"找到 {len(all_files)} 個檔案")
        print("=" * 50)
        
        for file_path in all_files:
            file_count += 1
            relative_path = file_path.relative_to(directory_path)
            print(f"處理檔案 ({file_count}/{len(all_files)}): {relative_path}")
            
            # 如果要包含檔案名稱，先將檔案路徑加入 hash
            if include_filenames:
                hash_md5.update(str(relative_path).encode('utf-8'))
            
            # 讀取檔案內容並加入 hash
            try:
                with open(file_path, 'rb') as f:
                    while chunk := f.read(8192):
                        hash_md5.update(chunk)
            except Exception as e:
                print(f"警告: 無法讀取檔案 {relative_path}: {str(e)}")
                continue
        
        directory_md5 = hash_md5.hexdigest()
        print("=" * 50)
        print(f"✅ 目錄 MD5 計算完成!")
        print(f"📁 目錄: {directory_path}")
        print(f"🔐 MD5: {directory_md5}")
        print(f"📊 檔案數量: {file_count}")
        
        return directory_md5
        
    except Exception as e:
        print(f"錯誤: 計算目錄 MD5 時發生錯誤")
        print(f"錯誤詳情: {str(e)}")
        return None

def save_directory_md5_result(directory_path, md5_hash, output_file=None):
    """保存目錄 MD5 結果到檔案"""
    directory_path = Path(directory_path)
    directory_name = directory_path.name  # 只取目錄名稱，不要完整路徑
    
    if output_file is None:
        output_file = f"{directory_name}_directory_md5.txt"
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"目錄 MD5 計算結果\n")
            f.write(f"=" * 30 + "\n")
            f.write(f"目錄名稱: {directory_name}\n")  # 只顯示目錄名稱
            f.write(f"計算時間: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"目錄 MD5: {md5_hash}\n")
            f.write(f"\n注意事項:\n")
            f.write(f"- 此 MD5 值代表整個 '{directory_name}' 目錄的內容指紋\n")
            f.write(f"- 任何檔案的新增、刪除、修改都會改變此 MD5 值\n")
            f.write(f"- 檔案名稱和路徑也會影響 MD5 計算\n")
        
        print(f"✅ 結果已保存到: {output_file}")
        return True
    except Exception as e:
        print(f"錯誤: 無法保存結果到 {output_file}")
        print(f"錯誤詳情: {str(e)}")
        return False

def main():
    """主函數"""
    print("=" * 60)
    print("📁 目錄 MD5 計算工具")
    print("=" * 60)
    print("此工具會計算整個目錄的單一 MD5 值")
    print("適用於目錄版本比較和完整性驗證")
    print()
    
    while True:
        # 讓使用者輸入目標目錄
        target_dir = input("請輸入目標目錄路徑 (輸入 'q' 退出): ").strip()
        
        if target_dir.lower() == 'q':
            print("👋 程式結束")
            break
        
        if not target_dir:
            print("⚠️  請輸入有效的目錄路徑")
            continue
        
        # 使用固定的計算選項（根據使用者偏好設定）
        include_filenames = True   # 將檔案名稱納入 MD5 計算
        sort_files_flag = False    # 不對檔案進行排序
        
        print("\n🔧 計算選項:")
        print("✅ 將檔案名稱納入 MD5 計算: 是")
        print("✅ 對檔案進行排序: 否")
        
        print()
        
        # 計算目錄 MD5
        md5_result = calculate_directory_md5(
            target_dir, 
            include_filenames=include_filenames,
            sort_files=sort_files_flag
        )
        
        if md5_result:
            # 詢問是否保存結果
            save_choice = input("\n是否要將結果保存到檔案? (y/n): ").strip().lower()
            if save_choice == 'y':
                save_directory_md5_result(target_dir, md5_result)
        
        print("\n" + "=" * 60)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程式被使用者中斷")
    except Exception as e:
        print(f"\n❌ 程式發生未預期的錯誤: {str(e)}")
