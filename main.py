"""
自動化爬蟲與轉換主程式
依序執行：
1. 爬取 URL 並生成 Markdown 檔案
2. 將 Markdown 檔案合併轉換為 DOCX/PDF
3. 刪除中間生成的 Markdown 檔案
"""

import sys
import os
import glob
from scrape_to_markdown import process_urls_from_file
from convert_md_to_docx import main as convert_main

def cleanup_markdown_files():
    """刪除所有生成的 Markdown 檔案"""
    print("\n步驟 3: 清理中間檔案...")
    print("=" * 60)
    
    # 找出所有符合 *_數字.md 格式的檔案
    md_files = glob.glob("*_*.md")
    md_files = [f for f in md_files if f.endswith('.md')]
    
    if not md_files:
        print("沒有找到需要刪除的 Markdown 檔案")
        return
    
    deleted_count = 0
    for file in md_files:
        try:
            os.remove(file)
            print(f"已刪除：{file}")
            deleted_count += 1
        except Exception as e:
            print(f"無法刪除 {file}: {e}")
    
    print(f"\n清理完成，共刪除 {deleted_count} 個 Markdown 檔案")
    print("=" * 60)

def main():
    """主程式：自動執行爬蟲和轉換流程"""
    # 檢查 urls.txt 是否存在
    txt_file = "urls.txt"
    if not os.path.exists(txt_file):
        print(f"錯誤：找不到檔案 {txt_file}")
        print("請創建一個 urls.txt 檔案，每行一個 URL")
        return 1
    
    # 步驟 1: 爬取 URL 並生成 Markdown 檔案
    try:
        process_urls_from_file(txt_file)
    except Exception as e:
        print(f"爬取過程發生錯誤: {e}")
        return 1
    
    # 步驟 2: 將 Markdown 檔案合併轉換為 DOCX/PDF
    print()  # 空行分隔兩個步驟
    try:
        result = convert_main()
        
        # 如果轉換成功，刪除中間生成的 Markdown 檔案
        if result == 0:
            cleanup_markdown_files()
        
        return result
    except Exception as e:
        print(f"轉換過程發生錯誤: {e}")
        return 1

if __name__ == '__main__':
    sys.exit(main())
