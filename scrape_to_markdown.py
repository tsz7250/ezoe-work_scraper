import requests
from bs4 import BeautifulSoup, NavigableString
import re

def fullwidth_to_halfwidth(text):
    """
    將全型英文字母和數字轉換為半型
    """
    result = []
    for char in text:
        code = ord(char)
        # 全型英文字母和數字的範圍是 0xFF01-0xFF5E
        # 對應的半型是 0x0021-0x007E
        if 0xFF01 <= code <= 0xFF5E:
            # 轉換為半型
            result.append(chr(code - 0xFEE0))
        else:
            result.append(char)
    return ''.join(result)

def convert_quotes(text):
    """
    將各種英文引號轉換為中文引號
    "" → 「」
    "" → 「」 (彎引號)
    '' → 『』 (單引號)
    """
    # 處理各種類型的雙引號
    result = []
    in_quote = False
    for i, char in enumerate(text):
        char_code = ord(char)
        
        # ASCII 雙引號 " (U+0022)
        if char_code == 0x0022:
            if in_quote:
                result.append('」')
                in_quote = False
            else:
                result.append('「')
                in_quote = True
        # 左雙引號 " (U+201C)
        elif char_code == 0x201C:
            result.append('「')
        # 右雙引號 " (U+201D)
        elif char_code == 0x201D:
            result.append('」')
        # 左單引號 ' (U+2018)
        elif char_code == 0x2018:
            result.append('『')
        # 右單引號 ' (U+2019)
        elif char_code == 0x2019:
            result.append('』')
        # ASCII 單引號 ' (U+0027)
        elif char_code == 0x0027:
            # 簡單處理：如果前後都是字母，可能是所有格，保留
            if i > 0 and i < len(text) - 1 and text[i-1].isalpha() and text[i+1].isalpha():
                result.append(char)
            else:
                result.append('『' if i == 0 or not text[i-1].isalpha() else '』')
        else:
            result.append(char)
    return ''.join(result)

def extract_book_name(soup):
    """
    從網頁中提取書名
    從 <a class="brandtitle header__link"> 元素中提取
    """
    brandtitle = soup.find('a', class_='brandtitle')
    if brandtitle:
        return brandtitle.get_text(strip=True)
    return None

def extract_article_number_from_url(url):
    """
    從 URL 中提取篇數
    例如：從 "https://ezoe.work/books/3/3061-11.html" 提取 "11"
    """
    # 匹配 URL 中最後的數字（在 .html 之前）
    match = re.search(r'-(\d+)\.html', url)
    if match:
        return match.group(1)
    return None

def scrape_to_markdown(url, output_file=None):
    """
    從指定網頁抓取內容並轉換為 Markdown 格式
    移除所有超連結及其內部文字，保持原有排版和分段
    """
    # 獲取網頁內容
    response = requests.get(url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # 如果沒有指定輸出檔案名，自動生成
    if output_file is None:
        book_name = extract_book_name(soup)
        # 從 URL 中提取篇數
        article_num = extract_article_number_from_url(url)
        
        # 根據書名和篇數生成檔案名
        if book_name and article_num:
            output_file = f"{book_name}_{article_num}.md"
        elif book_name:
            output_file = f"{book_name}.md"
        elif article_num:
            output_file = f"第{article_num}篇.md"
        else:
            output_file = "output.md"
    
    markdown_lines = []
    
    # 1. 提取主標題（feature 區域）
    feature = soup.find('header', class_='feature')
    if feature:
        chap_titles = feature.find_all('div', id='chap1')
        for title in chap_titles:
            title_text = title.get_text(strip=True)
            markdown_lines.append(f"# {title_text}\n")
    
    # 2. 提取主要內容區域
    main = soup.find('div', class_='main')
    if not main:
        print("警告：找不到主要內容區域")
        return
    
    # 遍歷主要內容區域的所有直接子元素
    for element in main.find_all(recursive=False):
        # 處理 cn1 層級標題（二級標題）
        if element.get('class') and 'cn1' in element.get('class'):
            title_text = fullwidth_to_halfwidth(element.get_text(strip=True))
            markdown_lines.append(f"\n## {title_text}\n")
        
        # 處理 cn2 層級標題（三級標題）
        elif element.get('class') and 'cn2' in element.get('class'):
            title_text = fullwidth_to_halfwidth(element.get_text(strip=True))
            markdown_lines.append(f"\n### {title_text}\n")
        
        # 處理 cn3 層級標題（四級標題）
        elif element.get('class') and 'cn3' in element.get('class'):
            title_text = fullwidth_to_halfwidth(element.get_text(strip=True))
            markdown_lines.append(f"\n#### {title_text}\n")
        
        # 處理 cn4 層級標題（五級標題）
        elif element.get('class') and 'cn4' in element.get('class'):
            title_text = fullwidth_to_halfwidth(element.get_text(strip=True))
            markdown_lines.append(f"\n##### {title_text}\n")
        
        # 處理內容區域（id='c' 的 div）
        elif element.get('id') == 'c':
            # 先嘗試尋找所有 cont 類的段落
            cont_divs = element.find_all('div', class_='cont')
            
            if cont_divs:
                # 如果有 cont 子元素，處理每個 cont
                for cont in cont_divs:
                    # 先移除所有 modal div 元素
                    for modal in cont.find_all('div', class_='modal'):
                        modal.decompose()
                    
                    # 將 <a> 標籤替換為其文字內容（保留顯示的文字，移除超連結）
                    for a_tag in cont.find_all('a'):
                        a_tag.replace_with(a_tag.get_text())
                    
                    # 獲取純文字內容
                    text = cont.get_text(strip=True)
                    if text:
                        markdown_lines.append(f"\n{text}\n")
            else:
                # 如果沒有 cont 子元素，直接提取 id='c' 元素的內容
                # 先移除所有 modal div 元素
                for modal in element.find_all('div', class_='modal'):
                    modal.decompose()
                
                # 將 <a> 標籤替換為其文字內容
                for a_tag in element.find_all('a'):
                    a_tag.replace_with(a_tag.get_text())
                
                # 獲取純文字內容
                text = element.get_text(strip=True)
                if text:
                    markdown_lines.append(f"\n{text}\n")
    
    # 3. 轉換英文引號為中文引號
    content = ''.join(markdown_lines)
    content = convert_quotes(content)
    
    # 4. 寫入 Markdown 文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"成功生成 Markdown 文件：{output_file}")
    print(f"總行數：{len(markdown_lines)}")

def process_urls_from_file(txt_file):
    """
    從 txt 檔案讀取 URL 列表並批次處理
    每行一個 URL
    """
    try:
        with open(txt_file, 'r', encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip() and not line.strip().startswith('#')]
        
        total = len(urls)
        print(f"讀取到 {total} 個 URL")
        print("=" * 50)
        
        success_count = 0
        fail_count = 0
        
        for idx, url in enumerate(urls, 1):
            print(f"\n[{idx}/{total}] 處理：{url}")
            try:
                scrape_to_markdown(url)
                success_count += 1
            except Exception as e:
                print(f"錯誤：{e}")
                fail_count += 1
            print("-" * 50)
        
        print(f"\n完成！成功：{success_count}，失敗：{fail_count}")
        
    except FileNotFoundError:
        print(f"錯誤：找不到檔案 {txt_file}")
        print("請創建一個 urls.txt 檔案，每行一個 URL")

if __name__ == '__main__':
    import sys
    
    # 檢查是否有命令列參數
    if len(sys.argv) > 1:
        txt_file = sys.argv[1]
    else:
        txt_file = "urls.txt"
    
    print(f"從檔案讀取 URL：{txt_file}")
    process_urls_from_file(txt_file)
