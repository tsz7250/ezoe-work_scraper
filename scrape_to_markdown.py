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

def extract_book_name(soup, url=None):
    """
    從網頁中提取書名
    從 <a class="brandtitle header__link"> 元素中提取
    對於三級結構（X-Y-Z.html），會嘗試從書籍首頁（X.html）獲取真正的書名
    """
    brandtitle = soup.find('a', class_='brandtitle')
    if brandtitle:
        book_name = brandtitle.get_text(strip=True)
        
        # 檢測是否為三級結構（brandtitle 包含卷/冊等關鍵字）
        if url and re.search(r'(卷[一二三四五六七八九十百]+|第\d+[册冊]|第[一二三四五六七八九十]+[册冊])', book_name):
            # 嘗試從 URL 提取書號並獲取書籍首頁的書名
            match = re.search(r'/books/\d+/(\d+)-\d+-\d+\.html', url)
            if match:
                book_id = match.group(1)
                book_base_url = re.sub(r'/(\d+)-\d+-\d+\.html', f'/{book_id}.html', url)
                try:
                    response = requests.get(book_base_url)
                    response.encoding = 'utf-8'
                    base_soup = BeautifulSoup(response.text, 'html.parser')
                    base_brandtitle = base_soup.find('a', class_='brandtitle')
                    if base_brandtitle:
                        book_name = base_brandtitle.get_text(strip=True)
                except:
                    # 如果獲取失敗，移除卷/冊等前綴
                    book_name = re.sub(r'^(卷[一二三四五六七八九十百]+|第\d+[册冊]|第[一二三四五六七八九十]+[册冊])\s+', '', book_name)
        
        return book_name
    return None

def extract_volume_name(soup):
    """
    從網頁中提取卷名（用於三級結構）
    從 <a class="brandtitle"> 元素中提取，只有當包含卷/冊關鍵字時才返回
    """
    brandtitle = soup.find('a', class_='brandtitle')
    if brandtitle:
        volume_name = brandtitle.get_text(strip=True)
        # 如果包含卷/冊關鍵字，這就是卷名
        if re.search(r'(卷[一二三四五六七八九十百]+|第\d+[册冊]|第[一二三四五六七八九十]+[册冊])', volume_name):
            return volume_name
    return None

def extract_chapter_number(url):
    """
    從 URL 提取章號（只取最後一個數字）
    例如：從 "https://ezoe.work/books/2/2022-1-1.html" 提取 "1"（只要章號）
    例如：從 "https://ezoe.work/books/1/1046-1.html" 提取 "1"
    """
    # 三級結構：2022-1-1.html → "1"（返回第二個數字，即章號）
    match = re.search(r'-(\d+)-(\d+)\.html', url)
    if match:
        return match.group(2)  # 返回章號
    
    # 兩級結構：1046-1.html → "1"
    match = re.search(r'-(\d+)\.html', url)
    if match:
        return match.group(1)
    return None

def extract_article_number_from_url(url):
    """
    從 URL 中提取篇數（保留用於向後兼容）
    例如：從 "https://ezoe.work/books/3/3061-11.html" 提取 "11"
    例如：從 "https://ezoe.work/books/2/2022-1-1.html" 提取 "1-1"（三級結構）
    """
    # 先嘗試匹配三級結構（X-Y-Z.html）
    match = re.search(r'-(\d+)-(\d+)\.html', url)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    
    # 再匹配兩級結構（X-Y.html）
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
        book_name = extract_book_name(soup, url)
        volume_name = extract_volume_name(soup)
        chapter_num = extract_chapter_number(url)
        
        # 三級結構：書名_卷名_章號.md
        if book_name and volume_name and chapter_num:
            output_file = f"{book_name}_{volume_name}_{chapter_num}.md"
        # 兩級結構：書名_章號.md
        elif book_name and chapter_num:
            output_file = f"{book_name}_{chapter_num}.md"
        elif book_name:
            output_file = f"{book_name}.md"
        elif chapter_num:
            output_file = f"第{chapter_num}篇.md"
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
    
    # 2. 提取主要內容區域（使用多層級備用選擇器）
    main = None
    selectors = [
        ('div', {'class': 'main'}),
        ('div', {'class': 'content'}),
        ('div', {'class': 'container'}),
        ('article', {}),
        ('main', {}),
        ('body', {})  # 最後備用方案
    ]
    
    for tag_name, attrs in selectors:
        main = soup.find(tag_name, attrs) if attrs else soup.find(tag_name)
        if main:
            selector_desc = f"<{tag_name} class='{attrs.get('class')}'>" if attrs.get('class') else f"<{tag_name}>"
            if tag_name != 'div' or attrs.get('class') != 'main':
                print(f"資訊：使用備用選擇器 {selector_desc}")
            break
    
    if not main:
        print("警告：找不到主要內容區域")
        print("除錯資訊：網頁的前幾層標籤：")
        for tag in soup.find_all(limit=10):
            print(f"  <{tag.name}> class={tag.get('class')}, id={tag.get('id')}")
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
