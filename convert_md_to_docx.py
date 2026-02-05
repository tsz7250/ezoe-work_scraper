#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown 轉 DOCX/PDF 合併轉換工具
將所有 Markdown 檔案合併並轉換為 DOCX 和 PDF 檔案
"""

import os
import re
import glob
import sys
import traceback
import pypandoc
from docx import Document
from docx.shared import Pt
from docx2pdf import convert as docx2pdf_convert

# Word WdTCSCConverterDirection 常數
# SCTC = 0 = Simplified to Traditional Chinese（簡體→繁體）
# TCSC = 1 = Traditional to Simplified Chinese（繁體→簡體）
# Auto = 2 = 自動判斷
WD_TCSC_CONVERTER_DIRECTION_SCTC = 0

def get_article_number(filename):
    """從檔案名提取文章編號"""
    match = re.search(r'_(\d+)\.md$', filename)
    if match:
        return int(match.group(1))
    return 0

def extract_book_name(filename):
    """從檔案名提取書名（去除 _數字.md 後綴）"""
    match = re.search(r'^(.+?)_\d+\.md$', filename)
    if match:
        return match.group(1)
    return None

def collect_markdown_files_by_book():
    """收集所有 Markdown 檔案並按書名分組"""
    # 找出所有 *_數字.md 格式的檔案
    all_files = glob.glob("*_*.md")
    
    # 過濾出符合 書名_數字.md 格式的檔案
    valid_files = [f for f in all_files if re.match(r'^.+_\d+\.md$', f)]
    
    if not valid_files:
        return {}
    
    # 按書名分組
    books = {}
    for file in valid_files:
        book_name = extract_book_name(file)
        if book_name:
            if book_name not in books:
                books[book_name] = []
            books[book_name].append(file)
    
    # 每本書的檔案按編號排序
    for book_name in books:
        books[book_name].sort(key=get_article_number)
    
    return books

def process_markdown_headings(content, article_num):
    """處理 Markdown 標題：
    將連續的一級標題合併為一個「第X篇 XXXX」格式的標題
    保留其他級別標題的原有格式（保持視覺差異）
    """
    lines = content.split('\n')
    processed_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # 處理一級標題
        if line.strip().startswith('# '):
            # 收集所有連續的一級標題（包括中間有空行的情況）
            h1_titles = []
            j = i
            
            # 收集連續的一級標題
            while j < len(lines):
                current_line = lines[j].strip()
                if current_line.startswith('# '):
                    title_text = current_line[2:].strip()
                    h1_titles.append(title_text)
                    j += 1
                elif current_line == '':
                    # 空行，繼續檢查下一行是否是一級標題
                    j += 1
                    continue
                else:
                    # 遇到非空行且非一級標題，停止收集
                    break
            
            # 處理收集到的一級標題
            if len(h1_titles) == 0:
                # 不應該發生，但安全處理
                processed_lines.append(line)
                i += 1
            elif len(h1_titles) == 1:
                # 只有一個標題
                title_text = h1_titles[0]
                # 檢查是否已經包含「第X篇」格式（中文數字或阿拉伯數字）
                if re.match(r'^第[一二三四五六七八九十百千萬〇\d]+篇\s+.+', title_text):
                    # 已經包含「第X篇 標題」，直接使用（保留原本的中文數字）
                    processed_lines.append(f"# {title_text}")
                elif re.match(r'^第[一二三四五六七八九十百千萬〇\d]+篇$', title_text):
                    # 只有「第X篇」，沒有標題內容，直接使用（保留原本的中文數字）
                    processed_lines.append(f"# {title_text}")
                else:
                    # 不包含篇數，添加篇數（使用檔案編號轉換為中文數字，但這裡先簡單處理）
                    # 如果沒有篇數，使用檔案編號
                    processed_lines.append(f"# 第{article_num}篇 {title_text}")
            else:
                # 多個標題，合併處理
                # 第一個標題通常是篇數（如「第四〇四篇」或「第一篇」），後續是標題內容
                first_title = h1_titles[0]
                other_titles = ' '.join(h1_titles[1:]) if len(h1_titles) > 1 else ''
                
                # 提取第一個標題中的篇數部分（保留中文數字）
                article_match = re.match(r'^(第[一二三四五六七八九十百千萬〇\d]+篇)', first_title)
                if article_match:
                    # 找到篇數部分，保留原本的中文數字格式
                    article_part = article_match.group(1)
                    # 檢查第一個標題是否只包含篇數（沒有其他內容）
                    if re.match(r'^第[一二三四五六七八九十百千萬〇\d]+篇$', first_title):
                        # 第一個標題只包含篇數，使用這個篇數，加上後續標題
                        if other_titles:
                            processed_lines.append(f"# {article_part} {other_titles}")
                        else:
                            processed_lines.append(f"# {article_part}")
                    else:
                        # 第一個標題包含篇數和內容，提取內容並與後續標題合併
                        first_content = re.sub(r'^第[一二三四五六七八九十百千萬〇\d]+篇\s+', '', first_title)
                        all_content = ' '.join([first_content] + h1_titles[1:]) if first_content else other_titles
                        if all_content:
                            processed_lines.append(f"# {article_part} {all_content}")
                        else:
                            processed_lines.append(f"# {article_part}")
                else:
                    # 第一個標題不包含篇數，合併所有標題並添加篇數（使用檔案編號）
                    all_titles = ' '.join(h1_titles)
                    processed_lines.append(f"# 第{article_num}篇 {all_titles}")
            
            # 跳過已處理的行（包括空行）
            i = j
        else:
            # 保留其他所有內容（包括其他級別標題）不變
            processed_lines.append(line)
            i += 1
    
    return '\n'.join(processed_lines)

def merge_markdown_files(files):
    """合併所有 Markdown 檔案內容，每篇之間添加分頁符號"""
    merged_content = []
    
    for i, file in enumerate(files):
        print(f"讀取檔案 [{i+1}/{len(files)}]: {file}")
        
        try:
            with open(file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                
                # 將波浪號轉義，避免被 pandoc 誤解為下標標記
                # 在 Markdown 中，~text~ 會被轉換為下標
                # 使用反斜線轉義波浪號：\~
                content = content.replace('~', r'\~')
                
                # 獲取文章編號
                article_num = get_article_number(file)
                
                # 處理標題：只保留一級標題作為標題，其他轉為普通段落
                content = process_markdown_headings(content, article_num)
                
                merged_content.append(content)
                
                # 在每篇文章之間添加分頁符號（除了最後一篇）
                if i < len(files) - 1:
                    # 使用 pandoc 的原始 OpenXML 區塊插入 Word 分頁符
                    merged_content.append('\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n')
        
        except Exception as e:
            print(f"警告：無法讀取檔案 {file}: {e}")
            continue
    
    return '\n\n'.join(merged_content)

def set_line_spacing(docx_file, line_spacing=1.5):
    """設定 DOCX 檔案的行距"""
    try:
        print(f"設定行距為 {line_spacing}...")
        doc = Document(docx_file)
        
        # 遍歷所有段落並設定行距
        for paragraph in doc.paragraphs:
            paragraph.paragraph_format.line_spacing = line_spacing
        
        # 遍歷所有表格中的段落並設定行距
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.line_spacing = line_spacing
        
        # 儲存修改
        doc.save(docx_file)
        print(f"行距設定完成")
        return True
        
    except Exception as e:
        print(f"\n警告：設定行距時發生錯誤: {e}")
        return False

def process_docx_headings(docx_file):
    """處理 DOCX 檔案中的標題樣式：
    只保留一級標題（Heading 1）作為標題樣式，其他級別的標題改為自訂格式
    保持視覺差異（字體大小、粗體等），但不會成為 PDF 書籤
    """
    try:
        print("處理標題樣式，確保只有一級標題用於 PDF 書籤...")
        doc = Document(docx_file)
        
        def get_heading_format(style_name):
            """從標題樣式中獲取格式資訊（字體大小、是否粗體）"""
            try:
                style = doc.styles[style_name]
                # 獲取字體大小（如果有的話）
                font_size = None
                is_bold = False
                
                if hasattr(style, 'font') and style.font.size:
                    font_size = style.font.size
                if hasattr(style, 'font') and style.font.bold is not None:
                    is_bold = style.font.bold
                
                # 如果沒有從樣式獲取到，使用預設值
                if font_size is None:
                    # 根據標題級別設定預設字體大小
                    if '2' in style_name or '二' in style_name:
                        font_size = Pt(18)
                    elif '3' in style_name or '三' in style_name:
                        font_size = Pt(16)
                    elif '4' in style_name or '四' in style_name:
                        font_size = Pt(14)
                    elif '5' in style_name or '五' in style_name:
                        font_size = Pt(12)
                    else:
                        font_size = Pt(14)  # 預設值
                
                # 預設標題為粗體
                if is_bold is False and style_name.startswith(('Heading', '標題')):
                    is_bold = True
                
                return font_size, is_bold
            except:
                # 如果無法獲取樣式，使用預設值
                return Pt(14), True
        
        def apply_custom_format(paragraph, original_style_name):
            """將段落改為 Normal 樣式，但保留原有的視覺格式"""
            # 先獲取原樣式的格式資訊
            font_size, is_bold = get_heading_format(original_style_name)
            
            # 改為 Normal 樣式（不會成為 PDF 書籤）
            paragraph.style = doc.styles['Normal']
            
            # 應用原樣式的格式到所有 runs
            for run in paragraph.runs:
                if font_size:
                    run.font.size = font_size
                run.font.bold = is_bold
        
        # 需要處理的非一級標題樣式
        heading_styles_to_process = [
            'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5',
            '標題 2', '標題 3', '標題 4', '標題 5'
        ]
        
        # 遍歷所有段落
        for paragraph in doc.paragraphs:
            style_name = paragraph.style.name if paragraph.style else None
            
            # 如果段落使用的是非一級標題樣式，改為自訂格式
            if style_name in heading_styles_to_process:
                apply_custom_format(paragraph, style_name)
        
        # 遍歷所有表格中的段落
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        style_name = paragraph.style.name if paragraph.style else None
                        if style_name in heading_styles_to_process:
                            apply_custom_format(paragraph, style_name)
        
        # 儲存修改
        doc.save(docx_file)
        print("標題樣式處理完成")
        return True
        
    except Exception as e:
        print(f"\n警告：處理標題樣式時發生錯誤: {e}")
        traceback.print_exc()
        return False

def convert_to_docx(markdown_content, output_file):
    """使用 pypandoc 將 Markdown 轉換為 DOCX"""
    try:
        # 檢查 pandoc 是否已安裝
        try:
            pypandoc.get_pandoc_version()
        except OSError:
            print("未檢測到 Pandoc，正在下載並安裝...")
            pypandoc.download_pandoc()
        
        print(f"\n開始轉換為 DOCX 格式...")
        
        # 先將內容寫入臨時 Markdown 檔案，避免編碼問題
        temp_md = "temp_merged.md"
        with open(temp_md, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        # 設定轉換選項
        extra_args = [
            '--standalone',  # 產生獨立的文件
        ]
        
        # 執行轉換（從檔案轉換）
        pypandoc.convert_file(
            temp_md,
            'docx',
            outputfile=output_file,
            extra_args=extra_args
        )
        
        # 刪除臨時檔案
        if os.path.exists(temp_md):
            os.remove(temp_md)
        
        print(f"\n成功生成 DOCX 檔案：{output_file}")
        print(f"檔案大小：{os.path.getsize(output_file) / 1024:.2f} KB")
        
        # 設定行距為 1.5
        set_line_spacing(output_file, 1.5)
        
        # 處理標題樣式：只保留一級標題作為標題樣式
        process_docx_headings(output_file)
        
        return True
        
    except Exception as e:
        print(f"\n錯誤：轉換過程發生錯誤: {e}")
        return False

def convert_docx_to_pdf(docx_file):
    """將 DOCX 檔案轉換為 PDF"""
    try:
        pdf_file = docx_file.replace('.docx', '.pdf')
        print(f"\n開始轉換為 PDF 格式：{pdf_file}")
        
        # 執行轉換
        docx2pdf_convert(docx_file, pdf_file)
        
        print(f"成功生成 PDF 檔案：{pdf_file}")
        print(f"檔案大小：{os.path.getsize(pdf_file) / 1024:.2f} KB")
        
        return True
        
    except Exception as e:
        print(f"\n警告：轉換 PDF 時發生錯誤: {e}")
        traceback.print_exc()
        return False

def convert_string_simplified_to_traditional(text):
    """
    使用 Word COM 將字串由簡體轉為繁體。
    需在 Windows 上執行且已安裝 Microsoft Word。
    回傳轉換後的字串；若失敗則回傳原字串。
    """
    if sys.platform != 'win32':
        return text
    try:
        import win32com.client
        app = win32com.client.Dispatch('Word.Application')
        app.Visible = False
        doc = app.Documents.Add()
        doc.Content.Text = text
        doc.Content.TCSCConverter(WD_TCSC_CONVERTER_DIRECTION_SCTC, True, True)
        result = doc.Content.Text.strip()
        doc.Close(SaveChanges=False)
        app.Quit()
        return result
    except Exception as e:
        print(f"警告：字串簡轉繁失敗: {e}")
        return text

def convert_docx_simplified_to_traditional(docx_path):
    """
    使用 Word COM 將 DOCX 整份內容由簡體轉為繁體並存檔。
    需在 Windows 上執行且已安裝 Microsoft Word。
    回傳 True 成功，False 失敗。
    """
    if sys.platform != 'win32':
        print("警告：簡轉繁需在 Windows 上使用 Word COM，已略過")
        return False
    try:
        import win32com.client
        abs_path = os.path.abspath(docx_path)
        if not os.path.isfile(abs_path):
            print(f"警告：找不到檔案 {abs_path}")
            return False
        app = win32com.client.Dispatch('Word.Application')
        app.Visible = False
        # 使用完整路徑開啟，避免中文路徑問題
        doc = app.Documents.Open(FileName=abs_path, ConfirmConversions=False, ReadOnly=False)
        doc.Activate()
        # 選取整份文件後再轉換（與 Word 手動「簡體轉繁體」行為一致）
        app.Selection.WholeStory()
        app.Selection.Range.TCSCConverter(WD_TCSC_CONVERTER_DIRECTION_SCTC, True, True)
        doc.Save()
        doc.Close(SaveChanges=True)
        app.Quit()
        print("已使用 Word 將整份文件轉為繁體並存檔")
        return True
    except Exception as e:
        print(f"\n警告：Word 簡轉繁時發生錯誤: {e}")
        traceback.print_exc()
        return False

def main():
    """主程式"""
    print("=" * 60)
    print("Markdown 轉 DOCX/PDF 合併轉換工具")
    print("=" * 60)
    
    # 1. 收集 Markdown 檔案並按書名分組
    print("\n步驟 1: 收集 Markdown 檔案...")
    books = collect_markdown_files_by_book()
    
    if not books:
        print("錯誤：未找到任何符合 書名_數字.md 格式的 Markdown 檔案")
        return 1
    
    print(f"\n找到 {len(books)} 本書，共 {sum(len(files) for files in books.values())} 個檔案：")
    for book_name, files in books.items():
        print(f"\n【{book_name}】：{len(files)} 篇")
        for i, file in enumerate(files, 1):
            article_num = get_article_number(file)
            print(f"  {i}. {file} (第 {article_num} 篇)")
    
    # 2. 處理每本書
    print("\n" + "=" * 60)
    success_count = 0
    fail_count = 0
    
    for book_idx, (book_name, files) in enumerate(books.items(), 1):
        print(f"\n[{book_idx}/{len(books)}] 處理：{book_name}")
        print("-" * 60)
        
        # 合併檔案內容
        print(f"步驟 2: 合併 {len(files)} 個 Markdown 檔案...")
        merged_content = merge_markdown_files(files)
        
        if not merged_content:
            print(f"警告：無法讀取 {book_name} 的內容")
            fail_count += 1
            continue
        
        print(f"合併完成，總字數：{len(merged_content)} 字元")
        
        # 轉換為 DOCX
        print(f"步驟 3: 轉換為 DOCX 格式...")
        output_file = f"{book_name}.docx"
        
        if convert_to_docx(merged_content, output_file):
            # 使用 Word COM 將 DOCX 簡體轉繁體並存檔
            print(f"步驟 4: 使用 Word 將 DOCX 簡體轉繁體...")
            convert_docx_simplified_to_traditional(output_file)
            
            # 將書名轉為繁體並重新命名檔案
            print(f"步驟 5: 將檔案名稱轉為繁體...")
            traditional_book_name = convert_string_simplified_to_traditional(book_name)
            if traditional_book_name != book_name:
                new_output_file = f"{traditional_book_name}.docx"
                try:
                    if os.path.exists(new_output_file):
                        os.remove(new_output_file)
                    os.rename(output_file, new_output_file)
                    output_file = new_output_file
                    print(f"已將檔案重新命名為：{output_file}")
                except Exception as e:
                    print(f"警告：重新命名失敗: {e}")
            
            # PDF 由已轉繁的 .docx 產生
            print(f"步驟 6: 轉換為 PDF 格式...")
            convert_docx_to_pdf(output_file)  # PDF 轉換失敗不影響整體成功狀態
            success_count += 1
        else:
            fail_count += 1
    
    # 總結
    print("\n" + "=" * 60)
    print("轉換完成！")
    print(f"成功：{success_count} 本，失敗：{fail_count} 本")
    print("\n已生成 DOCX 和 PDF 檔案")
    print("=" * 60)
    
    return 0 if fail_count == 0 else 1

if __name__ == '__main__':
    sys.exit(main())
