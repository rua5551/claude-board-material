import os
import re
import sys

def find_docx_files(folder_path):
    """查找文件夹中所有.docx文件"""
    docx_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.docx'):
                docx_files.append(os.path.join(root, file))
    return docx_files

def replace_in_docx(file_path, meeting_num, month, day):
    """替换docx文件内容"""
    try:
        from docx import Document

        doc = Document(file_path)
        modified = False

        def replace_text_in_element(element, meeting_num, month, day):
            """递归替换元素中的文本，保留格式"""
            if hasattr(element, 'text') and hasattr(element, 'runs'):
                # 处理带runs的元素（paragraph）
                if element.runs:
                    for run in element.runs:
                        original = run.text
                        if original:
                            new_text = original
                            new_text = re.sub(r'第(\d+)次', f'第{meeting_num}次', new_text)
                            new_text = re.sub(r'2026年(\d+)月(\d+)日', f'2026年{month}月{day}日', new_text)
                            if new_text != original:
                                run.text = new_text
                                return True
                return False

        # 替换段落中的文本
        for paragraph in doc.paragraphs:
            if replace_text_in_element(paragraph, meeting_num, month, day):
                modified = True

        # 替换表格中的文本
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if replace_text_in_element(cell, meeting_num, month, day):
                        modified = True

        if modified:
            doc.save(file_path)
        return modified

    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return False

def replace_in_filename(file_path, meeting_num, month, day):
    """替换文件名"""
    filename = os.path.basename(file_path)
    directory = os.path.dirname(file_path)

    new_filename = filename

    # 替换文件名中的"第n次"
    new_filename = re.sub(r'第(\d+)次', f'第{meeting_num}次', new_filename)

    # 替换文件名中的"2026年n月n日"
    new_filename = re.sub(r'2026年(\d+)月(\d+)日', f'2026年{month}月{day}日', new_filename)

    if new_filename != filename:
        new_path = os.path.join(directory, new_filename)
        if not os.path.exists(new_path):
            os.rename(file_path, new_path)
            return True
        else:
            print(f"目标文件名已存在: {new_path}")
            return False
    return False

def main():
    print("=" * 50)
    print("Word文档批量替换工具")
    print("=" * 50)

    # 1. 输入文件夹路径
    folder_path = input("请输入文件夹路径：").strip()
    folder_path = folder_path.strip('"').strip("'")

    if not os.path.isdir(folder_path):
        print("错误：指定的路径不是有效的文件夹！")
        input("按回车键退出...")
        return

    # 2. 输入第几次会议
    meeting_num = input("请输入第几次会议（数字）：").strip()
    if not meeting_num.isdigit():
        print("错误：请输入有效的数字！")
        input("按回车键退出...")
        return

    # 3. 输入月
    month = input("请输入月（数字）：").strip()
    if not month.isdigit() or not (1 <= int(month) <= 12):
        print("错误：请输入1-12之间的数字！")
        input("按回车键退出...")
        return

    # 4. 输入日
    day = input("请输入日（数字）：").strip()
    if not day.isdigit() or not (1 <= int(day) <= 31):
        print("错误：请输入1-31之间的数字！")
        input("回车键退出...")
        return

    print("\n开始处理...")
    print(f"文件夹: {folder_path}")
    print(f"会议次数: 第{meeting_num}次")
    print(f"日期: 2026年{month}月{day}日")
    print("-" * 50)

    # 查找所有docx文件
    docx_files = find_docx_files(folder_path)

    if not docx_files:
        print("未找到任何.docx文件！")
        input("按回车键退出...")
        return

    print(f"找到 {len(docx_files)} 个.docx文件\n")

    # 处理每个文件
    content_modified_count = 0
    filename_modified_count = 0

    for file_path in docx_files:
        filename = os.path.basename(file_path)
        print(f"处理: {filename}")

        # 替换内容
        if replace_in_docx(file_path, meeting_num, month, day):
            content_modified_count += 1
            print(f"  - 内容已替换")

        # 替换文件名
        if replace_in_filename(file_path, meeting_num, month, day):
            filename_modified_count += 1
            print(f"  - 文件名已修改")

    print("-" * 50)
    print(f"处理完成！")
    print(f"- 内容替换: {content_modified_count} 个文件")
    print(f"- 文件名修改: {filename_modified_count} 个文件")
    print("\n按回车键退出...")
    input()

if __name__ == "__main__":
    main()