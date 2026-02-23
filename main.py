import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import os
import shutil
import pathlib
import sys
import re

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

root = tk.Tk()
root.geometry("1200x800")
root.title("Luffa ChessBase Strings (修复版)")

# 图标（增加异常处理）
try:
    root.iconbitmap(resource_path("s.ico"))
except Exception as e:
    print(f"图标加载失败：{e}")

# 全局变量（统一管理）
original_stdout = sys.stdout
root_path = None
EN_DIR = "English".lower()
CN_DIR = "Chinese".lower()
CN_NEW_DIR = "Chinese_new"
done_init_file_tree = False
nb1 = None
created_tabs = {}  # {文件名：(frame3, treeview)}
current_tree = None
current_file = None
current_selected_index = -1
current_selected_iid = None
search_results = []  # 存储当前搜索结果的 iid 列表
current_search_index = -1  # 当前高亮的搜索结果索引
base_font_size = 14  # 基础字号，用于全局缩放


# ==================== 【修复核心】保存翻译到文件 ====================
def save_translation_to_file(event=None, force_iid=None):
    """
    ✅ 修复：使用精确匹配，不再用 startswith() 导致前缀误匹配
    ✅ 修复：只替换第一个匹配项，避免重复写入
    ✅ 修复：保留原始空格格式
    """
    global current_tree, current_file, current_selected_iid
    if not current_tree or not current_file or not root_path:
        return False

    new_translation = text1.get().strip()
    if not new_translation:
        return False

    target_iid = force_iid if force_iid else current_selected_iid
    if not target_iid or target_iid not in current_tree.get_children():
        return False

    item_data = current_tree.item(target_iid)
    cap_word = item_data["values"][0].strip()  # 去除可能的空格
    english_text = item_data["values"][1]

    file_path = os.path.join(root_path, CN_NEW_DIR, current_file)
    if not os.path.exists(file_path):
        messagebox.showerror("错误", f"文件不存在：{file_path}")
        return False

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        saved = False
        with open(file_path, "w", encoding="utf-8") as f:
            for line in lines:
                # ✅ 使用正则精确匹配：确保逗号前完全等于 cap_word
                # 匹配模式：^M_TXTR_METAL,  或 ^M_TXTR_METAL   ,
                pattern = rf'^{re.escape(cap_word)}\s*,'
                if re.match(pattern, line) and line.strip().endswith('";'):
                    if not saved:  # ✅ 只替换第一个匹配，避免重复
                        parts = line.split(',', 1)
                        if len(parts) >= 2:
                            spaces = parts[1].split('"')[0]  # 保留原始空格
                            new_line = f'{cap_word},{spaces}"{new_translation}";\n'
                            f.write(new_line)
                            saved = True
                            continue
                f.write(line)

        if saved:
            current_tree.item(target_iid, values=(cap_word, english_text, new_translation))
            return True
        return False
    except Exception as e:
        messagebox.showerror("错误", f"保存翻译失败：{e}")
        return False


# ==================== Tree 选择事件 ====================
def on_tree_select(event):
    global current_tree, current_file, current_selected_index, current_selected_iid

    # 🔧 修复：先保存当前行的翻译（无论焦点在哪里）
    save_translation_to_file(force_iid=current_selected_iid)

    tree = event.widget
    current_tree = tree
    selected_items = tree.selection()
    if not selected_items:
        current_selected_iid = None
        current_selected_index = -1
        text1.delete(0, tk.END)
        return

    iid = selected_items[0]
    current_selected_iid = iid
    children = tree.get_children()
    current_selected_index = children.index(iid)

    item_data = tree.item(iid)
    cn_translation = item_data["values"][2]
    text1.delete(0, tk.END)
    text1.insert(0, cn_translation)


# ==================== Notebook 切换事件 ====================
def on_tab_changed(event):
    global current_tree, current_file, current_selected_iid, search_results, current_search_index
    # 🔧 修复：切换前强制保存
    save_translation_to_file(force_iid=current_selected_iid)

    if not nb1:
        return

    current_tab_id = nb1.select()
    if not current_tab_id:
        return

    current_tree = None
    current_file = None
    current_selected_iid = None
    current_selected_index = -1
    # 🔧 修复：清空搜索结果
    search_results = []
    current_search_index = -1
    
    for file_name, (frame3, treeview) in created_tabs.items():
        if str(frame3) == current_tab_id:
            current_file = file_name
            current_tree = treeview
            break

    text1.delete(0, tk.END)


# ==================== 上一条 / 下一条 ====================
def select_prev_row():
    global current_tree, current_selected_index, current_selected_iid
    save_translation_to_file(force_iid=current_selected_iid)
    if not current_tree or current_selected_index <= 0:
        return
    current_selected_index -= 1
    children = current_tree.get_children()
    target_iid = children[current_selected_index]
    current_selected_iid = target_iid
    current_tree.selection_set(target_iid)
    current_tree.see(target_iid)
    item_data = current_tree.item(target_iid)
    text1.delete(0, tk.END)
    text1.insert(0, item_data["values"][2])


def select_next_row():
    global current_tree, current_selected_index, current_selected_iid
    save_translation_to_file(force_iid=current_selected_iid)
    if not current_tree:
        return
    children = current_tree.get_children()
    if current_selected_index >= len(children) - 1:
        return
    current_selected_index += 1
    target_iid = children[current_selected_index]
    current_selected_iid = target_iid
    current_tree.selection_set(target_iid)
    current_tree.see(target_iid)
    item_data = current_tree.item(target_iid)
    text1.delete(0, tk.END)
    text1.insert(0, item_data["values"][2])


# ==================== 复制英文 ====================
def copy_english_to_clipboard():
    global current_tree, current_selected_iid
    if not current_tree or not current_selected_iid:
        return
    if current_selected_iid not in current_tree.get_children():
        return
    item_data = current_tree.item(current_selected_iid)
    english_text = item_data["values"][1]
    root.clipboard_clear()
    root.clipboard_append(english_text)
    root.update()


# ==================== 搜索功能 ====================
def search_in_tree():
    global current_tree, search_results, current_search_index, current_selected_iid
    # 🔧 修复：搜索前先保存当前行
    save_translation_to_file(force_iid=current_selected_iid)
    
    if not current_tree:
        messagebox.showinfo("提示", "请先打开一个文件标签页")
        return
    keyword = search_entry.get().strip()
    if not keyword:
        messagebox.showinfo("提示", "请输入搜索关键词")
        return

    # 🔧 修复：清空之前的搜索结果
    search_results = []
    current_search_index = -1

    children = current_tree.get_children()
    for iid in children:
        item_data = current_tree.item(iid)
        values = item_data["values"]
        if any(keyword.lower() in str(val).lower() for val in values):
            search_results.append(iid)

    if not search_results:
        messagebox.showinfo("提示", f"在当前文件 '{current_file}' 中未找到匹配内容：'{keyword}'")
        return

    current_search_index = 0
    highlight_search_result()


def goto_prev_search():
    global current_search_index, current_selected_iid
    # 🔧 修复：移动前先保存当前行
    save_translation_to_file(force_iid=current_selected_iid)
    
    if not search_results:
        return
    if current_search_index > 0:
        current_search_index -= 1
    else:
        current_search_index = len(search_results) - 1
    highlight_search_result()


def goto_next_search():
    global current_search_index, current_selected_iid
    # 🔧 修复：移动前先保存当前行
    save_translation_to_file(force_iid=current_selected_iid)
    
    if not search_results:
        return
    if current_search_index < len(search_results) - 1:
        current_search_index += 1
    else:
        current_search_index = 0
    highlight_search_result()


def highlight_search_result():
    """根据 current_search_index 高亮对应的行"""
    global current_tree, current_selected_iid, current_selected_index
    
    if not current_tree or not (0 <= current_search_index < len(search_results)):
        if current_tree:
            current_tree.selection_remove(current_tree.selection())
        current_selected_iid = None
        current_selected_index = -1
        text1.delete(0, tk.END)
        return
        
    target_iid = search_results[current_search_index]
    current_selected_iid = target_iid
    children = current_tree.get_children()
    current_selected_index = children.index(target_iid)

    current_tree.selection_set(target_iid)
    current_tree.see(target_iid)

    item_data = current_tree.item(target_iid)
    text1.delete(0, tk.END)
    text1.insert(0, item_data["values"][2])


# ==================== 原有工具函数 ====================
def generate_tip():
    return messagebox.askokcancel(
        parent=root,
        title="打开前注意",
        message="仅适用于 CB26 和 Fritz20 及以下，未来更高版本未测试。\n先选择 Messages 目录\n请不要在原目录打开 Messages，建议复制"
    )


def get_msgs_dir():
    if generate_tip():
        return filedialog.askdirectory(title="打开 Messages 目录")


def set_root_path():
    global root_path
    root_path = get_msgs_dir()


# ==================== 【修复核心】提取枚举名列表 ====================
def genCapWordsList(file_name, language):
    """
    ✅ 修复：使用正则精确提取枚举名，不用 split(',') 避免空格问题
    """
    if not root_path:
        messagebox.showerror("错误", "请先选择 Messages 目录！")
        return []
    try:
        file_path = os.path.join(root_path, language, file_name)
        with open(file_path, mode="r", encoding='utf-8') as file:
            cap_words_list = []
            for line in file:
                line_stripped = line.strip()
                if line_stripped.startswith('M_') and line_stripped.endswith('";'):
                    # ✅ 使用正则精确匹配：M_开头，逗号前结束
                    match = re.match(r'^(M_\w+)\s*,', line_stripped)
                    if match:
                        cap_words_list.append(match.group(1))
            return cap_words_list
    except FileNotFoundError:
        messagebox.showerror("错误", f"文件不存在：{file_path}")
        return []
    except Exception as e:
        messagebox.showerror("错误", f"读取文件失败：{e}")
        return []


# ==================== 【修复核心】提取语言字典 ====================
def genLangDict(file_name, language):
    """
    ✅ 修复：使用正则精确提取，不用 split(',') 避免字符串内含逗号时出错
    ✅ 修复：自动去除枚举名和翻译文本的首尾空格
    """
    if not root_path:
        messagebox.showerror("错误", "请先选择 Messages 目录！")
        return {}
    try:
        file_path = os.path.join(root_path, language, file_name)
        with open(file_path, mode="r", encoding='utf-8') as file:
            lang_dict = {}
            for line in file:
                line_stripped = line.strip()
                if line_stripped.startswith('M_') and line_stripped.endswith('";'):
                    # ✅ 使用正则精确匹配整个结构
                    match = re.match(r'^(M_\w+)\s*,\s*"([^"]*)"\s*;', line_stripped)
                    if match:
                        cap_words = match.group(1)
                        translation = match.group(2)
                        lang_dict[cap_words] = translation
            return lang_dict
    except FileNotFoundError:
        messagebox.showerror("错误", f"文件不存在：{file_path}")
        return {}
    except Exception as e:
        messagebox.showerror("错误", f"读取文件失败：{e}")
        return {}


def isCnEnEqual(file_name):
    en_list = genCapWordsList(file_name, EN_DIR)
    cn_list = genCapWordsList(file_name, CN_DIR)
    return en_list == cn_list


def getMaxLen(file_name):
    cap_list = genCapWordsList(file_name, EN_DIR)
    if not cap_list:
        return 20
    max_len = len(sorted(cap_list, key=len, reverse=True)[0]) + 4
    return max_len


def writeHeaders(file_name, create_time="yyyy/mm/dd", author="yourName"):
    dir_path = os.path.join(root_path, CN_NEW_DIR)
    os.makedirs(dir_path, exist_ok=True)
    try:
        with open(os.path.join(dir_path, file_name), mode='w', encoding='utf-8') as cn_new_file:
            uni_headers = f'''/*
    Generated with Luffa ChessBase Strings
    版本：0.1 (修复版)
    创建日期：{create_time}
    作者：{author}
    IMPORTANT: 
    0. 不要在";后面再加任何注释//等字符（空白字符没关系）
    1. Please do not change or translate the symbols in capital letters.
    2. Do not remove or add commas and semicolons.
    3. Make sure that every string is included into quotes.
*/

'''
            cn_new_file.write(uni_headers)
    except Exception as e:
        messagebox.showerror("错误", f"写入头部失败：{e}")


# ==================== 【修复核心】写入主要内容 ====================
def writeMainContents(file_name):
    """
    ✅ 修复：读取英文原始文件，逐行保留原始空格格式
    ✅ 修复：不再用固定公式计算空格，而是保留原始行的空格
    """
    if not root_path:
        return
    dir_path = os.path.join(root_path, CN_NEW_DIR)
    try:
        # ✅ 读取英文原始文件，保留格式
        en_file_path = os.path.join(root_path, EN_DIR, file_name)
        cn_new_file_path = os.path.join(dir_path, file_name)
        
        en_dict = genLangDict(file_name, EN_DIR)
        cn_dict = genLangDict(file_name, CN_DIR)

        with open(en_file_path, 'r', encoding='utf-8') as en_file, \
             open(cn_new_file_path, 'w', encoding='utf-8') as cn_new_file:
            
            for line in en_file:
                # ✅ 精确匹配每一行，保留原始空格
                match = re.match(r'^(M_\w+)(\s*,\s*)(")([^"]*)(";.*)$', line)
                if match:
                    cap_word = match.group(1)
                    spaces_and_comma = match.group(2)
                    quote_open = match.group(3)
                    en_text = match.group(4)
                    quote_close_and_end = match.group(5)
                    
                    # 查中文翻译，如果没有则用英文
                    cn_text = cn_dict.get(cap_word, en_text)
                    
                    new_line = f'{cap_word}{spaces_and_comma}{quote_open}{cn_text}{quote_close_and_end}\n'
                    cn_new_file.write(new_line)
                else:
                    cn_new_file.write(line)  # 非数据行原样保留（如注释、空行）
    except Exception as e:
        messagebox.showerror("错误", f"写入内容失败：{e}")


def compare_origin_en_cn(file_name):
    cn_file_path = os.path.join(root_path, CN_DIR, file_name)
    en_file_path = os.path.join(root_path, EN_DIR, file_name)
    if not os.path.exists(cn_file_path):
        print(f"中文文件夹里缺失{file_name}。将把英文对应 strings 复制到中文文件夹里。")
        try:
            shutil.copy2(en_file_path, cn_file_path)
            return True
        except Exception as e:
            print(f"复制失败：{e}")
            return False
    else:
        return isCnEnEqual(file_name)


def generateNewCn(file_name):
    if not compare_origin_en_cn(file_name):
        print(f"*{file_name}中文对比英文不同，将会生成新的{file_name}中文 strings")
        writeHeaders(file_name, create_time="yyyy/mm/dd", author="854982774@qq.com")
        writeMainContents(file_name)
        print(f"新建 Chinese_new {file_name}成功\n")
    else:
        print(f"{file_name}中文对比英文相同，原有的中文{file_name}可以继续使用\n")
        source_path = os.path.join(root_path, CN_DIR, file_name)
        dest_path = os.path.join(root_path, CN_NEW_DIR, file_name)
        if not os.path.exists(source_path):
            print(f"源文件 {source_path} 不存在")
        else:
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            shutil.copy2(source_path, dest_path)
            print(f"{file_name}已复制到 Chinese_new 文件夹\n")


def copyFile(file_name):
    if not root_path:
        return
    source_path = os.path.join(root_path, CN_DIR, file_name.lower())
    dest_path = os.path.join(root_path, CN_NEW_DIR, file_name.lower())
    if not os.path.exists(source_path):
        print(f"源文件 {source_path} 不存在")
    else:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        shutil.copy2(source_path, dest_path)
        print(f"{file_name}已复制到 Chinese_new 文件夹\n")


def generate():
    set_root_path()
    if not root_path:
        return

    target_dir = pathlib.Path(os.path.join(root_path, EN_DIR))
    if not target_dir.exists():
        messagebox.showerror("错误", "未找到 English 目录")
        return

    suffix = ".strings"
    files = [f.name.lower() for f in target_dir.glob(f"*{suffix}") if f.is_file()]
    modifiable_files = list(set(files) - {"openings.strings", "cities.strings", "mia.strings"})

    try:
        with open("output.txt", "w", encoding="utf-8") as f:
            sys.stdout = f
            nums = 0
            for name in modifiable_files:
                nums += 1
                print(nums, end=' ')
                generateNewCn(name)
            copyFile("cities.strings")
            copyFile("openings.strings")
            copyFile("mia.strings")
            print("成功！")
    finally:
        sys.stdout = original_stdout

    try:
        with open("output.txt", "r", encoding="utf-8") as f2:
            popout1 = tk.Toplevel(master=root)
            popout1.title("生成日志")
            popout1.geometry("800x600")
            msg_popout1 = tk.Text(master=popout1, font=("Microsoft Yahei", 12))
            msg_popout1.insert("1.0", f2.read())
            msg_popout1.pack(fill="both", expand=True, padx=10, pady=10)
    except Exception as e:
        messagebox.showerror("错误", f"读取日志失败：{e}")


# ==================== Notebook 初始化 ====================
def init_notebook():
    global nb1
    if nb1 is not None:
        return
    style2 = ttk.Style()
    style2.configure("TNotebook.Tab", font=("Microsoft Yahei", base_font_size + 4), padding=(15, 5), foreground="#336699")
    nb1 = ttk.Notebook(master=frame2)
    nb1.bind("<<NotebookTabChanged>>", on_tab_changed)
    nb1.pack(fill="both", expand=True)


def init_contents_and_scrollbar(file_name):
    global current_file
    current_file = file_name
    frame3 = ttk.Frame(master=nb1)
    data1 = ttk.Treeview(
        master=frame3,
        columns=("大写符号", "英文", "中文翻译"),
        show=("tree", "headings")
    )
    data1.heading("#0", text="行号")
    data1.heading("大写符号", text="大写符号")
    data1.heading("英文", text="英文")
    data1.heading("中文翻译", text="中文翻译")
    data1.column("#0", width=80, minwidth=50, anchor="center")
    data1.column("大写符号", width=250, minwidth=150)
    data1.column("英文", width=300, minwidth=150)
    data1.column("中文翻译", width=300, minwidth=150)

    data1.tag_configure("even_color", background="#edf3e5")
    row_count = 1

    cap_list = genCapWordsList(file_name, EN_DIR)
    en_dict = genLangDict(file_name, EN_DIR)
    cn_new_dict = genLangDict(file_name, CN_NEW_DIR)

    for cap_word in cap_list:
        en_text = en_dict.get(cap_word, "None")
        cn_text = cn_new_dict.get(cap_word, "None")
        tag = ("even_color",) if row_count % 2 == 0 else ()
        data1.insert("", "end", text=row_count, values=(cap_word, en_text, cn_text), tags=tag)
        row_count += 1

    scrollbar = ttk.Scrollbar(master=frame3, orient="vertical", command=data1.yview)
    data1.config(yscrollcommand=scrollbar.set)
    data1.bind("<ButtonRelease-1>", on_tree_select)

    scrollbar.pack(side="right", fill="y")
    data1.pack(side="left", fill="both", expand=True)
    return frame3, data1


def add_tab(file_name):
    global created_tabs, current_file, current_tree, current_selected_iid, search_results, current_search_index
    
    save_translation_to_file(force_iid=current_selected_iid)
    current_file = file_name
    tab_text = file_name[:-8]

    if file_name in created_tabs:
        frame3, treeview = created_tabs[file_name]
        nb1.select(frame3)
        current_tree = treeview
        current_selected_iid = None
        current_selected_index = -1
        text1.delete(0, tk.END)
        search_results = []
        current_search_index = -1
        return

    frame3, treeview = init_contents_and_scrollbar(file_name)
    nb1.add(frame3, text=tab_text)
    created_tabs[file_name] = (frame3, treeview)
    nb1.select(frame3)
    current_tree = treeview
    current_selected_iid = None
    current_selected_index = -1
    text1.delete(0, tk.END)
    search_results = []
    current_search_index = -1


# ==================== 加载文件到 Notebook ====================
def list_msgs_dir():
    global root_path, done_init_file_tree, created_tabs, current_tree, current_selected_iid, current_file, search_results, current_search_index

    set_root_path()
    if not root_path:
        return

    cn_new_path = os.path.join(root_path, CN_NEW_DIR)
    if not os.path.exists(cn_new_path):
        messagebox.showinfo(title="提示", message="未找到 Chinese_new 文件夹，请先点击'生成新 Chinese'")
        return

    target_dir = pathlib.Path(os.path.join(root_path, EN_DIR))
    if not target_dir.exists():
        messagebox.showerror("错误", "未找到 English 目录")
        return

    suffix = ".strings"
    files = [f.name.lower() for f in target_dir.glob(f"*{suffix}") if f.is_file()]
    modifiable_files = list(set(files) - {"openings.strings", "cities.strings", "mia.strings"})

    if not modifiable_files:
        messagebox.showinfo("提示", "未找到可编辑的 .strings 文件")
        return

    init_notebook()

    # 清空已有标签页（避免重复）
    for frame, _ in created_tabs.values():
        nb1.forget(frame)
    created_tabs.clear()
    current_tree = None
    current_selected_iid = None
    current_file = None
    current_selected_index = -1
    text1.delete(0, tk.END)
    search_results = []
    current_search_index = -1

    for file_name in sorted(modifiable_files):
        add_tab(file_name)

    done_init_file_tree = True


# ==================== 全局字体更新函数 ====================
def update_all_fonts(event=None):
    global base_font_size
    try:
        offset = int(spinbox.get())
    except ValueError:
        offset = 0
    base_font_size = 14 + offset
    font_family = "Microsoft Yahei"

    style.configure("TButton", font=(font_family, base_font_size + 2))
    style2 = ttk.Style()
    style2.configure("TNotebook.Tab", font=(font_family, base_font_size + 4), padding=(15, 5), foreground="#336699")
    style_tv.configure("Treeview", font=(font_family, base_font_size+4), rowheight=40 + max(0, offset * 2))
    style_tv.configure("Treeview.Heading", font=(font_family, base_font_size + 2))

    search_entry.config(font=(font_family, base_font_size + 2))
    text1.config(font=(font_family, base_font_size + 2))

    for widget in statusbar.winfo_children():
        if isinstance(widget, (ttk.Label, ttk.Spinbox)):
            widget.config(font=(font_family, base_font_size))

    for _, tree in created_tabs.values():
        tree.configure(font=(font_family, base_font_size))


# ==================== 程序关闭前保存 ====================
def on_closing():
    save_translation_to_file(force_iid=current_selected_iid)
    root.destroy()


# ==================== 界面布局 ====================

statusbar = tk.Frame(master=root, relief="sunken", bd=1, pady=5)
statusbar.pack(side="bottom", fill="x")

frame_toolbar = tk.Frame(master=root)
frame_toolbar.pack(side="top", fill="x", padx=10, pady=2)

style = ttk.Style(master=frame_toolbar)
style.configure("TButton", font=("Microsoft Yahei", 16))

tool_button_frame = tk.Frame(frame_toolbar)
tool_button_frame.pack(side="left", anchor="w")

select_msgs_btn = ttk.Button(
    master=tool_button_frame,
    text="生成新 Chinese",
    command=generate
)
select_msgs_btn.pack(side="top", fill="x", pady=2)

open_msgs_btn = ttk.Button(
    master=tool_button_frame,
    text="打开 Messages",
    command=list_msgs_dir
)
open_msgs_btn.pack(side="top", fill="x", pady=2)

spacer = tk.Frame(frame_toolbar)
spacer.pack(side="left", fill="both", expand=True)

search_frame = tk.Frame(spacer)
search_entry = ttk.Entry(search_frame, font=("Microsoft Yahei", 16))
search_entry.delete(0, tk.END)
search_entry.pack(side="left", padx=5)

search_btn = ttk.Button(search_frame, text="搜索", command=search_in_tree)
search_btn.pack(side="left", padx=5)

search_nav_frame = tk.Frame(search_frame)
prev_search_btn = ttk.Button(search_nav_frame, text="上一个", command=goto_prev_search)
next_search_btn = ttk.Button(search_nav_frame, text="下一个", command=goto_next_search)
prev_search_btn.pack(side="top", padx=2)
next_search_btn.pack(side="bottom", padx=2)
search_nav_frame.pack(side="left")

search_frame.pack(side="left", anchor="center", fill="y", pady=16)

right_container = tk.Frame(frame_toolbar)
right_container.pack(side="right", anchor="e")

pre_next_frame = tk.Frame(right_container)
pre_btn = ttk.Button(pre_next_frame, text="上一条", command=select_prev_row)
next_btn = ttk.Button(pre_next_frame, text="下一条", command=select_next_row)
pre_btn.pack(side="top")
next_btn.pack(side="bottom")
pre_next_frame.pack(side="right", anchor="e")

cp_En_btn = ttk.Button(right_container, text="复制英文", command=copy_english_to_clipboard)
cp_En_btn.pack(side="left")

text1 = ttk.Entry(right_container, font=("Microsoft Yahei", 16), width=40)
text1.bind("<FocusOut>", lambda e: save_translation_to_file(force_iid=current_selected_iid))
text1.bind("<Return>", lambda e: save_translation_to_file(force_iid=current_selected_iid))
text1.pack(side="right", padx=10, pady=5)

frame2 = tk.Frame(root)
frame2.pack(side="top", fill="both", expand=True)

style_tv = ttk.Style()
style_tv.configure("Treeview", font=("Microsoft Yahei", 14), rowheight=30)
style_tv.configure("Treeview.Heading", font=("Microsoft Yahei", 14, "bold"))

ttk.Label(statusbar, text="字号增减:").pack(side="left", padx=(10, 5))
spinbox = ttk.Spinbox(statusbar, from_=-10, to=20, width=8, command=update_all_fonts)
spinbox.bind("<KeyRelease>", update_all_fonts)
spinbox.set(0)
spinbox.pack(side="left", padx=(0, 20))

update_all_fonts()

# 🔧 修复：注册关闭事件
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()