# ==============================================================================
# This file was generated/modified by the AtaFilter program.
# ==============================================================================
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re 
from typing import List, Tuple, Any, Dict, Set

# ================= Globals =================
df_fix: pd.DataFrame = None
column_checkboxes = {}      
display_checkboxes = {}     
invalid_value_vars = {}     
# دیکشنری برای جستجوی نقشه تثبیت: {normalized_input: original_target}
input_to_target_map: Dict[str, str] = {} 
# لیست ورودی‌های نقشه که به صورت Regex برای جستجو استفاده خواهند شد
consolidation_patterns: List[Tuple[re.Pattern, str, str]] = [] 
last_clicked_cell = {"item": None, "column_name": None, "column_index": -1} 
temp_editor = None 
tooltip_label = None 

# ================= Language/I18N Globals =================
current_language = "en" # پیش‌فرض انگلیسی
translations = {
    "fa": {
        "title": "AtaFilter - ابزار هوشمند اصلاح و تثبیت داده‌ها",
        "load_file": "۱. انتخاب فایل",
        "choose_fix_col": "۲. ستون اصلاح را انتخاب کنید:",
        "file_not_loaded": "ابتدا فایل را بارگذاری کنید",
        "error_conditions": "۳. شرایط خطا:",
        "cond_empty": "مقدار ستون اصلاح خالی باشد",
        "invalid_values_title": "مقادیر یکتای ستون اصلاح (نامعتبرها را تیک بزنید):",
        "search_cols": "۴. ستون‌های جستجو (برای یافتن پیشنهاد):",
        "load_map_title": "۵. بارگذاری نقشه تثبیت (استان و ورودی‌ها):",
        "select_map_file": "انتخاب فایل مرجع تثبیت",
        "clear_map": "حذف نقشه تثبیت",
        "display_cols": "۶. ستون‌های نمایش در جدول (برای بررسی ردیف):",
        "find_rows": "۷. پیدا کردن ردیف‌ها",
        "select_all": "انتخاب همه ردیف‌ها",
        "row_count": "تعداد ردیف‌های یافت شده: {} | دارای پیشنهاد: {}",
        "apply_save": "۸. اعمال اصلاح و ذخیره فایل جدید",
        "msg_success": "موفق",
        "msg_error": "خطا",
        "msg_file_loaded": "فایل بارگذاری و حروف عربی به فارسی تبدیل شدند. لطفاً ستون اصلاح را انتخاب کنید.",
        "msg_map_loaded": "نقشه تثبیت با {} ورودی یکپارچه‌سازی بارگذاری شد. (حروف عربی در نقشه نیز تبدیل شدند)",
        "msg_map_cleared": "نقشه تثبیت حذف (خالی) شد. جستجو بدون استفاده از آن انجام خواهد شد.",
        "msg_map_col_error": "فایل مرجع باید حداقل ۲ ستون داشته باشد",
        "msg_no_file": "ابتدا فایل و ستون اصلاح را انتخاب کنید",
        "msg_no_fix_col": "لطفاً ستون اصلاح را انتخاب کنید",
        "msg_no_condition": "برای شروع جستجو، باید حداقل یکی از 'شرایط خطا' (ستون خالی یا مقادیر نامعتبر) را در مرحله ۳ فعال نمایید.",
        "msg_no_search_map": "لطفاً حداقل یک ستون را برای جستجو فعال نمایید یا نقشه تثبیت را بارگذاری کنید.",
        "msg_no_fixes": "هیچ مقداری در ستون 'انتخاب نهایی' برای اعمال کردن وجود نداشت. فایل اصلی بدون تغییرات اضافی ذخیره می‌شود.",
        "msg_fixes_applied": "{} اصلاح از 'انتخاب نهایی' در دیتافریم اعمال شد. اکنون پنجره ذخیره فایل باز می‌شود.",
        "msg_save_success": "فایل اصلاح‌شده با موفقیت در {} ذخیره شد.",
        "msg_save_error": "خطا در ذخیره",
        "msg_copy_error": "خطا در عملیات کپی: {}",
        "msg_update_error": "خطا در به‌روزرسانی ردیف: {}",
        "context_copy": "کپی:",
        "context_copy_rows": "کپی ردیف‌های انتخاب شده",
        "context_select_suggestion": "انتخاب پیشنهاد نهایی",
        "context_select": "انتخاب: {}",
        "context_clear_final": "خالی کردن 'انتخاب نهایی'",
        "fixed_cols": ["ردیف", "نوع خطا", "پیشنهاد اصلاح", "منبع پیشنهاد", "مقدار یافت شده", "انتخاب نهایی"],
        "reason_empty": "خالی",
        "reason_invalid": "نامعتبر",
        "reason_consolidation": "تثبیت",
        "source_direct": "تثبیت مستقیم",
        "source_standard": " (استان)",
        "source_consolidation": " (تثبیت)",
        "tooltip_normalized": "نرمال‌شده: {}",
        "creation_text": "ساخته شده توسط برنامه AtaFilter" # متن امضای جدید
    },
    "en": {
        "title": "AtaFilter - Smart Data Cleaning and Consolidation Tool",
        "load_file": "1. Select File",
        "choose_fix_col": "2. Select the Fix Column:",
        "file_not_loaded": "Load File First",
        "error_conditions": "3. Error Conditions:",
        "cond_empty": "Fix Column value is empty",
        "invalid_values_title": "Unique Fix Column Values (Tick Invalid Ones):",
        "search_cols": "4. Search Columns (for suggestions):",
        "load_map_title": "5. Load Consolidation Map (Target and Inputs):",
        "select_map_file": "Select Consolidation Reference File",
        "clear_map": "Clear Consolidation Map",
        "display_cols": "6. Display Columns in Table (for row inspection):",
        "find_rows": "7. Find Rows",
        "select_all": "Select All Rows",
        "row_count": "Found Rows: {} | With Suggestion: {}",
        "apply_save": "8. Apply Fixes and Save New File",
        "msg_success": "Success",
        "msg_error": "Error",
        "msg_file_loaded": "File loaded and Arabic characters converted to Persian. Please select the fix column.",
        "msg_map_loaded": "Consolidation map loaded with {} consolidation entries. (Arabic chars in map also converted)",
        "msg_map_cleared": "Consolidation map cleared. Searching will be done without it.",
        "msg_map_col_error": "Reference file must have at least 2 columns",
        "msg_no_file": "Please load a file and select the fix column first",
        "msg_no_fix_col": "Please select the fix column",
        "msg_no_condition": "To start searching, you must enable at least one 'Error Condition' (Empty Column or Invalid Values) in step 3.",
        "msg_no_search_map": "Please enable at least one Search Column or load a Consolidation Map.",
        "msg_no_fixes": "No values in the 'Final Selection' column were found to apply. The original file will be saved without extra changes.",
        "msg_fixes_applied": "{} fixes from 'Final Selection' were applied to the DataFrame. The save file dialog will now open.",
        "msg_save_success": "The fixed file was successfully saved to {}.",
        "msg_save_error": "Error during save",
        "msg_copy_error": "Error in copy operation: {}",
        "msg_update_error": "Error updating row: {}",
        "context_copy": "Copy:",
        "context_copy_rows": "Copy Selected Rows",
        "context_select_suggestion": "Select Final Suggestion",
        "context_select": "Select: {}",
        "context_clear_final": "Clear 'Final Selection'",
        "fixed_cols": ["Row", "Error Type", "Suggestion", "Source", "Found Value", "Final Selection"],
        "reason_empty": "Empty",
        "reason_invalid": "Invalid",
        "reason_consolidation": "Consolidation",
        "source_direct": "Direct Consolidation",
        "source_standard": " (Standard)",
        "source_consolidation": " (Consolidation)",
        "tooltip_normalized": "Normalized: {}",
        "creation_text": "Created by AtaFilter program" # متن امضای جدید
    }
}

def get_text(key: str, lang: str = None) -> Any:
    """Helper to get text based on current language."""
    if lang is None:
        lang = current_language
    return translations.get(lang, translations["en"]).get(key, f"MISSING_TEXT_{key}")


# ---------------- Utility Functions ----------------

def convert_arabic_to_persian_in_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    تبدیل حروف عربی (ك, ي) به معادل فارسی آنها (ک, ی) در تمام ستون‌های متنی.
    """
    def convert_text(text):
        if pd.isna(text) or text is None:
            return text
        text = str(text)
        # انجام تبدیل حروف عربی به فارسی
        text = text.replace('ي', 'ی').replace('ك', 'ک')
        return text

    # تلاش برای اعمال تبدیل فقط روی ستون‌های نوع 'object' (که معمولاً رشته هستند)
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                df[col] = df[col].apply(convert_text)
            except Exception:
                pass # در صورت بروز خطا در تبدیل، ادامه می‌دهد
    return df

def destroy_temp_editor():
    """ویجت ویرایشگر موقت را حذف می‌کند و متغیر آن را پاک می‌کند."""
    global temp_editor
    if temp_editor and temp_editor.winfo_exists():
        temp_editor.destroy()
    temp_editor = None

def normalize_persian_string(text: Any) -> str:
    """
    نرمال‌سازی کامل و دقیق.
    - جداکننده‌های رایج (مثل خط تیره) را به فاصله تبدیل می‌کند.
    - فضاهای نامرئی و فاصله‌های اضافی را حذف می‌کند.
    """
    if pd.isna(text) or text is None: return ""
    text = str(text)
    
    # 1. NEW: جایگزینی جداکننده‌های رایج (خط تیره، کاما، اسلش، نقطه، پرانتز) با فاصله
    # این کار تضمین می‌کند که کلمات کنار جداکننده‌ها به درستی تفکیک شوند.
    text = re.sub(r'[-،,/\.\(\)]', ' ', text)
    
    # 2. حذف کاراکترهای کنترلی، نامرئی و فضاهای نیم‌فاصله
    text = re.sub(r'[\u200c\u200b\u202e\ufeff\xad\xa0\u200f]', '', text) 
    
    # 3. کاهش چندین فاصله متوالی به یک فاصله
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip().lower()

def get_word_tokens(text: str) -> Set[str]:
    """
    متن را به کلمات (توکن‌ها) تقسیم می‌کند.
    """
    if pd.isna(text) or text is None: return set()
    
    # 1. نرمال‌سازی کلی (شامل تبدیل جداکننده‌ها به فاصله)
    text = normalize_persian_string(text)
    
    # 2. شکستن به کلمات بر اساس فاصله و فیلتر کردن کلمات خالی
    tokens = {t for t in text.split(' ') if t}
    return tokens


# ---------------- Tooltip Functions (Debug) ----------------

def show_tooltip(event):
    """نمایش tooltip حاوی محتوای نرمال‌شده سلول."""
    global tooltip_label
    
    item = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)
    
    if item and col_id:
        try:
            content = tree.set(item, col_id)
            normalized_content = normalize_persian_string(content)
            
            if tooltip_label: tooltip_label.destroy()
            
            tooltip_label = tk.Label(root, 
                                     text=get_text("tooltip_normalized").format(normalized_content),
                                     relief='solid', 
                                     borderwidth=1, 
                                     bg='light yellow',
                                     fg='black',
                                     font=("Arial", 9))
            
            tooltip_label.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            tooltip_label.lift()
        
        except Exception:
            hide_tooltip()
    else:
        hide_tooltip()

def hide_tooltip(event=None):
    """پنهان کردن tooltip."""
    global tooltip_label
    if tooltip_label:
        tooltip_label.destroy()
        tooltip_label = None

# ---------------- Sorting Functions ----------------

def treeview_sort_column(tv, col, reverse):
    """مرتب‌سازی استاندارد Treeview بر اساس ستون کلیک شده."""
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    try:
        l.sort(key=lambda t: float(re.sub(r'[^\d.]', '', t[0])), reverse=reverse)
    except ValueError:
        l.sort(key=lambda t: t[0], reverse=reverse)

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col, command=lambda _col=col: treeview_sort_column(tv, _col, not reverse))


def treeview_sort_by_suggestion_count(tv):
    """مرتب‌سازی Treeview بر اساس تعداد پیشنهادات (نزولی)."""
    l = []
    for k in tv.get_children(''):
        suggestion_string = tv.set(k, 'suggestion')
        if suggestion_string:
            count = len(suggestion_string.split(" | "))
        else:
            count = 0
        l.append((count, k))

    l.sort(key=lambda t: t[0], reverse=True)

    for index, (count, k) in enumerate(l):
        tv.move(k, '', index)

    for col in tv['columns']:
        tv.heading(col, command=lambda _col=col: treeview_sort_column(tv, _col, False))


# ---------------- Copy Function ----------------
def copy_selected_content(item_id, column_name):
    """محتوای یک سلول مشخص یا محتوای کل ردیف‌های انتخاب شده را کپی می‌کند."""
    try:
        root.clipboard_clear()
        selected_items = tree.selection()

        if not selected_items:
            return

        if len(selected_items) == 1:
            if item_id and column_name and column_name in tree['columns']:
                content = str(tree.set(item_id, column_name))
            else:
                row_data = [str(v) for v in tree.item(item_id, 'values')]
                content = "\t".join(row_data) 
            
            root.clipboard_append(content)

        else:
            all_rows = []
            for item in selected_items:
                row_data = [str(v) for v in tree.item(item, 'values')]
                all_rows.append("\t".join(row_data))
            content = "\n".join(all_rows)
            root.clipboard_append(content)

    except Exception as e:
        messagebox.showerror(get_text("msg_error"), get_text("msg_copy_error").format(e))
        pass
        

# ---------------- Set Final Selection ----------------
def set_final_selection(selected_value):
    """مقدار نهایی انتخاب شده را در ستون 'انتخاب نهایی' ردیف کلیک شده قرار می‌دهد."""
    item = last_clicked_cell.get("item")
    
    if item and tree.exists(item): 
        try:
            current_values = list(tree.item(item, 'values'))
            
            if len(current_values) > 5:
                current_values[5] = selected_value
                tree.item(item, values=current_values) 
                tree.selection_set(item)
                last_clicked_cell["item"] = None
        except Exception as e:
            messagebox.showerror(get_text("msg_error"), get_text("msg_update_error").format(e))

# ---------------- Open Editor (for Double Click) ----------------
def open_editor(event):
    """باز کردن ویرایشگر متن هنگام دوبار کلیک."""
    global last_clicked_cell, temp_editor
    
    if temp_editor and temp_editor.winfo_exists():
        temp_editor.destroy()
        temp_editor = None
        return

    item = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)
    
    if not item or not col_id: return

    try:
        col_index = int(col_id.replace('#', '')) - 1
        column_name = tree['columns'][col_index]
    except (ValueError, IndexError):
        return

    last_clicked_cell["item"] = item
    last_clicked_cell["column_name"] = column_name
    last_clicked_cell["column_index"] = col_index

    if column_name in ["final_selection"]: 
        x, y, width, height = tree.bbox(item, col_id)
        
        temp_editor = ttk.Entry(tree, font=("Arial", 10))
        temp_editor.place(x=x, y=y, width=width, height=height)
        
        current_value = tree.set(item, col_id)
        temp_editor.insert(0, current_value)
        
        temp_editor.focus_set()
        temp_editor.select_range(0, 'end')
        
        def on_editor_exit(e):
            tree.set(item, col_id, temp_editor.get())
            destroy_temp_editor()
            
        temp_editor.bind("<Return>", on_editor_exit)
        temp_editor.bind("<FocusOut>", lambda e: destroy_temp_editor())

# ---------------- Load File / Columns / Map / Find Rows / Select All / Apply & Save ----------------
def browse_file():
    global df_fix
    path = filedialog.askopenfilename(filetypes=[("CSV or Excel","*.csv *.xlsx *.xls")])
    if not path: return
    try:
        if path.endswith(".csv"):
            try: df_fix = pd.read_csv(path, encoding='utf-8-sig') # Try utf-8-sig first for better CSV support
            except UnicodeDecodeError: 
                 try: df_fix = pd.read_csv(path, encoding='utf-8')
                 except UnicodeDecodeError: df_fix = pd.read_csv(path, encoding='latin-1')
        else: df_fix = pd.read_excel(path)
        
        df_fix.columns = df_fix.columns.astype(str).str.strip()
        
        # تبدیل حروف عربی به فارسی
        df_fix = convert_arabic_to_persian_in_df(df_fix)
        
        fix_column_cb["values"] = df_fix.columns.tolist()
        fix_column_cb.config(state="readonly")
        fix_column_cb.set(get_text("choose_fix_col"))
        load_search_columns(); load_display_columns() 
        messagebox.showinfo(get_text("msg_success"), get_text("msg_file_loaded"))
    except Exception as e: messagebox.showerror(get_text("msg_error"), str(e))


def load_invalid_values(event=None):
    global invalid_value_vars
    for w in invalid_inner_frame.winfo_children(): w.destroy()
    invalid_value_vars.clear()
    if df_fix is None: return
    col = fix_column_cb.get()
    if col not in df_fix.columns: return
    
    # --- مرتب‌سازی الفبایی ---
    values = sorted(df_fix[col].dropna().astype(str).str.strip().unique())
    
    max_cols = 2 
    for i, v in enumerate(values):
        var = tk.BooleanVar(value=False); r = i // max_cols; c = i % max_cols
        
        # --- استفاده از Checkbutton استاندارد ---
        tk.Checkbutton(invalid_inner_frame, text=v, variable=var).grid(row=r, column=c, sticky="w", padx=5, pady=2)
        
        invalid_value_vars[v] = var

def load_search_columns():
    for w in search_inner_frame.winfo_children(): w.destroy()
    column_checkboxes.clear()
    if df_fix is None: return
    max_cols = 2 
    for i, col in enumerate(df_fix.columns):
        var = tk.BooleanVar(value=False); r = i // max_cols; c = i % max_cols
        tk.Checkbutton(search_inner_frame, text=col, variable=var).grid(row=r, column=c, sticky="w", padx=5, pady=2)
        column_checkboxes[col] = var
        
def load_display_columns():
    for w in display_inner_frame.winfo_children(): w.destroy()
    display_checkboxes.clear()
    if df_fix is None: return
    fix_col_name = fix_column_cb.get() if fix_column_cb.get() in df_fix.columns else None
    max_cols = 2 
    for i, col in enumerate(df_fix.columns):
        is_checked = (col == fix_col_name); var = tk.BooleanVar(value=is_checked)
        r = i // max_cols; c = i % max_cols
        tk.Checkbutton(display_inner_frame, text=col, variable=var).grid(row=r, column=c, sticky="w", padx=5, pady=2)
        display_checkboxes[col] = var

def load_consolidation_map():
    global input_to_target_map, consolidation_patterns
    path = filedialog.askopenfilename(title=get_text("select_map_file") + " (CSV or Excel)", filetypes=[("CSV or Excel","*.csv *.xlsx *.xls")])
    if not path: return
    try:
        if path.endswith(".csv"):
            try: df_map = pd.read_csv(path, encoding='utf-8-sig')
            except UnicodeDecodeError: 
                 try: df_map = pd.read_csv(path, encoding='utf-8')
                 except UnicodeDecodeError: df_map = pd.read_csv(path, encoding='latin-1')
        else: df_map = pd.read_excel(path)
        
        # تبدیل حروف عربی در نقشه تثبیت نیز ضروری است
        df_map = convert_arabic_to_persian_in_df(df_map) 
        
        if len(df_map.columns) < 2: messagebox.showerror(get_text("msg_error"), get_text("msg_map_col_error")); return
        
        input_to_target_map = {}
        consolidation_patterns = []
        target_col_name = df_map.columns[0]; input_col_name = df_map.columns[1]
        normalized_inputs_set = set()

        for index, row in df_map.iterrows():
            target_value = str(row[target_col_name]).strip()
            input_cell = str(row[input_col_name]).strip()
            if target_value and input_cell: 
                input_list = re.split(r'[,،]', input_cell) 
                for input_value in input_list:
                    input_value_stripped = input_value.strip()
                    if input_value_stripped: 
                        normalized_input = normalize_persian_string(input_value_stripped)
                        if normalized_input and normalized_input not in normalized_inputs_set:
                            
                            # 1. ذخیره مقدار هدف
                            input_to_target_map[normalized_input] = target_value
                            normalized_inputs_set.add(normalized_input) 
                            
                            # 2. ذخیره الگو برای جستجوی دقیق زیررشته (برای حفظ چند کلمه‌ای‌ها)
                            # این الگو به شکل (الگوی Regex کامپایل شده، مقدار ورودی نرمال‌شده، مقدار هدف) ذخیره می‌شود
                            # از \b برای پیدا کردن کل کلمه استفاده می‌کنیم.
                            regex_pattern = re.compile(r'\b' + re.escape(normalized_input) + r'\b', re.IGNORECASE)
                            consolidation_patterns.append((regex_pattern, normalized_input, target_value))
                            
                            
        messagebox.showinfo(get_text("msg_success"), get_text("msg_map_loaded").format(len(input_to_target_map)))
    except Exception as e: messagebox.showerror(get_text("msg_error"), str(e))
        
def clear_consolidation_map():
    """خالی کردن نقشه تثبیت از حافظه."""
    global input_to_target_map, consolidation_patterns
    input_to_target_map.clear()
    consolidation_patterns.clear()
    messagebox.showinfo(get_text("msg_success"), get_text("msg_map_cleared"))

def update_treeview_columns(display_cols):
    fixed_cols = ["row", "reason", "suggestion", "source", "found_value", "final_selection"]
    fixed_titles = get_text("fixed_cols")
    dynamic_cols = display_cols; dynamic_titles = display_cols
    all_cols = fixed_cols + dynamic_cols; all_titles = fixed_titles + dynamic_titles
    tree.config(columns=all_cols); tree.delete(*tree.get_children()) 
    for col_id, title in zip(all_cols, all_titles):
        tree.heading(col_id, text=title, command=lambda _col=col_id: treeview_sort_column(tree, _col, False)) 
        if col_id == "row": tree.column(col_id, width=50, anchor="center")
        elif col_id == "reason": tree.column(col_id, width=100, anchor="center")
        elif col_id == "suggestion": tree.column(col_id, width=150, anchor="w")
        elif col_id == "source": tree.column(col_id, width=80, anchor="w")
        elif col_id == "found_value": tree.column(col_id, width=150, anchor="w") 
        elif col_id == "final_selection": tree.column(col_id, width=150, anchor="w") 
        else: tree.column(col_id, width=120, anchor="w")

def find_rows():
    tree.delete(*tree.get_children())
    if df_fix is None: messagebox.showerror(get_text("msg_error"), get_text("msg_no_file")); return
    fix_col = fix_column_cb.get()
    if fix_col not in df_fix.columns: messagebox.showerror(get_text("msg_error"), get_text("msg_no_fix_col")); return
    
    search_cols = [c for c,v in column_checkboxes.items() if v.get()]
    display_cols = [c for c,v in display_checkboxes.items() if v.get()] 
    invalid_values = [v for v,var in invalid_value_vars.items() if var.get()]
    
    has_error_condition = cond_empty.get() or invalid_values
    if not has_error_condition:
        messagebox.showerror(get_text("msg_error"), get_text("msg_no_condition"))
        return
        
    has_search_cols = bool(search_cols)

    # اگر ستون جستجو نداریم و نقشه تثبیت هم نداریم، نمی‌توانیم پیشنهادی بدهیم
    if not has_search_cols and not bool(input_to_target_map):
         messagebox.showerror(get_text("msg_error"), get_text("msg_no_search_map"))
         return

    has_consolidation_map = bool(input_to_target_map)
    update_treeview_columns(display_cols)
    count = 0 
    count_with_suggestion = 0 

    # ساخت نقشه مقادیر استاندارد (فقط ورودی‌های تک کلمه‌ای)
    unique_fix_values = df_fix[fix_col].dropna().astype(str).str.strip().unique()
    original_value_map = {normalize_persian_string(v): v for v in unique_fix_values if v.strip() != ""}; 
    
    # {توکن_تک_کلمه‌ای: مقدار_اصلی_ستون_اصلاح}
    standard_search_tokens = {}
    for normalized_val, original_val in original_value_map.items():
        # فقط توکن‌هایی به دیکشنری استاندارد اضافه می‌شوند که تک کلمه‌ای باشند
        if len(normalized_val.split()) == 1:
            standard_search_tokens[normalized_val] = original_val
    
    
    for idx,row in df_fix.iterrows():
        value = str(row[fix_col]).strip() if pd.notna(row[fix_col]) else ""
        normalized_value = normalize_persian_string(value)
        
        reasons = []; 
        
        is_empty_error = (value == "" and cond_empty.get())
        is_invalid_error = (value in invalid_values)
        
        # A. بررسی پیشینی: پیشنهاد تثبیت مستقیم (اولویت 0 - سناریو 2)
        is_direct_consolidation_error = False
        if has_consolidation_map and normalized_value in input_to_target_map:
            target_value_map = input_to_target_map[normalized_value]
            # اگر مقدار فعلی با هدف تثبیت متفاوت است، پس خطای تثبیت است.
            if target_value_map != value: 
                is_direct_consolidation_error = True

        # *** فیلتر اولیه ردیف‌ها: اگر شرط خطا برقرار نیست، رد شو ***
        if not (is_empty_error or is_invalid_error or is_direct_consolidation_error): continue 
        
        if is_empty_error: reasons.append(get_text("reason_empty"))
        if is_invalid_error: reasons.append(get_text("reason_invalid"))
        # توجه: خطای تثبیت را فقط در صورتی اضافه می‌کنیم که خطا "خالی" نباشد (تا اولویت‌بندی صحیح حفظ شود)
        if is_direct_consolidation_error and not is_empty_error: reasons.append(get_text("reason_consolidation"))
        
        # 2. ردیابی پیشنهادات
        # (اولویت, موقعیت_پیدایش, پیشنهاد, منبع, مقدار یافت شده)
        # اولویت: 0 (Direct Map), 1 (Standard Token), 2 (Map Regex in Search Col)
        # موقعیت پیدایش: -1 (برای P0، یعنی بدون موقعیت) یا Index شروع کلمه/الگو در رشته جستجو
        final_suggestions_data: List[Tuple[int, int, str, str, str]] = [] 
        suggestion_list = set() 
        
        # ----------------------------------------------------------------------
        # P0: تثبیت مستقیم (نقشه تثبیت با خود مقدار ستون اصلاح) - بالاترین اولویت
        # ----------------------------------------------------------------------
        if is_direct_consolidation_error and not is_empty_error:
            suggestion = input_to_target_map[normalized_value]
            if suggestion not in suggestion_list:
                suggestion_list.add(suggestion)
                # اولویت 0، موقعیت -1 (موقعیت برای این نوع جستجو اهمیتی ندارد)
                final_suggestions_data.append((0, -1, suggestion, get_text("source_direct"), value))

        cols_to_search_in = search_cols.copy()
        if fix_col in cols_to_search_in:
            cols_to_search_in.remove(fix_col)
            
        if cols_to_search_in:
            
            # ----------------------------------------------------------------------
            # P1: جستجو با توکن‌های تک کلمه‌ای استاندارد (سناریو 1)
            # ----------------------------------------------------------------------
            for col in cols_to_search_in:
                original_cell = str(row[col])
                # از نرمال‌شده برای پیدا کردن موقعیت توکن استفاده می‌کنیم
                normalized_cell = normalize_persian_string(original_cell) 
                
                # توجه: normalized_cell_tokens اکنون با توکن‌سازی بهبودیافته کار می‌کند.
                normalized_cell_tokens = get_word_tokens(original_cell) 
                
                for token in normalized_cell_tokens:
                    # توکن باید یک کلمه باشد و در نقشه توکن‌های استاندارد (که فقط تک کلمات را دارد) وجود داشته باشد
                    if token in standard_search_tokens:
                        suggestion = standard_search_tokens[token]
                        
                        # اگر مقدار استاندارد پیدا شده، همان مقدار نامعتبر فعلی بود، آن را پیشنهاد نده
                        if is_invalid_error and normalize_persian_string(suggestion) == normalized_value:
                            continue
                            
                        if suggestion not in suggestion_list:
                            suggestion_list.add(suggestion)
                            
                            # پیدا کردن موقعیت توکن نرمال‌شده در رشته نرمال‌شده ستون جستجو
                            match_position = normalized_cell.find(token)
                            
                            # اولویت 1
                            final_suggestions_data.append((1, match_position, suggestion, col + get_text("source_standard"), token))

            # ----------------------------------------------------------------------
            # P2: جستجو با الگوهای Regex نقشه تثبیت در ستون‌های جستجو (سناریو 3)
            # ----------------------------------------------------------------------
            if has_consolidation_map:
                for pattern, normalized_input, suggestion in consolidation_patterns:

                    # اگر مقدار پیشنهاد تثبیت همان مقدار استاندارد فعلی بود، آن را پیشنهاد نده
                    if is_invalid_error and normalize_persian_string(suggestion) == normalized_value:
                        continue
                        
                    for col in cols_to_search_in:
                        original_cell = str(row[col])
                        
                        # جستجو با الگوی Regex در کل محتوای سلول جستجو
                        match = pattern.search(original_cell) 
                        if match:
                            if suggestion not in suggestion_list: 
                                suggestion_list.add(suggestion)
                                
                                match_position = match.start() # موقعیت دقیق در رشته اصلی
                                # اولویت 2
                                final_suggestions_data.append((2, match_position, suggestion, col + get_text("source_consolidation"), normalized_input))
                                
                                # بلافاصله پس از یافتن در یک ستون، به سراغ الگوی بعدی برو
                                break 
                
        # 3. مرتب‌سازی و استخراج نتایج نهایی
        initial_final_selection = ""
        final_suggestion_str = ""
        source = ""; found_value = ""

        if final_suggestions_data:
            
            # مرتب‌سازی بر اساس:
            # 1. موقعیت پیدایش (Match_Position) - صعودی (کوچکترین ایندکس برنده است)
            # 2. اولویت (Priority) - صعودی (برای شکستن تساوی در موقعیت، P0 برنده P1 است)
            final_suggestions_data.sort(key=lambda x: (x[1], x[0])) 
            
            # بهترین پیشنهاد، اولین مقدار پس از مرتب‌سازی است
            # best_match: (Priority, Match_Position, Suggestion, Source, Found_Value)
            best_match = final_suggestions_data[0]
            initial_final_selection = best_match[2].strip() 
            source = best_match[3]
            found_value = best_match[4]
            
            # ساخت لیست پیشنهادات مرتب شده
            suggestion_list_ordered = []
            seen_suggestions = set()
            # استفاده از لیست مرتب شده برای ساخت رشته پیشنهاد
            for priority, position, suggestion, _, _ in final_suggestions_data:
                 if suggestion not in seen_suggestions:
                     suggestion_list_ordered.append(suggestion)
                     seen_suggestions.add(suggestion)
                     
            final_suggestion_str = " | ".join(suggestion_list_ordered)
            
        
        # 4. افزودن به Treeview
        fixed_values = [idx+1, ",".join(reasons), final_suggestion_str, source, found_value, initial_final_selection] 
        display_values = [str(row[col]) for col in display_cols]
        
        tree.insert("", "end", values=fixed_values + display_values)
        count +=1
        
        if final_suggestion_str:
            count_with_suggestion += 1
            
    count_label.config(text=get_text("row_count").format(count, count_with_suggestion))
    treeview_sort_by_suggestion_count(tree)


def select_all_rows():
    for item in tree.get_children(): tree.selection_add(item)

# ----------------- تابع اعمال اصلاح و ذخیره (اصلاح شده) -----------------
def apply_fixes_and_save():
    """اعمال اصلاحات از ستون 'انتخاب نهایی' به دیتافریم اصلی و ذخیره فایل جدید."""
    if df_fix is None: 
        messagebox.showerror(get_text("msg_error"), get_text("msg_no_file"))
        return

    fix_col = fix_column_cb.get()
    if fix_col not in df_fix.columns: 
        messagebox.showerror(get_text("msg_error"), get_text("msg_no_fix_col"))
        return

    # --- 1. اعمال اصلاحات ---
    fixes = []
    fixed_indices = set()
    
    for item in tree.get_children(''):
        values = tree.item(item, 'values')
        try:
            # ردیف از ایندکس 1 شروع می‌شود، پس باید 1 واحد کم شود
            idx = int(values[0]) - 1 
        except (ValueError, IndexError):
            continue
        
        final_selection = values[5].strip() if len(values) > 5 else "" 
        
        if final_selection != "":
            if idx not in fixed_indices:
                df_fix.loc[idx, fix_col] = final_selection
                fixes.append((idx + 1, final_selection))
                fixed_indices.add(idx)

    if fixes:
        messagebox.showinfo(get_text("msg_success"), get_text("msg_fixes_applied").format(len(fixes)))
    else:
        messagebox.showwarning(get_text("msg_error"), get_text("msg_no_fixes"))

    # --- 2. ذخیره فایل (پیش فرض CSV) ---
    path = filedialog.asksaveasfilename(defaultextension=".csv",
                                        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    if not path: 
        if fixes: find_rows() 
        return

    try:
        if path.endswith(".csv"):
            df_fix.to_csv(path, index=False, encoding='utf-8-sig') 
        elif path.endswith((".xlsx", ".xls")):
            # متن امضا
            creation_text = get_text("creation_text")
            
            # استفاده از ExcelWriter برای نوشتن دیتافریم و متن
            with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                # 1. نوشتن دیتافریم
                # نیازی به writer.book و writer.sheets['Sheet1'] نیست اگر index=False باشد، 
                # اما برای نوشتن متن در سطر بعد، باید به شیء sheet دسترسی داشته باشیم.
                
                # برای سادگی، از یک شیت با نام 'Data' استفاده می‌کنیم
                df_fix.to_excel(writer, sheet_name='Data', index=False)
                
                # 2. نوشتن متن در سطر انتهایی
                workbook = writer.book
                worksheet = writer.sheets['Data'] 
                
                # سطر برای نوشتن: یک سطر فاصله بعد از آخرین سطر داده (len(df_fix) + 1 + 1)
                # +1 برای هدرها و +1 برای یک خط خالی
                start_row = len(df_fix) + 2 
                
                # یک فرمت ساده برای متن امضا
                text_format = workbook.add_format({'bold': False, 'italic': True, 'color': '#777777', 'font_size': 9})
                
                # نوشتن متن در ستون A (این یک شیت بدون ایندکس است، بنابراین ستون اول A است)
                worksheet.write(start_row, 0, creation_text, text_format)

        messagebox.showinfo(get_text("msg_success"), get_text("msg_save_success").format(path))
    except Exception as e:
        messagebox.showerror(get_text("msg_save_error"), str(e))
        
    # --- 3. حفظ جدول نتایج: پس از ذخیره، Treeview را با داده‌های جدید به‌روز می‌کنیم ---
    find_rows()
# ----------------- پایان تابع اعمال اصلاح و ذخیره (اصلاح شده) -----------------


# ---------------- Context Menu (راست کلیک) ----------------
def show_context_menu(event):
    """نمایش منوی راست کلیک و به‌روزرسانی سلول کلیک شده."""
    destroy_temp_editor() 
    hide_tooltip() 
    context_menu.delete(0, 'end')
    
    tree.focus_set() 
    
    item = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)
    
    column_name_actual = None
    col_index = -1
    
    if item and col_id:
        tree.focus(item); tree.selection_set(item)
        
        all_column_names = tree['columns']
        column_name_raw = col_id.replace('#', '')

        if column_name_raw.isdigit():
            col_index = int(column_name_raw) - 1
            if 0 <= col_index < len(all_column_names):
                column_name_actual = all_column_names[col_index]
        else:
            column_name_actual = column_name_raw
            try: col_index = all_column_names.index(column_name_actual)
            except ValueError: pass

        last_clicked_cell["item"] = item; 
        last_clicked_cell["column_name"] = column_name_actual; 
        last_clicked_cell["column_index"] = col_index
        
        if column_name_actual:
            context_menu.add_command(label=f"{get_text('context_copy')} {tree.heading(column_name_actual, 'text')}", 
                                     command=lambda i=item, c=column_name_actual: copy_selected_content(i, c))
        
        context_menu.add_command(label=get_text("context_copy_rows"), 
                                 command=lambda: copy_selected_content(None, None))
        
        if column_name_actual == "suggestion": 
            values = tree.item(item)["values"]
            if len(values) > 2:
                suggestion_full = values[2] 
                
                if suggestion_full:
                    suggestions = [s.strip() for s in suggestion_full.split(" | ")]
                    context_menu.add_separator()
                    
                    select_menu = tk.Menu(context_menu, tearoff=0)
                    for s in suggestions:
                        select_menu.add_command(label=get_text("context_select").format(s), command=lambda val=s: set_final_selection(val))
                    
                    context_menu.add_cascade(label=get_text("context_select_suggestion"), menu=select_menu)
        
        if column_name_actual == "final_selection":
             current_final = tree.item(item)["values"][5].strip() if len(tree.item(item)["values"]) > 5 else ""
             if current_final:
                 context_menu.add_command(label=get_text("context_clear_final"), 
                                         command=lambda: set_final_selection(""))
    
    if tree.selection():
        context_menu.post(event.x_root, event.y_root)

# ---------------- Language Selection ----------------
def change_language(event=None):
    """
    تغییر زبان فعلی و به‌روزرسانی تمام متون در رابط کاربری.
    """
    global current_language
    selected_lang_display = language_cb.get()
    
    # نگاشت از نام نمایش به کد زبان
    lang_map = {"فارسی": "fa", "English": "en"}
    new_lang_code = lang_map.get(selected_lang_display, "en") 
    
    if current_language != new_lang_code:
        current_language = new_lang_code
        update_texts()

def update_texts():
    """به‌روزرسانی تمام متون ثابت در رابط کاربری بر اساس current_language."""
    
    # 1. عنوان
    root.title(get_text("title"))
    
    # 2. دکمه‌ها و لیبل‌ها
    load_file_btn.config(text=get_text("load_file"))
    fix_col_label.config(text=get_text("choose_fix_col"))
    
    # به‌روزرسانی Combobox ستون اصلاح (نیاز به بازنشانی متن پیش‌فرض)
    # این فقط برای حالت Disabled است
    if fix_column_cb["state"] == "disabled":
        fix_column_cb.set(get_text("file_not_loaded"))
    elif fix_column_cb["state"] == "readonly":
        # اگر ستونی انتخاب نشده است، متن پیش‌فرض را نمایش دهد
        if fix_column_cb.get() not in fix_column_cb["values"]: 
            fix_column_cb.set(get_text("choose_fix_col")) 
        
    error_cond_label.config(text=get_text("error_conditions"))
    cond_empty_cb.config(text=get_text("cond_empty"))
    invalid_val_label.config(text=get_text("invalid_values_title"))
    
    # ۴. ستون‌های جستجو (متن لیبل)
    search_col_label.config(text=get_text("search_cols"))
    
    # ۵. نقشه تثبیت
    load_map_label.config(text=get_text("load_map_title"))
    load_map_btn.config(text=get_text("select_map_file"))
    clear_map_btn.config(text=get_text("clear_map"))
    
    # ۶. ستون‌های نمایش
    display_col_label.config(text=get_text("display_cols"))
    
    # ۷. دکمه پیدا کردن ردیف‌ها
    find_rows_btn.config(text=get_text("find_rows"))
    select_all_btn.config(text=get_text("select_all"))
    
    # ۸. دکمه نهایی
    apply_save_btn.config(text=get_text("apply_save"))
    
    # ۹. Treeview (به‌روزرسانی عنوان ستون‌های ثابت)
    display_cols = [c for c,v in display_checkboxes.items() if v.get()]
    update_treeview_columns(display_cols)
    
    # ۱۰. به‌روزرسانی تعداد ردیف‌ها (اگر جدولی پر است)
    try:
        current_count_text = count_label.cget("text")
        if " | " in current_count_text or "تعداد" in current_count_text or "Found Rows" in current_count_text:
            # سعی در استخراج اعداد از متن فعلی
            match = re.search(r'(\d+).*?(\d+)', current_count_text)
            if match:
                count, count_with_suggestion = match.groups()
                count_label.config(text=get_text("row_count").format(count, count_with_suggestion))
            else:
                count_label.config(text=get_text("row_count").format(0, 0))
        else:
            count_label.config(text=get_text("row_count").format(0, 0))
    except Exception:
        count_label.config(text=get_text("row_count").format(0, 0))
        

# ---------------- UI ----------------
root = tk.Tk()
root.title(get_text("title"))
root.geometry("1400x850") 

main_horizontal_frame = tk.Frame(root)
main_horizontal_frame.pack(fill="both", expand=True)

# ---------------- پنل سمت چپ (کنترل‌ها) ----------------
control_canvas = tk.Canvas(main_horizontal_frame, width=400) 
control_canvas.pack(side=tk.LEFT, fill="y")
control_scrollbar = tk.Scrollbar(main_horizontal_frame, orient="vertical", command=control_canvas.yview)
control_scrollbar.pack(side=tk.LEFT, fill="y")
control_canvas.configure(yscrollcommand=control_scrollbar.set)

control_frame = tk.Frame(control_canvas)
control_canvas.create_window((0,0), window=control_frame, anchor="nw", width=400) 

def on_control_frame_config(event):
    control_canvas.configure(scrollregion=control_canvas.bbox("all"))

control_frame.bind("<Configure>", on_control_frame_config)


# --- NEW: Language Selection ---
language_frame = tk.Frame(control_frame)
language_frame.pack(pady=5, padx=10, fill='x')
tk.Label(language_frame, text="Language:").pack(side=tk.LEFT)
language_cb = ttk.Combobox(language_frame, values=["English", "فارسی"], state="readonly", width=15)
language_cb.set("English") # پیش‌فرض انگلیسی
language_cb.pack(side=tk.RIGHT)
language_cb.bind("<<ComboboxSelected>>", change_language)
# --- END NEW ---

# ۱. انتخاب فایل و ستون اصلاح
load_file_btn = tk.Button(control_frame, text=get_text("load_file"), command=browse_file)
load_file_btn.pack(pady=5)
fix_col_label = tk.Label(control_frame, text=get_text("choose_fix_col"))
fix_col_label.pack(pady=(10, 0))
fix_column_cb = ttk.Combobox(control_frame, state="disabled", width=50, justify='center')
fix_column_cb.set(get_text("file_not_loaded"))
fix_column_cb.pack(pady=5, padx=10)
fix_column_cb.bind("<<ComboboxSelected>>", load_invalid_values)
fix_column_cb.bind("<<ComboboxSelected>>", load_display_columns, add="+")

# ۳. شرایط خطا
error_cond_label = tk.Label(control_frame, text=get_text("error_conditions"))
error_cond_label.pack(pady=(10, 0))
cond_empty = tk.BooleanVar(value=False)
cond_empty_cb = tk.Checkbutton(control_frame, text=get_text("cond_empty"), variable=cond_empty)
cond_empty_cb.pack(anchor="w", padx=10)

invalid_val_label = tk.Label(control_frame, text=get_text("invalid_values_title"))
invalid_val_label.pack(pady=(5, 0))
invalid_inner_frame = tk.Frame(control_frame)
invalid_inner_frame.pack(pady=5, padx=10, fill='x')
# load_invalid_values() دیگر در اینجا اجرا نمی‌شود و به رویداد ComboBox متصل است.

# ۴. ستون‌های جستجو
search_col_label = tk.Label(control_frame, text=get_text("search_cols"))
search_col_label.pack(pady=(10, 0))
search_inner_frame = tk.Frame(control_frame)
search_inner_frame.pack(pady=5, padx=10, fill='x')
load_search_columns()

# ۵. بارگذاری نقشه تثبیت (جدید)
load_map_label = tk.Label(control_frame, text=get_text("load_map_title"))
load_map_label.pack(pady=(10, 0))
load_map_btn = tk.Button(control_frame, text=get_text("select_map_file"), command=load_consolidation_map, bg="#E6EE9C")
load_map_btn.pack(pady=5, padx=10)
clear_map_btn = tk.Button(control_frame, text=get_text("clear_map"), command=clear_consolidation_map, bg="#FFCDD2")
clear_map_btn.pack(pady=5, padx=10)


# ۶. ستون‌های نمایش 
display_col_label = tk.Label(control_frame, text=get_text("display_cols"))
display_col_label.pack(pady=(10, 0))
display_inner_frame = tk.Frame(control_frame)
display_inner_frame.pack(pady=5, padx=10, fill='x')
load_display_columns()

# ---------------- پنل سمت راست (خروجی و جدول) ----------------
output_frame = tk.Frame(main_horizontal_frame)
output_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=10, pady=10)

# دکمه پیدا کردن ردیف‌ها
find_rows_btn = tk.Button(output_frame, text=get_text("find_rows"), command=find_rows, bg="#FFCC80", relief=tk.RAISED)
find_rows_btn.pack(pady=5)

# دکمه انتخاب همه
select_all_btn = tk.Button(output_frame, text=get_text("select_all"), command=select_all_rows, bg="#B3E5FC")
select_all_btn.pack(pady=2)

# نمایش تعداد ردیف‌ها (آپدیت شده)
count_label = tk.Label(output_frame, text=get_text("row_count").format(0, 0), font=("Arial", 10, "bold"))
count_label.pack(pady=5)

# ================= Treeview و Scrollbar =================
tree_frame = tk.Frame(output_frame)
tree_frame.pack(expand=True, fill="both", pady=5) 

tree = ttk.Treeview(tree_frame, selectmode="extended", show='headings') 
update_treeview_columns([])

tree_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)

tree_scrollbar_y.pack(side="right", fill="y")
tree.pack(side="top", expand=True, fill="both")
tree_scrollbar_x.pack(side="bottom", fill="x")

# ---------------- منوی کلیک راست و اشکال‌زدایی ----------------
context_menu = tk.Menu(root, tearoff=0)

tree.bind("<Motion>", show_tooltip)
tree.bind("<Leave>", hide_tooltip)
tree.bind("<Button-3>", show_context_menu) 
tree.bind("<Button-2>", show_context_menu) 
tree.bind("<Double-1>", open_editor) 


# دکمه نهایی و ادغام شده
apply_save_btn = tk.Button(output_frame, text=get_text("apply_save"), command=apply_fixes_and_save, bg="#81C784", fg="black", font=("Arial", 11, "bold"))
apply_save_btn.pack(pady=10)

root.bind('<Control-c>', lambda event: copy_selected_content(tree.focus(), tree.column(tree.identify_column(event.x), 'id') if tree.identify_column(event.x) else None)) 
root.bind('<Command-c>', lambda event: copy_selected_content(tree.focus(), tree.column(tree.identify_column(event.x), 'id') if tree.identify_column(event.x) else None)) 

# Initial text update (to ensure all texts are set based on the default 'en')
update_texts()

root.mainloop()
