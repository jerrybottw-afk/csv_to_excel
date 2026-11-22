import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# --- 色彩配置常數（方便集中調整 UI 配色） ---
BG_MAIN = "#F4F7FB"
BG_CARD = "#FFFFFF"
BG_HEADER = "#1E3A8A"
FG_HEADER = "#FFFFFF"
BG_ACCENT_SOFT = "#EFF6FF"
BORDER_COLOR = "#CBD5E1"
TEXT_PRIMARY = "#0F172A"
TEXT_MUTED = "#475569"

BTN_PRIMARY_BG = "#059669"
BTN_PRIMARY_FG = "#FFFFFF"
BTN_BLUE_BG = "#2563EB"
BTN_BLUE_FG = "#FFFFFF"
BTN_LIGHT_BG = "#E2E8F0"
BTN_LIGHT_FG = "#1E293B"

FONT_TITLE = ("Microsoft JhengHei UI", 13, "bold")
FONT_SUB = ("Microsoft JhengHei UI", 9)
FONT_UI = ("Microsoft JhengHei UI", 10)
FONT_BTN = ("Microsoft JhengHei UI", 10, "bold")


# 存放目前清單項目的狀態：[(檔名, BooleanVar), ...]
csv_items = []


def get_output_dir():
    """取得執行檔所在目錄下的 output 資料夾路徑，若不存在則自動建立"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base_dir, "output")
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def read_csv_safely(csv_path):
    """嘗試以常見中文編碼讀取 CSV，並將所有欄位保留為字串避免前導零遺失"""
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5"]
    last_error = None

    for enc in encodings:
        try:
            df = pd.read_csv(csv_path, encoding=enc, dtype=str, keep_default_na=False)
            for col in df.columns:
                df[col] = df[col].str.strip()
            return df
        except Exception as e:
            last_error = e

    raise last_error


def make_valid_sheet_name(raw_name, used_names):
    """清理 Excel 分頁名稱（移除非法字元、限制 31 字元、避免重複）"""
    cleaned = re.sub(r"[:\\/?*\[\]]", "_", raw_name).strip() or "Sheet"
    base = cleaned[:31]
    sheet_name = base
    counter = 1

    while sheet_name in used_names:
        suffix = f"_{counter}"
        sheet_name = base[: 31 - len(suffix)] + suffix
        counter += 1

    used_names.add(sheet_name)
    return sheet_name


def auto_fit_columns(ws, padding=4, max_width=60):
    """根據工作表內容字元長度自動適配欄寬（中文字算 2 單位）"""
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = 0
        col_letter = get_column_letter(col_idx)

        for cell in col:
            text = "" if cell.value is None else str(cell.value)
            width_units = sum(2 if ord(ch) > 127 else 1 for ch in text)
            if width_units > max_len:
                max_len = width_units

        fitted_width = min(max_len + padding, max_width)
        ws.column_dimensions[col_letter].width = max(fitted_width, 10)


def convert_single_csv(csv_path, excel_path):
    """獨立模式：將單一 CSV 轉為 Excel 並自動調整欄寬"""
    df = read_csv_safely(csv_path)
    df.to_excel(excel_path, index=False, sheet_name="Sheet1")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Sheet1"]
    ws.freeze_panes = "A2"
    auto_fit_columns(ws)
    wb.save(excel_path)


def convert_merged_csvs(folder_path, selected_files, output_excel_path):
    """合併模式：將多個 CSV 寫入同一個 Excel 的不同分頁，並調整所有分頁欄寬"""
    used_sheet_names = set()
    failed_files = []
    success_count = 0

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for filename in selected_files:
            csv_path = os.path.join(folder_path, filename)
            base_name = os.path.splitext(filename)[0]
            sheet_name = make_valid_sheet_name(base_name, used_sheet_names)

            try:
                df = read_csv_safely(csv_path)
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                success_count += 1
            except Exception as e:
                failed_files.append(f"{filename}：{e}")

    if success_count == 0:
        if os.path.exists(output_excel_path):
            os.remove(output_excel_path)
        return 0, failed_files

    wb = openpyxl.load_workbook(output_excel_path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        auto_fit_columns(ws)
    wb.save(output_excel_path)

    return success_count, failed_files


def clear_checkbox_list():
    """清空目前的勾選清單畫面與資料"""
    for widget in list_inner_frame.winfo_children():
        widget.destroy()
    csv_items.clear()


def on_canvas_configure(event):
    """讓內部勾選區塊寬度跟隨捲動區"""
    list_canvas.itemconfig(canvas_window_id, width=event.width)


def on_mousewheel(event):
    """支援滑鼠滾輪捲動檔案清單"""
    list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def load_csv_files(folder_path):
    """讀取指定目錄下的所有 .csv 檔案並建立勾選項目"""
    clear_checkbox_list()

    if not folder_path or not os.path.isdir(folder_path):
        status_var.set("請選擇有效的 CSV 目錄（輸出：./output）")
        return

    files = [
        f
        for f in sorted(os.listdir(folder_path))
        if f.lower().endswith(".csv") and os.path.isfile(os.path.join(folder_path, f))
    ]

    if not files:
        status_var.set("此目錄下找不到任何 .csv 檔案（輸出：./output）")
        return

    for idx, filename in enumerate(files):
        row_bg = BG_CARD if idx % 2 == 0 else BG_ACCENT_SOFT
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(
            list_inner_frame,
            text=f"  {filename}",
            variable=var,
            anchor="w",
            bg=row_bg,
            activebackground="#DBEAFE",
            fg=TEXT_PRIMARY,
            activeforeground=TEXT_PRIMARY,
            selectcolor=BG_CARD,
            font=FONT_UI,
            padx=8,
            pady=4,
            bd=0,
            highlightthickness=0,
        )
        chk.pack(fill="x", anchor="w")
        csv_items.append((filename, var))

    status_var.set(f"已載入 {len(files)} 個 CSV 檔案（預設全選｜輸出：./output）")


def browse_folder():
    """開啟資料夾選擇視窗"""
    selected_dir = filedialog.askdirectory(title="選擇 CSV 目錄")
    if selected_dir:
        folder_var.set(selected_dir)
        load_csv_files(selected_dir)


def select_all(state):
    """全選或取消全選"""
    for _, var in csv_items:
        var.set(state)


def start_convert():
    """執行已勾選檔案的批次轉換"""
    folder_path = folder_var.get().strip()
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showwarning("提示", "請先選擇有效的 CSV 目錄。")
        return

    selected_files = [filename for filename, var in csv_items if var.get()]
    if not selected_files:
        messagebox.showwarning("提示", "請至少勾選一個 CSV 檔案。")
        return

    out_dir = get_output_dir()
    mode = mode_var.get()

    if mode == "merge":
        merged_path = os.path.join(out_dir, "merged_output.xlsx")
        try:
            success_count, failed_files = convert_merged_csvs(
                folder_path, selected_files, merged_path
            )
        except Exception as e:
            messagebox.showerror(
                "轉換失敗", f"寫入 merged_output.xlsx 時發生錯誤：\n{e}"
            )
            return

        status_var.set(
            f"合併完成：成功 {success_count} 分頁，失敗 {len(failed_files)} 筆"
        )

        if failed_files:
            messagebox.showerror(
                "部分分頁轉換失敗",
                f"已輸出至：{merged_path}\n成功 {success_count} 個分頁，以下檔案發生錯誤：\n\n"
                + "\n".join(failed_files),
            )
        else:
            messagebox.showinfo(
                "完成", f"已合併 {success_count} 個 CSV 分頁至：\n{merged_path}"
            )

    else:
        success_count = 0
        failed_files = []

        for filename in selected_files:
            csv_path = os.path.join(folder_path, filename)
            base_name = os.path.splitext(filename)[0]
            excel_path = os.path.join(out_dir, f"{base_name}.xlsx")

            try:
                convert_single_csv(csv_path, excel_path)
                success_count += 1
            except Exception as e:
                failed_files.append(f"{filename}：{e}")

        status_var.set(
            f"獨立完成：成功 {success_count} 筆，失敗 {len(failed_files)} 筆"
        )

        if failed_files:
            messagebox.showerror(
                "部分轉換失敗",
                f"輸出目錄：{out_dir}\n成功 {success_count} 個檔案，以下檔案發生錯誤：\n\n"
                + "\n".join(failed_files),
            )
        else:
            messagebox.showinfo(
                "完成", f"已成功將 {success_count} 個檔案輸出至：\n{out_dir}"
            )


# --- 建立主視窗與介面元件（程序化佈局 + Colorful 配色） ---
root = tk.Tk()
root.title("CSV 批次轉 Excel 工具")
root.geometry("680x540")
root.configure(bg=BG_MAIN)

# 頂部色彩標題列
banner = tk.Frame(root, bg=BG_HEADER, padx=16, pady=12)
banner.pack(fill="x")

tk.Label(
    banner,
    text="CSV 批次轉 Excel 工具",
    bg=BG_HEADER,
    fg=FG_HEADER,
    font=FONT_TITLE,
    anchor="w",
).pack(side="left")

tk.Label(
    banner,
    text="自動欄寬適配・預設合併分頁輸出至 ./output",
    bg=BG_HEADER,
    fg="#BFDBFE",
    font=FONT_SUB,
    anchor="e",
).pack(side="right")

# 1. 目錄選擇區
top_frame = tk.Frame(root, bg=BG_MAIN, padx=14, pady=10)
top_frame.pack(fill="x")

tk.Label(top_frame, text="CSV 目錄：", bg=BG_MAIN, fg=TEXT_PRIMARY, font=FONT_UI).pack(
    side="left"
)

folder_var = tk.StringVar()
folder_entry = tk.Entry(
    top_frame,
    textvariable=folder_var,
    bg=BG_CARD,
    fg=TEXT_PRIMARY,
    font=FONT_UI,
    relief="solid",
    bd=1,
    highlightthickness=1,
    highlightcolor=BTN_BLUE_BG,
)
folder_entry.pack(side="left", fill="x", expand=True, padx=8, ipady=3)

tk.Button(
    top_frame,
    text="瀏覽目錄",
    command=browse_folder,
    bg=BTN_BLUE_BG,
    fg=BTN_BLUE_FG,
    activebackground="#1D4ED8",
    activeforeground=BTN_BLUE_FG,
    font=FONT_BTN,
    relief="flat",
    padx=10,
    pady=4,
    cursor="hand2",
).pack(side="left", padx=(0, 6))

tk.Button(
    top_frame,
    text="重新載入",
    command=lambda: load_csv_files(folder_var.get().strip()),
    bg=BTN_LIGHT_BG,
    fg=BTN_LIGHT_FG,
    activebackground="#CBD5E1",
    activeforeground=BTN_LIGHT_FG,
    font=FONT_UI,
    relief="flat",
    padx=10,
    pady=4,
    cursor="hand2",
).pack(side="left")

# 2. 輸出模式選擇區（預設合併模式）
mode_frame = tk.Frame(
    root,
    bg=BG_ACCENT_SOFT,
    padx=14,
    pady=8,
    highlightbackground="#BFDBFE",
    highlightthickness=1,
)
mode_frame.pack(fill="x", padx=14, pady=(0, 8))

tk.Label(
    mode_frame, text="輸出模式：", bg=BG_ACCENT_SOFT, fg=BG_HEADER, font=FONT_BTN
).pack(side="left")
mode_var = tk.StringVar(value="merge")

tk.Radiobutton(
    mode_frame,
    text="合併模式（單一 merged_output.xlsx，一檔一分頁）",
    variable=mode_var,
    value="merge",
    bg=BG_ACCENT_SOFT,
    activebackground=BG_ACCENT_SOFT,
    fg=TEXT_PRIMARY,
    selectcolor=BG_CARD,
    font=FONT_UI,
).pack(side="left", padx=(6, 14))

tk.Radiobutton(
    mode_frame,
    text="獨立模式（各自轉成獨立 Excel）",
    variable=mode_var,
    value="single",
    bg=BG_ACCENT_SOFT,
    activebackground=BG_ACCENT_SOFT,
    fg=TEXT_PRIMARY,
    selectcolor=BG_CARD,
    font=FONT_UI,
).pack(side="left")

# 3. 全選 / 取消全選工具列
action_bar = tk.Frame(root, bg=BG_MAIN, padx=14, pady=4)
action_bar.pack(fill="x")

tk.Label(
    action_bar, text="檔案清單（可勾選）：", bg=BG_MAIN, fg=TEXT_PRIMARY, font=FONT_BTN
).pack(side="left")

tk.Button(
    action_bar,
    text="取消全選",
    command=lambda: select_all(False),
    bg=BTN_LIGHT_BG,
    fg=BTN_LIGHT_FG,
    activebackground="#CBD5E1",
    activeforeground=BTN_LIGHT_FG,
    font=FONT_SUB,
    relief="flat",
    padx=8,
    pady=2,
    cursor="hand2",
).pack(side="right", padx=(6, 0))

tk.Button(
    action_bar,
    text="全選",
    command=lambda: select_all(True),
    bg="#DBEAFE",
    fg=BG_HEADER,
    activebackground="#BFDBFE",
    activeforeground=BG_HEADER,
    font=FONT_SUB,
    relief="flat",
    padx=10,
    pady=2,
    cursor="hand2",
).pack(side="right")

# 4. 可捲動勾選清單區
list_container = tk.Frame(
    root,
    bg=BG_CARD,
    padx=4,
    pady=4,
    highlightbackground=BORDER_COLOR,
    highlightthickness=1,
)
list_container.pack(fill="both", expand=True, padx=14, pady=(0, 8))

list_canvas = tk.Canvas(list_container, bg=BG_CARD, highlightthickness=0)
scrollbar = tk.Scrollbar(list_container, orient="vertical", command=list_canvas.yview)
list_inner_frame = tk.Frame(list_canvas, bg=BG_CARD)

canvas_window_id = list_canvas.create_window(
    (0, 0), window=list_inner_frame, anchor="nw"
)
list_canvas.configure(yscrollcommand=scrollbar.set)

list_canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

list_inner_frame.bind(
    "<Configure>", lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all"))
)
list_canvas.bind("<Configure>", on_canvas_configure)
list_canvas.bind_all("<MouseWheel>", on_mousewheel)

# 5. 底部轉換按鈕與狀態列
bottom_frame = tk.Frame(root, bg=BG_MAIN, padx=14, pady=10)
bottom_frame.pack(fill="x")

tk.Button(
    bottom_frame,
    text="一鍵轉換成 Excel",
    command=start_convert,
    bg=BTN_PRIMARY_BG,
    fg=BTN_PRIMARY_FG,
    activebackground="#047857",
    activeforeground=BTN_PRIMARY_FG,
    font=FONT_BTN,
    relief="flat",
    padx=18,
    pady=8,
    cursor="hand2",
).pack(side="right")

status_var = tk.StringVar(value="尚未選擇目錄（輸出位置：./output）")
tk.Label(
    bottom_frame,
    textvariable=status_var,
    bg=BG_MAIN,
    fg=TEXT_MUTED,
    font=FONT_SUB,
    anchor="w",
).pack(side="left", fill="x", expand=True)

root.mainloop()
