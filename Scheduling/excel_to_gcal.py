# ...existing code...
import argparse
import pandas as pd
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
##python 檔案名稱 --file 檔案路徑 --sheet 工作表名稱 --name 篩選人名 --out 輸出檔名
def find_header_row(df):
    # 找出包含 "Date" 或 "日期" 的列位置（優先偵測英文字 "Date"）
    indices = []
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.lower().fillna('')
        if any('date' in c for c in row) or any('日期' in c for c in row): indices.append(i)
    return indices if indices else None

def build_column_labels(df, header1_idx, header2_idx):
    ncols = df.shape[1]
    col_labels = []
    is_date = []
    for c in range(ncols):
        top = df.iat[header1_idx, c]
        bot = df.iat[header2_idx, c]
        dt = pd.to_datetime(top, errors='coerce')
        if not pd.isna(dt):
            col_labels.append(dt.strftime('%Y-%m-%d'))
            is_date.append(True)
        else:
            label = None
            if isinstance(bot, str) and bot.strip():
                label = bot.strip()
            elif isinstance(top, str) and top.strip():
                label = top.strip()
            else:
                label = f"col_{c}"
            col_labels.append(label)
            is_date.append(False)
    return col_labels, is_date

def find_col_by_candidates(labels, candidates):
    lower = [str(l).strip().lower() for l in labels]
    for cand in candidates:
        cand = cand.lower()
        for i, lab in enumerate(lower):
            if lab == cand or cand in lab:
                return labels[i]
    return None

def convert_to_gcal(df_raw, sheet_name=None, name_filter=None, out_csv='google_calendar.csv'):
    header_positions = find_header_row(df_raw)
    if not header_positions:
        raise RuntimeError("無法找到標題列（找不到 'Date' / '日期'）。請檢查 Excel 格式。")
    
    all_events = []
    # 針對每個找到的日期標題區段進行處理
    for i,header1 in enumerate(header_positions):
        # 計算本區段的結束位置
        next_header = header_positions[i + 1] if i + 1 < len(header_positions) else len(df_raw)
        header2 = header1 + 1

        # 建立此區段的欄位標籤
        col_labels, is_date = build_column_labels(df_raw, header1, header2)

        # 只取這個區段的資料
        data = df_raw.iloc[header2+1:next_header].copy()
        data.columns = col_labels

        # 找出此區段的成員欄和課程欄
        member_col = find_col_by_candidates(col_labels, ['member','姓名','成員','名字','name'])
        course_col = find_col_by_candidates(col_labels, ['課別','course','科別'])
        if member_col is None:
            # 嘗試推定為第二欄或第一個非日期欄
            non_date = [col_labels[i] for i,v in enumerate(is_date) if not v]
            if not non_date:
                raise RuntimeError("無法找到成員欄位。")
            member_col = non_date[1] if len(non_date) > 1 else non_date[0]
        if course_col is None:
            non_date = [col_labels[i] for i,v in enumerate(is_date) if not v]
            try:
                idx = non_date.index(member_col)
                course_col = non_date[idx+1] if idx+1 < len(non_date) else None
            except ValueError:
                course_col = non_date[1] if len(non_date) > 1 else None

        # 取得日期欄位名稱列表
        date_cols = [col_labels[i] for i,v in enumerate(is_date) if v]
        for _, row in data.iterrows():
            member = row.get(member_col, '')
            course = row.get(course_col, '')
            if pd.isna(member) or str(member).strip() == '':
                continue
            if name_filter and name_filter.strip() not in str(member):
                continue
            for dcol in date_cols:
                val = row.get(dcol, '')
                if pd.isna(val) or str(val).strip() == '':
                    continue
                val_s = str(val).strip()
                if '休' in val_s:
                    continue
                first = val_s[0].upper() if val_s else ''
                # 解析日期
                try:
                    start_dt = pd.to_datetime(dcol)
                except Exception:
                    start_dt = pd.to_datetime(str(dcol))
                start_date = start_dt.date()
                # 預設為全天事件
                start_time = ''
                end_time = ''
                end_date = start_date
                all_day = 'TRUE'
                description = val_s
                # 如果開頭是 D 或 N，填入時間並把 All Day 設為 FALSE
                if first == 'D':
                    start_time = '07:30'
                    end_time = '19:30'
                    all_day = 'FALSE'
                    description = val_s+' '+'07:30~19:30'
                elif first == 'N':
                    start_time = '19:30'
                    end_time = '07:30'
                    all_day = 'FALSE'
                    description = val_s+' '+'19:30~07:30'
                    end_date = (pd.to_datetime(start_date) + pd.Timedelta(days=1)).date()
                # 建事件
                all_events.append({
                    'Subject': f"{str(member).strip()} - {str(course).strip()}",
                    'Start Date': start_date.strftime('%Y-%m-%d'),
                    'Start Time': start_time,
                    'End Date': end_date.strftime('%Y-%m-%d'),
                    'End Time': end_time,
                    'All Day Event': all_day,
                    'Description': description,
                    'Location': '',
                    'Private': 'TRUE'
                })

    if not all_events:
        return 0

    out_df = pd.DataFrame(all_events, columns=[
        'Subject','Start Date','Start Time','End Date','End Time',
        'All Day Event','Description','Location','Private'
    ])
    out_df.to_csv(out_csv, index=False, encoding='utf-8-sig')
    return len(out_df)

# GUI 部分
def launch_gui():
    root = tk.Tk()
    root.title("Excel -> Google Calendar CSV")
    root.geometry("520x220")
    root.resizable(False, False)

    file_var = tk.StringVar()
    sheet_var = tk.StringVar(value='各組值班表')
    name_var = tk.StringVar()
    out_var = tk.StringVar(value='google_calendar.csv')

    def choose_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            file_var.set(path)

    def do_convert():
        file = file_var.get().strip()
        sheet = sheet_var.get().strip()
        name = name_var.get().strip()
        out = out_var.get().strip() or 'google_calendar.csv'
        if not file:
            messagebox.showwarning("缺少檔案", "請選擇 Excel 檔案。")
            return
        try:
            df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        except Exception as e:
            messagebox.showerror("讀取失敗", f"讀取 Excel 失敗：{e}")
            return
        try:
            count = convert_to_gcal(df_raw, sheet_name=sheet, name_filter=name if name else None, out_csv=out)
        except Exception as e:
            messagebox.showerror("轉換錯誤", f"轉換失敗：{e}")
            return
        if count == 0:
            messagebox.showinfo("完成", "未找到任何事件（可能篩選條件或資料格式）。")
        else:
            messagebox.showinfo("完成", f"成功輸出 {count} 筆事件到：{out}")

    # 介面佈局
    tk.Label(root, text="Excel 檔案：").place(x=10, y=14)
    tk.Entry(root, textvariable=file_var, width=54).place(x=90, y=12)
    tk.Button(root, text="選擇檔案", command=choose_file).place(x=430, y=8)

    tk.Label(root, text="工作表名稱：").place(x=10, y=50)
    tk.Entry(root, textvariable=sheet_var, width=20).place(x=110, y=48)

    tk.Label(root, text="篩選人名（可留空）：").place(x=10, y=86)
    tk.Entry(root, textvariable=name_var, width=30).place(x=140, y=84)

    tk.Label(root, text="輸出檔名：").place(x=10, y=122)
    tk.Entry(root, textvariable=out_var, width=30).place(x=110, y=120)

    tk.Button(root, text="轉換", width=12, command=do_convert).place(x=200, y=160)

    root.mainloop()

def main():
    parser = argparse.ArgumentParser(description='將值班表 Excel 轉成 Google Calendar 匯入 CSV')
    parser.add_argument('--file', '-f', default=None, help='Excel 檔案路徑')
    parser.add_argument('--sheet', '-s', default='各組值班表', help='工作表名稱')
    parser.add_argument('--name', '-n', default=None, help='篩選特定人名（完全或部分匹配）')
    parser.add_argument('--out', '-o', default='google_calendar.csv', help='輸出 CSV 檔名')
    parser.add_argument('--gui', action='store_true', help='啟動圖形介面')
    args = parser.parse_args()

    if args.gui:
        launch_gui()
        return

    if not args.file:
        print("未指定檔案。可使用 --gui 啟動視窗選擇檔案。")
        sys.exit(1)

    try:
        df_raw = pd.read_excel(args.file, sheet_name=args.sheet, header=None)
    except Exception as e:
        print("讀取 Excel 失敗：", e)
        sys.exit(1)

    try:
        count = convert_to_gcal(df_raw, sheet_name=args.sheet, name_filter=args.name, out_csv=args.out)
    except Exception as e:
        print("轉換失敗：", e)
        sys.exit(1)

    if count == 0:
        print("沒有找到符合條件的事件。")
    else:
        print(f"已輸出 {count} 筆事件到：{args.out}")

if __name__ == '__main__':
    main()
# ...existing code...