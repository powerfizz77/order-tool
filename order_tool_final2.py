# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# GUI
import tkinter as tk
from tkinter import filedialog, messagebox

# ===================== 跨平台打开文件（Windows/Mac通用）=====================
def open_file(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception:
        messagebox.showinfo("完成", f"文件已生成：\n{path}\n请手动打开")

# ===================== 核心统计逻辑 =====================
def auto_statistics(filepath):
    df = pd.read_excel(filepath)

    col_map = {}
    for col in df.columns:
        c = str(col).strip()
        if '店铺' in c:
            col_map['shop'] = col
        elif '原始单号' in c:
            col_map['order_id'] = col
        elif '货品成交总价' in c:
            col_map['sales'] = col
        elif '固定总成本' in c:
            col_map['cost'] = col
        elif '预估邮资' in c or '邮资' in c or '快递费' in c:
            col_map['express'] = col
        elif '服务费' in c:
            col_map['service'] = col

    required = ['shop', 'order_id', 'sales', 'cost', 'express', 'service']
    for k in required:
        if k not in col_map:
            raise Exception(f"缺少必要列：{k}")

    df_use = df.rename(columns={
        col_map['shop']: '店铺',
        col_map['order_id']: '原始单号',
        col_map['sales']: '货品成交总价',
        col_map['cost']: '固定总成本',
        col_map['express']: '快递费',
        col_map['service']: '服务费',
    })[['店铺', '原始单号', '货品成交总价', '固定总成本', '快递费', '服务费']]

    df_use = df_use.dropna(subset=['店铺', '原始单号'])
    num_cols = ['货品成交总价', '固定总成本', '快递费', '服务费']
    for c in num_cols:
        df_use[c] = pd.to_numeric(df_use[c], errors='coerce').fillna(0)

    money_sum = df_use.groupby('店铺')[num_cols].sum().reset_index()
    unique_orders = df_use[['店铺', '原始单号']].drop_duplicates()
    ship_count = unique_orders.groupby('店铺').size().reset_index(name='发件总票数')

    result = pd.merge(money_sum, ship_count, on='店铺', how='left')
    result['发件总票数'] = result['发件总票数'].fillna(0).astype(int)
    result = result.round(2)
    result = result.sort_values('货品成交总价', ascending=False).reset_index(drop=True)
    return result

# ===================== GUI 主界面（跨平台美化）=====================
def main_gui():
    root = tk.Tk()
    root.title('订单统计工具 · 跨平台版')
    root.geometry('620x340')
    root.resizable(False, False)

    # 跨平台字体
    if sys.platform == "darwin":
        default_font = ("PingFang SC", 11)
        bold_font = ("PingFang SC", 18, "bold")
        small_font = ("PingFang SC", 9)
    else:
        default_font = ("Microsoft YaHei", 11)
        bold_font = ("Microsoft YaHei", 18, "bold")
        small_font = ("Microsoft YaHei", 9)

    # 标题
    tk.Label(root, text='📊 订单店铺统计工具', font=bold_font).pack(pady=18)

    # 说明
    tk.Label(root, text='自动统计：成交总价 | 总成本 | 邮资 | 服务费 | 去重发件票数',
             font=default_font).pack(pady=2)

    # 选择文件
    def select_file():
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, path)

    # 执行统计
    def run_task():
        fpath = entry_file.get().strip()
        if not fpath or not os.path.isfile(fpath):
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return

        try:
            btn_start.config(text="处理中...", state=tk.DISABLED, bg="#666")
            root.update()

            res = auto_statistics(fpath)
            out_dir = os.path.dirname(fpath)
            out_file = os.path.join(out_dir, "店铺统计结果.xlsx")
            res.to_excel(out_file, index=False)

            messagebox.showinfo("完成", "✅ 统计完成！")
            open_file(out_file)

        except Exception as e:
            messagebox.showerror("处理失败", str(e))
        finally:
            btn_start.config(text="开始统计", state=tk.NORMAL, bg="#2E86AB")

    # 文件路径框
    entry_file = tk.Entry(root, font=default_font, width=60)
    entry_file.pack(pady=12)

    # 按钮区域
    frm = tk.Frame(root)
    frm.pack(pady=12)

    # 修复：-pad 改成 -padx
    tk.Button(frm, text="选择Excel文件", font=default_font, width=16,
              command=select_file).grid(row=0, column=0, padx=10)

    btn_start = tk.Button(frm, text="开始统计", font=(default_font[0], default_font[1], "bold"),
                          width=16, bg="#2E86AB", fg="white", command=run_task)
    btn_start.grid(row=0, column=1, padx=10)

    # 底部提示
    tk.Label(root, text='规则：按【店铺】汇总 | 【原始单号】自动去重 | 支持任意列顺序',
             font=small_font, fg="#666").pack(pady=10)

    root.mainloop()

if __name__ == '__main__':
    main_gui()