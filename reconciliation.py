import os
import re
import datetime
import threading
from decimal import Decimal, InvalidOperation

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import openpyxl

# ==================== 业务配置常量 ====================
PR_KEYWORDS = ["预收款"]
PR_HEADER_ROW = 3
PR_DATA_START_ROW = 4

# 三种模式的匹配配置
MODE_CONFIGS = {
    "银行": {
        "bank_keywords": ["银行", "流水", "对账单"],
        "bank_header_row": 16,
        "bank_data_start_row": 17,
        "pr_settle_account": "中信银行5031",
        "bank_cols": {
            "amount": "贷方发生额",
            "date": "交易日期",
            "opp_name": "对方账户名称",
            "currency": "币种"
        },
        "output_cols": {
            15: ("单据编号", "billno"),  # O 列
            16: ("客户", "client"),      # P 列
            17: ("订单号", "order")       # Q 列
        }
    },
    "支付宝": {
        "bank_keywords": ["支付宝", "alipay"],
        "bank_header_row": 3,
        "bank_data_start_row": 4,
        "pr_settle_account": "支付宝",
        "bank_cols": {
            "amount": "收入（+元）",
            "date": "入账时间"
        },
        "output_cols": {
            23: ("单据编号", "billno"),  # W 列
            25: ("客户", "client"),      # Y 列
            26: ("订单号", "order"),     # Z 列
            27: ("收款备注", "comment")   # AA 列
        }
    },
    "微信": {
        "bank_keywords": ["微信", "wechat"],
        "bank_header_row": 1,
        "bank_data_start_row": 2,
        "pr_settle_account": "宇问测量微信",
        "bank_cols": {
            "amount": "实收金额（元）",
            "date": "?交易时间"
        },
        "output_cols": {
            39: ("单据编号", "billno"),  # AM 列
            41: ("客户", "client"),      # AO 列
            42: ("订单号", "order")       # AP 列
        }
    }
}
# ====================================================

class ReconciliationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("银行/支付宝/微信流水自动核对工具")
        self.geometry("750x660")
        self.minsize(650, 550)
        
        # 绑定的变量
        self.mode_var = tk.StringVar(value="银行")
        self.bank_file_path = tk.StringVar()
        self.pr_file_path = tk.StringVar()
        self.status_text = tk.StringVar(value="准备就绪")
        
        self.init_ui()
        self.auto_detect_files()

    def init_ui(self):
        # 1. 模式选择区域
        mode_frame = ttk.LabelFrame(self, text="请选择对账模式", padding=10)
        mode_frame.pack(fill=tk.X, padx=15, pady=8)
        
        ttk.Radiobutton(mode_frame, text=" 银行流水对账模式 ", variable=self.mode_var, value="银行", command=self.on_mode_change).pack(side=tk.LEFT, padx=15)
        ttk.Radiobutton(mode_frame, text=" 支付宝对账模式 ", variable=self.mode_var, value="支付宝", command=self.on_mode_change).pack(side=tk.LEFT, padx=15)
        ttk.Radiobutton(mode_frame, text=" 微信对账模式 ", variable=self.mode_var, value="微信", command=self.on_mode_change).pack(side=tk.LEFT, padx=15)

        # 2. 文件选择区域
        file_frame = ttk.LabelFrame(self, text="数据源选择", padding=10)
        file_frame.pack(fill=tk.X, padx=15, pady=5)

        # 流水文件
        self.bank_label = ttk.Label(file_frame, text="流水文件:")
        self.bank_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        self.bank_entry = ttk.Entry(file_frame, textvariable=self.bank_file_path, width=55)
        self.bank_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="选择...", command=self.browse_bank_file).grid(row=0, column=2, padx=5, pady=5)

        # 预收款文件
        ttk.Label(file_frame, text="预收款文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.pr_entry = ttk.Entry(file_frame, textvariable=self.pr_file_path, width=55)
        self.pr_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="选择...", command=self.browse_pr_file).grid(row=1, column=2, padx=5, pady=5)

        file_frame.columnconfigure(1, weight=1)
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.E)
        ttk.Button(btn_frame, text="重新自动识别当前目录文件", command=self.auto_detect_files).pack(side=tk.RIGHT)

        # 3. 控制与状态区域
        control_frame = ttk.Frame(self, padding=5)
        control_frame.pack(fill=tk.X, padx=15, pady=5)

        self.run_btn = ttk.Button(control_frame, text="开始执行对账核对", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, ipadx=15)

        self.status_label = ttk.Label(control_frame, textvariable=self.status_text, font=("Arial", 10, "bold"), foreground="blue")
        self.status_label.pack(side=tk.RIGHT, padx=10)

        # 4. 日志输出区域
        log_frame = ttk.LabelFrame(self, text="处理日志输出", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        self.log_area = ScrolledText(log_frame, wrap=tk.WORD, height=13)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.configure(state='disabled')

    def log(self, message):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def clear_log(self):
        self.log_area.configure(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state='disabled')

    def on_mode_change(self):
        mode = self.mode_var.get()
        self.bank_label.config(text=f"{mode}流水文件:")
        self.log(f"[模式切换] 当前对账模式切换为: 【{mode}】")
        self.auto_detect_files()

    def auto_detect_files(self):
        current_dir = os.getcwd()
        mode = self.mode_var.get()
        cfg = MODE_CONFIGS[mode]

        bank_file = self.find_file_by_keywords(current_dir, cfg["bank_keywords"])
        pr_file = self.find_file_by_keywords(current_dir, PR_KEYWORDS)
        
        if bank_file:
            self.bank_file_path.set(bank_file)
            self.log(f"[系统自动识别] 已定位到【{mode}】流水文件: {os.path.basename(bank_file)}")
        else:
            self.bank_file_path.set("")
            self.log(f"[提示] 未在当前目录自动识别到【{mode}】流水文件，请手动选择。")

        if pr_file:
            self.pr_file_path.set(pr_file)
            self.log(f"[系统自动识别] 已定位到预收款文件: {os.path.basename(pr_file)}")
        else:
            self.pr_file_path.set("")
            self.log("[提示] 未在当前目录自动识别到预收款文件，请手动选择。")

    def find_file_by_keywords(self, directory, keywords):
        for filename in os.listdir(directory):
            if not filename.endswith('.xlsx') or filename.startswith('~$'):
                continue
            if any(kw in filename for kw in keywords):
                return os.path.join(directory, filename)
        return None

    def browse_bank_file(self):
        mode = self.mode_var.get()
        file_selected = filedialog.askopenfilename(
            title=f"选择{mode}流水Excel文件",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if file_selected:
            normalized_path = os.path.normpath(file_selected)
            self.bank_file_path.set(normalized_path)
            self.log(f"[手动选择] {mode}流水文件已更改为: {os.path.basename(normalized_path)}")

    def browse_pr_file(self):
        file_selected = filedialog.askopenfilename(
            title="选择预收款Excel文件",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if file_selected:
            normalized_path = os.path.normpath(file_selected)
            self.pr_file_path.set(normalized_path)
            self.log(f"[手动选择] 预收款文件已更改为: {os.path.basename(normalized_path)}")

    def clean_amount(self, val):
        if val is None:
            return None
        val_str = str(val).strip().replace(',', '')
        if not val_str:
            return None
        try:
            cleaned = re.sub(r'[^\d\.\-]', '', val_str)
            if not cleaned:
                return None
            return Decimal(cleaned)
        except (InvalidOperation, ValueError):
            return None

    def parse_date_only(self, val):
        if val is None:
            return None
        if isinstance(val, datetime.datetime):
            return val.date()
        if isinstance(val, datetime.date):
            return val
        val_str = str(val).strip()
        if not val_str:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(val_str, fmt).date()
            except ValueError:
                continue
        match = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', val_str)
        if match:
            try:
                return datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            except ValueError:
                pass
        return None

    def is_chinese_name(self, name):
        if not name:
            return False
        return bool(re.search(r'[\u4e00-\u9fa5]', str(name)))

    def find_header_col(self, headers, targets):
        for idx, h in enumerate(headers, start=1):
            h_clean = h.strip()
            for t in targets:
                if t.lower() in h_clean.lower() or h_clean.lower() == t.lower():
                    return idx
        return None

    def get_sheet_insensitive(self, workbook, target_name="sheet1"):
        for name in workbook.sheetnames:
            if name.lower() == target_name.lower():
                return workbook[name]
        return workbook.active

    # ==================== 合并单元格处理辅助工具 ====================
    def get_merged_cells_map(self, sheet):
        """扫描全表，构建合并单元格坐标到左上角真实数据值的映射字典"""
        merged_map = {}
        for merged_range in sheet.merged_cells.ranges:
            # 找到合并单元格左上角的真实数据值
            top_left_val = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
            # 将合并区域内所有坐标映射至该真实数据
            for r in range(merged_range.min_row, merged_range.max_row + 1):
                for c in range(merged_range.min_col, merged_range.max_col + 1):
                    merged_map[(r, c)] = top_left_val
        return merged_map

    def get_cell_value(self, sheet, row, col, merged_map):
        """支持合并单元格的安全取值方法"""
        if (row, col) in merged_map:
            return merged_map[(row, col)]
        return sheet.cell(row=row, column=col).value
    # ==============================================================

    def start_processing_thread(self):
        if not self.bank_file_path.get() or not self.pr_file_path.get():
            messagebox.showwarning("警告", "请确保已指定或选择了流水文件与预收款文件！")
            return
        
        self.run_btn.configure(state=tk.DISABLED)
        self.status_text.set("正在核对数据中...")
        self.clear_log()
        
        t = threading.Thread(target=self.process_reconciliation)
        t.daemon = True
        t.start()

    def process_reconciliation(self):
        start_time = datetime.datetime.now()
        log_messages = []

        def local_log(msg):
            self.log(msg)
            log_messages.append(msg)

        mode = self.mode_var.get()
        cfg = MODE_CONFIGS[mode]

        bank_path = self.bank_file_path.get()
        pr_path = self.pr_file_path.get()
        target_output_dir = os.path.dirname(bank_path)

        local_log(f"开始处理时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        local_log(f"当前对账模式: 【{mode}】")
        local_log(f"流水文件路径: {bank_path}")
        local_log(f"预收款路径: {pr_path}")
        local_log(f"结果输出目录: {target_output_dir}")

        # 1. 读取并预解析预收款文件
        try:
            pr_wb = openpyxl.load_workbook(pr_path, data_only=True)
            pr_ws = self.get_sheet_insensitive(pr_wb, "sheet1")
            # 解析预收款工作表的合并单元格
            pr_merged_map = self.get_merged_cells_map(pr_ws)
        except Exception as e:
            local_log(f"错误: 读取预收款文件失败: {str(e)}")
            self.finalize_run(False, log_messages, target_output_dir)
            return

        # 获取预收款表头（兼容表头可能包含在合并单元格中的情况）
        pr_headers = [
            str(self.get_cell_value(pr_ws, PR_HEADER_ROW, col, pr_merged_map)).strip()
            if self.get_cell_value(pr_ws, PR_HEADER_ROW, col, pr_merged_map) is not None else ""
            for col in range(1, pr_ws.max_column + 1)
        ]
        
        # 匹配预收款各列索引
        col_pr_settle = self.find_header_col(pr_headers, ["*收款账户.名称 # settleaccount.name"])
        col_pr_amount = self.find_header_col(pr_headers, ["*收款金额 # amount"])
        col_pr_date = self.find_header_col(pr_headers, ["*单据日期 # billdate"])
        col_pr_billno = self.find_header_col(pr_headers, ["*单据编号 # billno"])
        col_pr_client = self.find_header_col(pr_headers, ["*结算客户.名称 # itemclass.name"])
        col_pr_order = self.find_header_col(pr_headers, ["*订单编号 # srcbillno"])
        col_pr_exrate = self.find_header_col(pr_headers, ["*汇率 # exchangerate"])
        col_pr_currency = self.find_header_col(pr_headers, ["*币别.名称 # currency.name"])
        col_pr_comment = self.find_header_col(pr_headers, ["收款备注 # comment"])

        # 针对当前模式检查必需的预收款字段
        required_pr = {
            "*收款账户.名称 # settleaccount.name": col_pr_settle,
            "*收款金额 # amount": col_pr_amount,
            "*单据日期 # billdate": col_pr_date,
            "*单据编号 # billno": col_pr_billno,
            "*结算客户.名称 # itemclass.name": col_pr_client,
            "*订单编号 # srcbillno": col_pr_order
        }
        if mode == "银行":
            required_pr["*汇率 # exchangerate"] = col_pr_exrate
            required_pr["*币别.名称 # currency.name"] = col_pr_currency
        elif mode == "支付宝":
            required_pr["*收款备注 # comment"] = col_pr_comment

        missing_pr_cols = [k for k, v in required_pr.items() if v is None]
        if missing_pr_cols:
            local_log(f"错误：预收款文件中缺少必要列: {', '.join(missing_pr_cols)}")
            self.finalize_run(False, log_messages, target_output_dir)
            return

        # 筛选并加载预收款数据（全程安全取值，兼容合并单元格）
        pr_items = []
        target_settle = cfg["pr_settle_account"]
        skipped_settle_count = 0

        for r in range(PR_DATA_START_ROW, pr_ws.max_row + 1):
            settle_val = self.get_cell_value(pr_ws, r, col_pr_settle, pr_merged_map)
            if settle_val is None:
                continue
            
            settle_str = str(settle_val).strip()
            if settle_str != target_settle:
                skipped_settle_count += 1
                continue

            amount_raw = self.get_cell_value(pr_ws, r, col_pr_amount, pr_merged_map)
            amount = self.clean_amount(amount_raw)
            if amount is None:
                local_log(f"提示 [预收款行号 {r}]: 金额为空或无法解析，已跳过。")
                continue

            date_raw = self.get_cell_value(pr_ws, r, col_pr_date, pr_merged_map)
            date_val = self.parse_date_only(date_raw)
            if date_val is None:
                local_log(f"提示 [预收款行号 {r}]: 日期为空或格式无法解析，已跳过。")
                continue

            billno = str(self.get_cell_value(pr_ws, r, col_pr_billno, pr_merged_map) or "").strip()
            client = str(self.get_cell_value(pr_ws, r, col_pr_client, pr_merged_map) or "").strip()
            order = str(self.get_cell_value(pr_ws, r, col_pr_order, pr_merged_map) or "").strip()
            
            exrate = Decimal('1.0')
            if col_pr_exrate:
                exrate_raw = self.get_cell_value(pr_ws, r, col_pr_exrate, pr_merged_map)
                cleaned_ex = self.clean_amount(exrate_raw)
                if cleaned_ex is not None:
                    exrate = cleaned_ex

            currency = ""
            if col_pr_currency:
                currency = str(self.get_cell_value(pr_ws, r, col_pr_currency, pr_merged_map) or "").strip()

            comment = ""
            if col_pr_comment:
                comment = str(self.get_cell_value(pr_ws, r, col_pr_comment, pr_merged_map) or "").strip()

            pr_items.append({
                "row_num": r,
                "amount": amount,
                "date": date_val,
                "billno": billno,
                "client": client,
                "order": order,
                "exchangerate": exrate,
                "currency": currency,
                "comment": comment
            })

        local_log(f"预收款中账户过滤统计: 过滤非【{target_settle}】的记录共计 {skipped_settle_count} 条。")
        local_log(f"预收款中对应账户为【{target_settle}】的有效记录数: {len(pr_items)}")

        # 2. 读取并预解析流水文件
        try:
            bank_wb_values = openpyxl.load_workbook(bank_path, data_only=True)
            bank_ws_values = self.get_sheet_insensitive(bank_wb_values, "sheet1")
            # 解析流水数据工作表的合并单元格
            bank_merged_map = self.get_merged_cells_map(bank_ws_values)
            
            bank_wb_save = openpyxl.load_workbook(bank_path, data_only=False)
            bank_ws_save = self.get_sheet_insensitive(bank_wb_save, "sheet1")
        except Exception as e:
            local_log(f"错误: 读取流水文件失败: {str(e)}")
            self.finalize_run(False, log_messages, target_output_dir)
            return

        # 获取流水表头（安全取值）
        bank_headers = [
            str(self.get_cell_value(bank_ws_values, cfg["bank_header_row"], col, bank_merged_map)).strip()
            if self.get_cell_value(bank_ws_values, cfg["bank_header_row"], col, bank_merged_map) is not None else ""
            for col in range(1, bank_ws_values.max_column + 1)
        ]
        
        # 寻找流水文件各目标列索引
        bank_cols_idx = {}
        missing_bank_cols = []
        for key, field_name in cfg["bank_cols"].items():
            idx = None
            if key == "date" and mode == "微信":
                idx = self.find_header_col(bank_headers, ["交易时间"])
            else:
                idx = self.find_header_col(bank_headers, [field_name])
            
            if idx is None:
                missing_bank_cols.append(field_name)
            else:
                bank_cols_idx[key] = idx

        if missing_bank_cols:
            local_log(f"错误：流水文件中未找到必要表头列: {', '.join(missing_bank_cols)}")
            self.finalize_run(False, log_messages, target_output_dir)
            return

        # 解析流水表数据（安全取值）
        bank_rows_data = []
        for r in range(cfg["bank_data_start_row"], bank_ws_values.max_row + 1):
            row_vals = [self.get_cell_value(bank_ws_values, r, c, bank_merged_map) for c in range(1, bank_ws_values.max_column + 1)]
            if all(v is None for v in row_vals):
                continue

            amount_raw = self.get_cell_value(bank_ws_values, r, bank_cols_idx["amount"], bank_merged_map)
            amount = self.clean_amount(amount_raw)
            if amount is None:
                local_log(f"提示 [流水行号 {r}]: 金额为空或无法解析。")
                continue

            date_raw = self.get_cell_value(bank_ws_values, r, bank_cols_idx["date"], bank_merged_map)
            date_val = self.parse_date_only(date_raw)

            opp_name = ""
            if "opp_name" in bank_cols_idx:
                opp_name = str(self.get_cell_value(bank_ws_values, r, bank_cols_idx["opp_name"], bank_merged_map) or "").strip()

            currency = ""
            if "currency" in bank_cols_idx:
                currency = str(self.get_cell_value(bank_ws_values, r, bank_cols_idx["currency"], bank_merged_map) or "").strip()

            bank_rows_data.append({
                "row_num": r,
                "amount": amount,
                "date": date_val,
                "opp_name": opp_name,
                "currency": currency
            })

        local_log(f"流水文件有效记录行数: {len(bank_rows_data)}")

        # 3. 对账匹配比对 (建立多对多候选映射)
        bank_to_pr = {}
        pr_to_bank = {}

        for br in bank_rows_data:
            br_num = br["row_num"]
            for pr in pr_items:
                pr_num = pr["row_num"]
                
                matched = False
                if mode == "银行":
                    is_chinese = self.is_chinese_name(br["opp_name"])
                    if is_chinese:
                        # 2.1 中文对方账户名称匹配
                        cond1 = (br["amount"] == pr["amount"])
                        cond2 = (br["date"] == pr["date"] and br["date"] is not None)
                        cond3 = (br["opp_name"] == pr["client"])
                        cond4 = (br["currency"] == "人民币" and pr["currency"] == "人民币")
                        if cond1 and cond2 and cond3 and cond4:
                            matched = True
                    else:
                        # 2.2 英文对方账户名称匹配
                        expected_usd_amount = (pr["amount"] * pr["exchangerate"]).quantize(Decimal('0.01'))
                        cond1 = (br["amount"] == expected_usd_amount)
                        cond2 = (pr["currency"] == "美元")
                        if cond1 and cond2:
                            matched = True
                            
                elif mode == "支付宝":
                    cond1 = (br["amount"] == pr["amount"])
                    cond2 = (br["date"] == pr["date"] and br["date"] is not None)
                    if cond1 and cond2:
                        matched = True
                        
                elif mode == "微信":
                    cond1 = (br["amount"] == pr["amount"])
                    cond2 = (br["date"] == pr["date"] and br["date"] is not None)
                    if cond1 and cond2:
                        matched = True

                if matched:
                    bank_to_pr.setdefault(br_num, []).append(pr)
                    pr_to_bank.setdefault(pr_num, []).append(br_num)

        # 4. 详细校验、填充以及未通过原因分析
        matched_count = 0
        unmatched_count = 0
        duplicate_count = 0

        # 在要保存的工作表头写入新增列的名称
        for col_idx, (header_name, _) in cfg["output_cols"].items():
            bank_ws_save.cell(row=cfg["bank_header_row"], column=col_idx, value=header_name)

        local_log("\n" + "-"*15 + " 开始执行对账比对校验 " + "-"*15)

        for br in bank_rows_data:
            br_num = br["row_num"]
            br_amount = br["amount"]
            candidates = bank_to_pr.get(br_num, [])

            if len(candidates) == 1:
                pr_item = candidates[0]
                associated_bank_rows = pr_to_bank.get(pr_item["row_num"], [])
                
                # 双向确认唯一匹配
                if len(associated_bank_rows) == 1:
                    for col_idx, (_, pr_field) in cfg["output_cols"].items():
                        val_to_write = pr_item.get(pr_field, "")
                        bank_ws_save.cell(row=br_num, column=col_idx, value=val_to_write)
                    matched_count += 1
                else:
                    unmatched_count += 1
                    duplicate_count += 1
                    local_log(
                        f"无法匹配 [重复对应 - 流水行号 {br_num}]: 虽然匹配到预收款行 {pr_item['row_num']} (金额: {br_amount})，"
                        f"但该预收款记录同时对应了其他多个流水行号 {associated_bank_rows}，已跳过以防错配。"
                    )
            elif len(candidates) > 1:
                unmatched_count += 1
                duplicate_count += 1
                candidate_rows = [c["row_num"] for c in candidates]
                local_log(
                    f"无法匹配 [特征重复 - 流水行号 {br_num}]: 金额 {br_amount} 在预收款中同时匹配到多条记录 (行号: {candidate_rows})，无法确保唯一，已跳过。"
                )
            else:
                unmatched_count += 1
                # 匹配失败，深度追溯原因并输出
                potential_prs = []
                for pr in pr_items:
                    is_amount_match = False
                    if mode == "银行":
                        is_chinese = self.is_chinese_name(br["opp_name"])
                        if is_chinese:
                            is_amount_match = (br["amount"] == pr["amount"])
                        else:
                            expected_usd_amount = (pr["amount"] * pr["exchangerate"]).quantize(Decimal('0.01'))
                            is_amount_match = (br["amount"] == expected_usd_amount)
                    else:
                        is_amount_match = (br["amount"] == pr["amount"])
                    
                    if is_amount_match:
                        potential_prs.append(pr)

                if potential_prs:
                    # 有相同金额的记录，但是由于其他属性限制导致过滤未通过
                    reasons_list = []
                    for p in potential_prs:
                        reasons = []
                        if mode == "银行":
                            is_chinese = self.is_chinese_name(br["opp_name"])
                            if is_chinese:
                                if br["date"] != p["date"]:
                                    reasons.append(f"日期不匹配(流水:{br['date']} vs 预收:{p['date']})")
                                if br["opp_name"] != p["client"]:
                                    reasons.append(f"名称不匹配(流水对方户名:{br['opp_name']} vs 预收客户名称:{p['client']})")
                                if br["currency"] != "人民币" or p["currency"] != "人民币":
                                    reasons.append(f"币别不是人民币(流水:{br['currency']} vs 预收:{p['currency']})")
                            else:
                                if p["currency"] != "美元":
                                    reasons.append(f"预收款币别不是美元(预收币别:{p['currency']})")
                        else:
                            if br["date"] != p["date"]:
                                reasons.append(f"日期不匹配(流水:{br['date']} vs 预收:{p['date']})")
                        
                        reasons_str = " 且 ".join(reasons) if reasons else "多重规则不满足"
                        reasons_list.append(f"预收款行 {p['row_num']}(原因: {reasons_str})")
                    
                    local_log(
                        f"未通过匹配 [过滤失败 - 流水行号 {br_num}]: 金额 {br_amount} 的预收款记录虽存在，"
                        f"但因其他规则被过滤。具体详情 -> {'; '.join(reasons_list)}"
                    )
                else:
                    local_log(
                        f"未通过匹配 [无匹配数据 - 流水行号 {br_num}]: 未在当前预收款库中找到匹配金额为 {br_amount} 且符合对应账户条件的记录。"
                    )

        # 5. 输出保存
        end_time = datetime.datetime.now()
        timestamp_str = end_time.strftime("%Y%m%d_%H%M%S")
        
        output_excel_path = os.path.join(target_output_dir, f"{timestamp_str}.xlsx")
        output_log_path = os.path.join(target_output_dir, f"{timestamp_str}.txt")
        
        try:
            bank_wb_save.save(output_excel_path)
            local_log(f"\n结果Excel文件已成功保存至: {output_excel_path}")
        except Exception as e:
            local_log(f"错误: 保存结果文件失败: {str(e)}")
            self.finalize_run(False, log_messages, target_output_dir)
            return

        local_log("\n" + "="*15 + " 核对统计汇总 " + "="*15)
        local_log(f"完成处理时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        local_log(f"数据比对耗时: {(end_time - start_time).total_seconds():.2f} 秒")
        local_log(f"有效对账记录总数: {len(bank_rows_data)}")
        local_log(f"成功唯一匹配关联: {matched_count}")
        local_log(f"未匹配/冲突记录: {unmatched_count}")
        local_log("="*44)

        try:
            with open(output_log_path, "w", encoding="utf-8") as f:
                for line in log_messages:
                    f.write(line + "\n")
            local_log(f"日志文件已成功保存至: {output_log_path}")
        except Exception as e:
            local_log(f"警告: 无法保存日志文件: {str(e)}")

        self.finalize_run(True, log_messages, target_output_dir)

    def finalize_run(self, success, log_messages, target_output_dir):
        self.run_btn.configure(state=tk.NORMAL)
        if success:
            self.status_text.set("核对完毕")
            messagebox.showinfo("成功", f"账目数据比对执行完成！\n\n核对后的Excel文件和TXT日志已保存至以下文件夹目录：\n\n{target_output_dir}")
        else:
            self.status_text.set("核对失败")
            messagebox.showerror("出错", "在对账过程中遇到了错误，请检查日志。")


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = ReconciliationApp()
    app.mainloop()