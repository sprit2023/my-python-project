import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from xpinyin import Pinyin
from datetime import datetime

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("报价单系统-Sprit.Zeng V3.0-测试版")  # 修改窗口标题
        self.root.geometry("1200x800")
        
        # 初始化
        self.excel_data = None
        self.quotation_items = []
        self.pinyin = Pinyin()
        self.filter_timer = None
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 主布局
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 工具栏
        toolbar = tk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=5)
        
        # 保存按钮
        btn_save = tk.Button(toolbar, text="保存报价单", command=self.save_quotation)
        btn_save.pack(side=tk.RIGHT, padx=5)
        
        # 导出按钮
        btn_export = tk.Button(toolbar, text="导出报价单", command=self.export_excel)
        btn_export.pack(side=tk.RIGHT, padx=5)
        
        # 导入按钮
        btn_import = tk.Button(toolbar, text="导入Excel", command=self.import_excel)
        btn_import.pack(side=tk.LEFT, padx=5)
        
        # 搜索框
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=10)
        self.search_entry.bind("<KeyRelease>", self.filter_products)
        
        # 产品表格
        self.product_frame = tk.LabelFrame(main_frame, text="产品列表")
        self.product_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.create_product_table()
        
        # 报价单表格
        self.quotation_frame = tk.LabelFrame(main_frame, text="报价单")
        self.quotation_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.create_quotation_table()
        
        # 底部信息
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        # 第一行：含税总价和毛利率
        row1_frame = tk.Frame(bottom_frame)
        row1_frame.pack(fill=tk.X, pady=5)
        
        # 含税总价
        tk.Label(row1_frame, text="含税总价：").pack(side=tk.LEFT, padx=(46, 0))  # 右移两个汉字的位置
        self.total_label = ttk.Entry(row1_frame, font=("Arial", 14, "bold"), state="readonly", width=15)
        self.total_label.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row1_frame, text="大写金额：").pack(side=tk.LEFT, padx=(20, 0))
        self.total_cn_label = ttk.Entry(row1_frame, font=("Arial", 14), state="readonly", width=40)  # 加大大写金额框
        self.total_cn_label.pack(side=tk.LEFT, padx=5)
        
        # 毛利率输入框
        tk.Label(row1_frame, text="毛利率（%）：").pack(side=tk.LEFT, padx=(40, 0))  # 向右移动一些
        self.profit_margin_entry = ttk.Entry(row1_frame, width=10)
        self.profit_margin_entry.pack(side=tk.LEFT, padx=5)
        self.profit_margin_entry.bind("<KeyRelease>", self.calculate_total)  # 绑定输入框内容变化事件
        
        # 第二行：最终含税总价
        row2_frame = tk.Frame(bottom_frame)
        row2_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(row2_frame, text="最终含税总价：").pack(side=tk.LEFT, padx=(20, 0))  # 向左移动2个汉字的位置
        self.final_total_label = ttk.Entry(row2_frame, font=("Arial", 14, "bold"), state="readonly", width=15)
        self.final_total_label.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row2_frame, text="大写金额：").pack(side=tk.LEFT, padx=(20, 0))
        self.final_total_cn_label = ttk.Entry(row2_frame, font=("Arial", 14), state="readonly", width=40)  # 加大大写金额框
        self.final_total_cn_label.pack(side=tk.LEFT, padx=5)
        
        # 主内容区域
        content_frame = tk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧区域
        left_frame = tk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 选中商品信息
        self.selected_item_frame = tk.LabelFrame(left_frame, text="选中商品信息")
        self.selected_item_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.selected_item_info = tk.Text(self.selected_item_frame, font=("Arial", 12))
        self.selected_item_info.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 右侧历史记录区域
        history_frame = tk.LabelFrame(content_frame, text="历史报价记录")
        history_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10, ipadx=10, ipady=10)
        
        # 历史记录表格
        self.history_tree = ttk.Treeview(history_frame, columns=("时间", "总金额"), show="headings")
        self.history_tree.heading("时间", text="时间")
        self.history_tree.heading("总金额", text="总金额")
        self.history_tree.column("时间", width=150)
        self.history_tree.column("总金额", width=100)
        self.history_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def create_product_table(self):
        columns = ("物料编码", "物料名称", "规格型号", "数量", "含税单价")
        self.product_tree = ttk.Treeview(self.product_frame, columns=columns, show="headings")
        
        # 设置列宽
        col_widths = [150, 200, 250, 80, 100]
        for col, width in zip(columns, col_widths):
            self.product_tree.heading(col, text=col)
            self.product_tree.column(col, width=width, anchor="center")
            
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.product_frame, orient=tk.VERTICAL, command=self.product_tree.yview)
        self.product_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.product_tree.pack(fill=tk.BOTH, expand=True)
        
        # 绑定双击事件
        self.product_tree.bind("<Double-1>", self.add_to_quotation)
        
        # 绑定 Ctrl+C 复制事件
        self.product_tree.bind("<Control-c>", self.copy_selected_text)
        
    def create_quotation_table(self):
        columns = ("物料编码", "物料名称", "规格型号", "数量", "含税单价", "小计", "操作")
        self.quotation_tree = ttk.Treeview(self.quotation_frame, columns=columns, show="headings")
        
        # 设置列宽
        col_widths = [120, 150, 250, 80, 100, 100, 50]
        for col, width in zip(columns, col_widths):
            self.quotation_tree.heading(col, text=col)
            self.quotation_tree.column(col, width=width, anchor="center")
            
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.quotation_frame, orient=tk.VERTICAL, command=self.quotation_tree.yview)
        self.quotation_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.quotation_tree.pack(fill=tk.BOTH, expand=True)
        
        # 绑定删除事件
        self.quotation_tree.bind("<Button-1>", self.delete_item)
        
        # 绑定 Ctrl+C 复制事件
        self.quotation_tree.bind("<Control-c>", self.copy_selected_text)
        
        # 绑定 DEL 键删除事件
        self.quotation_tree.bind("<Delete>", self.delete_item)
        
        # 绑定选中事件
        self.quotation_tree.bind("<<TreeviewSelect>>", self.show_selected_item_info)
        
    def copy_selected_text(self, event=None):
        # 获取当前焦点所在的 Treeview
        widget = self.root.focus_get()
        if isinstance(widget, ttk.Treeview):
            selected_item = widget.selection()
            if selected_item:
                item_values = widget.item(selected_item)['values']
                selected_text = "\t".join(map(str, item_values))
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.root.update()
                messagebox.showinfo("复制成功", "已复制到剪贴板！")
        
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.excel_data = pd.read_excel(file_path)
                self.load_products()
                messagebox.showinfo("成功", "Excel文件导入成功！")
            except Exception as e:
                messagebox.showerror("错误", f"导入Excel文件失败：{str(e)}")
                
    def load_products(self):
        # 清空表格
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
            
        # 加载产品
        for _, row in self.excel_data.iterrows():
            self.product_tree.insert("", "end", values=(
                row.get('物料编码', ''),
                row.get('物料名称', ''),
                row.get('规格型号', ''),
                row.get('数量', 1),  # 读取 Excel 中的数量，默认为 1
                f"{row.get('含税单价', 0):.2f}"
            ))
            
    def filter_products(self, event=None):
        keyword = self.search_var.get().strip().lower()
        if not keyword:
            self.load_products()
            return
            
        # 清空表格
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
            
        # 过滤产品
        for _, row in self.excel_data.iterrows():
            name = str(row.get('物料名称', '')).lower()
            spec = str(row.get('规格型号', '')).lower()
            pinyin_name = self.pinyin.get_pinyin(name, '').lower()
            
            if (keyword in name or 
                keyword in spec or
                keyword in pinyin_name):
                
                self.product_tree.insert("", "end", values=(
                    row.get('物料编码', ''),
                    row.get('物料名称', ''),
                    row.get('规格型号', ''),
                    row.get('数量', 1),  # 读取 Excel 中的数量，默认为 1
                    f"{row.get('含税单价', 0):.2f}"
                ))
                
    def add_to_quotation(self, event):
        selected = self.product_tree.selection()
        if not selected:
            return
            
        item = self.product_tree.item(selected)
        values = item['values']
        
        # 检查是否已存在
        existing = next((i for i in self.quotation_items 
                       if i['code'] == values[0]), None)
        
        if existing:
            existing['quantity'] += 1  # 如果已存在，数量加 1
        else:
            self.quotation_items.append({
                'code': values[0],
                'name': values[1],
                'spec': values[2],
                'price': float(values[4]),  # 含税单价
                'quantity': 1  # 默认数量为 1
            })
            
        self.update_quotation_table()
        self.calculate_total()
        
    def update_quotation_table(self):
        # 清空表格
        for item in self.quotation_tree.get_children():
            self.quotation_tree.delete(item)
            
        # 更新报价单
        for item in self.quotation_items:
            subtotal = item['price'] * item['quantity']
            self.quotation_tree.insert("", "end", values=(
                item['code'],
                item['name'],
                item['spec'],
                item['quantity'],
                f"{item['price']:.2f}",
                f"{subtotal:.2f}",
                "×"  # 使用“×”表示删除
            ))
            
        # 绑定双击事件
        self.quotation_tree.bind("<Double-1>", self.edit_item)
        
    def calculate_total(self, event=None):
        total = sum(item['price'] * item['quantity'] 
                   for item in self.quotation_items)
        # 使用千分位格式化
        self.total_label.config(state="normal")
        self.total_label.delete(0, tk.END)
        self.total_label.insert(0, f"{total:,.2f}")
        self.total_label.config(state="readonly")
        
        self.total_cn_label.config(state="normal")
        self.total_cn_label.delete(0, tk.END)
        self.total_cn_label.insert(0, self.to_chinese_amount(total))
        self.total_cn_label.config(state="readonly")
        
        # 计算最终含税总价
        try:
            profit_margin = float(self.profit_margin_entry.get()) / 100  # 将百分比转换为小数
            final_total = total * (1 + profit_margin)  # 含税总价乘以毛利率
            self.final_total_label.config(state="normal")
            self.final_total_label.delete(0, tk.END)
            self.final_total_label.insert(0, f"{final_total:,.2f}")
            self.final_total_label.config(state="readonly")
            
            self.final_total_cn_label.config(state="normal")
            self.final_total_cn_label.delete(0, tk.END)
            self.final_total_cn_label.insert(0, self.to_chinese_amount(final_total))
            self.final_total_cn_label.config(state="readonly")
        except ValueError:
            # 如果输入的不是有效数字，清空最终含税总价
            self.final_total_label.config(state="normal")
            self.final_total_label.delete(0, tk.END)
            self.final_total_label.config(state="readonly")
            
            self.final_total_cn_label.config(state="normal")
            self.final_total_cn_label.delete(0, tk.END)
            self.final_total_cn_label.config(state="readonly")
        
    def to_chinese_amount(self, amount):
        # 金额中文大写转换
        units = ["", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿"]
        nums = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
        decimal_units = ["角", "分"]

        # 处理整数部分
        integer_part = int(amount)
        integer_str = str(integer_part)
        result = ""
        zero_flag = False  # 标记是否需要添加“零”

        # 从高位到低位处理
        length = len(integer_str)
        for i, char in enumerate(integer_str):
            if char == '0':
                zero_flag = True
            else:
                if zero_flag:
                    result += nums[0]  # 添加“零”
                    zero_flag = False
                result += nums[int(char)] + units[length - i - 1]

        # 处理小数部分
        decimal_part = round(amount - integer_part, 2)
        decimal_str = ""
        if decimal_part > 0:
            decimal_str = "".join([
                nums[int(d)] + u 
                for d, u in zip(f"{decimal_part:.2f}".split('.')[1], decimal_units)
                if d != '0'
            ])

        # 拼接结果
        if not result:
            result = "零"
        if not decimal_str:
            decimal_str = "整"

        return result + "元" + decimal_str
        
    def delete_item(self, event):
        # 获取点击的行和列
        region = self.quotation_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.quotation_tree.identify_column(event.x)
        if column != "#7":  # 只允许点击操作列
            return
            
        item = self.quotation_tree.identify_row(event.y)
        code = self.quotation_tree.item(item)['values'][0]
        
        # 从报价单中删除
        self.quotation_items = [i for i in self.quotation_items if i['code'] != code]
        
        # 更新表格
        self.update_quotation_table()
        self.calculate_total()
        
    def edit_item(self, event):
        # 编辑数量或含税单价
        region = self.quotation_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.quotation_tree.identify_column(event.x)
        if column not in ("#4", "#5"):  # 只允许编辑数量列或含税单价列
            return
            
        item = self.quotation_tree.identify_row(event.y)
        column_box = self.quotation_tree.bbox(item, column)
        
        # 获取当前值
        current_value = self.quotation_tree.set(item, column)
        
        # 创建编辑框
        entry = ttk.Entry(self.quotation_tree, width=column_box[2])
        entry.place(x=column_box[0], y=column_box[1], 
                   width=column_box[2], height=column_box[3])
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()
        
        def save_edit(event=None):
            try:
                # 获取新值
                new_value = float(entry.get())  # 支持小数
                
                # 获取选中商品的编码
                code = self.quotation_tree.item(item)['values'][0]
                product = next((p for p in self.quotation_items 
                              if p['code'] == code), None)
                
                if product:
                    if column == "#4":  # 编辑数量
                        product['quantity'] = int(new_value)  # 数量为整数
                    elif column == "#5":  # 编辑含税单价
                        product['price'] = new_value  # 含税单价为浮点数
                    
                # 更新表格
                self.update_quotation_table()
                self.calculate_total()
                
            except ValueError as e:
                # 如果输入的不是有效数字，提示错误
                messagebox.showerror("错误", f"无效的输入值: {str(e)}")
            finally:
                entry.destroy()
            
        # 绑定回车事件
        entry.bind("<Return>", save_edit)
        
        # 绑定离开事件（点击其他地方）
        entry.bind("<FocusOut>", save_edit)
    
    def save_quotation(self):
        if not self.quotation_items:
            messagebox.showwarning("警告", "当前没有报价单数据可以保存！")
            return
            
        # 获取当前时间
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 计算总金额
        total = sum(item['price'] * item['quantity'] for item in self.quotation_items)
        
        # 保存完整报价单信息
        quotation_data = {
            'time': current_time,
            'total': total,
            'items': self.quotation_items.copy()  # 保存当前报价单的副本
        }
        
        # 添加到历史记录
        item_id = self.history_tree.insert("", "end", values=(current_time, f"{total:,.2f}"))
        self.history_tree.set(item_id, "data", quotation_data)  # 将完整数据存储在Treeview中
        
        messagebox.showinfo("成功", "报价单已保存到历史记录！")
        
    def export_excel(self):
        if not self.quotation_items:
            messagebox.showwarning("警告", "当前没有报价单数据可以导出！")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            try:
                # 创建DataFrame
                data = []
                for item in self.quotation_items:
                    data.append([
                        item['code'],
                        item['name'],
                        item['spec'],
                        item['quantity'],
                        item['price'],
                        item['price'] * item['quantity']
                    ])
                
                df = pd.DataFrame(data, columns=[
                    "物料编码", "物料名称", "规格型号", "数量", "含税单价", "小计"
                ])
                
                # 导出Excel
                df.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"报价单已成功导出到：\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出Excel失败：{str(e)}")

    def show_selected_item_info(self, event):
        # 获取选中的商品
        selected_item = self.quotation_tree.selection()
        if selected_item:
            item_values = self.quotation_tree.item(selected_item)['values']
            info_text = (
                f"物料编码: {item_values[0]}\n"
                f"物料名称: {item_values[1]}\n"
                f"规格型号: {item_values[2]}\n"
                f"数量: {item_values[3]}\n"
                f"含税单价: {item_values[4]}\n"
                f"小计: {item_values[5]}"
            )
            self.selected_item_info.config(state="normal")
            self.selected_item_info.delete(1.0, tk.END)
            self.selected_item_info.insert(tk.END, info_text)
            self.selected_item_info.config(state="disabled")

    def load_history_quotation(self, event):
        # 获取选中的历史记录
        selected_item = self.history_tree.selection()
        if not selected_item:
            return
            
        # 获取保存的报价单数据
        quotation_data = self.history_tree.set(selected_item, "data")
        
        if not quotation_data:
            messagebox.showwarning("警告", "无法加载历史报价单数据！")
            return
            
        try:
            # 清空当前报价单
            self.quotation_items.clear()
            
            # 加载历史报价单
            self.quotation_items = quotation_data['items']
            
            # 更新报价单表格
            self.update_quotation_table()
            
            # 更新总金额
            self.calculate_total()
            
            messagebox.showinfo("成功", "历史报价单已加载！")
        except Exception as e:
            messagebox.showerror("错误", f"加载历史报价单失败：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationApp(root)
    root.mainloop()
