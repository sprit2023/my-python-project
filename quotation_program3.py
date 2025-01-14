import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from xpinyin import Pinyin
from decimal import Decimal

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("报价单系统 v3.0")
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
        
        # 总价显示
        tk.Label(bottom_frame, text="含税总价：").pack(side=tk.LEFT)
        self.total_label = tk.Label(bottom_frame, text="0.00", font=("Arial", 14, "bold"))
        self.total_label.pack(side=tk.LEFT, padx=5)
        
        tk.Label(bottom_frame, text="大写金额：").pack(side=tk.LEFT, padx=(20, 0))
        self.total_cn_label = tk.Label(bottom_frame, text="零元整", font=("Arial", 14))
        self.total_cn_label.pack(side=tk.LEFT)
        
    def create_product_table(self):
        columns = ("物料编码", "物料名称", "规格型号", "含税单价")
        self.product_tree = ttk.Treeview(self.product_frame, columns=columns, show="headings")
        
        # 设置列宽
        col_widths = [150, 200, 300, 100]
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
        
    def create_quotation_table(self):
        columns = ("物料编码", "物料名称", "规格型号", "数量", "含税单价", "小计", "操作")
        self.quotation_tree = ttk.Treeview(self.quotation_frame, columns=columns, show="headings")
        
        # 设置列宽
        col_widths = [120, 150, 250, 80, 100, 100, 80]
        for col, width in zip(columns, col_widths):
            self.quotation_tree.heading(col, text=col)
            self.quotation_tree.column(col, width=width, anchor="center")
            
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.quotation_frame, orient=tk.VERTICAL, command=self.quotation_tree.yview)
        self.quotation_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.quotation_tree.pack(fill=tk.BOTH, expand=True)
        
        # 绑定编辑事件
        self.quotation_tree.bind("<Double-1>", self.edit_quantity)
        
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
            existing['quantity'] += 1
        else:
            self.quotation_items.append({
                'code': values[0],
                'name': values[1],
                'spec': values[2],
                'price': float(values[3]),
                'quantity': 1
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
                "编辑"
            ))
            
    def calculate_total(self):
        total = sum(item['price'] * item['quantity'] 
                   for item in self.quotation_items)
        self.total_label.config(text=f"{total:.2f}")
        self.total_cn_label.config(text=self.to_chinese_amount(total))
        
    def to_chinese_amount(self, amount):
        # 金额中文大写转换
        units = ['', '万', '亿']
        nums = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
        decimal_units = ['角', '分']
        integer_units = ['元', '拾', '佰', '仟']
        
        integer_part = int(amount)
        decimal_part = round(amount - integer_part, 2)
        
        # 处理整数部分
        result = ''
        unit_index = 0
        while integer_part > 0:
            section = integer_part % 10000
            if section > 0:
                section_str = ''
                for i, digit in enumerate(str(section)[::-1]):
                    if digit != '0':
                        section_str = nums[int(digit)] + integer_units[i] + section_str
                result = section_str + units[unit_index] + result
            unit_index += 1
            integer_part = integer_part // 10000
            
        # 处理小数部分
        decimal_str = ''
        if decimal_part > 0:
            decimal_str = ''.join([
                nums[int(d)] + u 
                for d, u in zip(f"{decimal_part:.2f}".split('.')[1], decimal_units)
                if d != '0'
            ])
            
        if not result:
            result = nums[0]
        if not decimal_str:
            decimal_str = '整'
            
        return result + decimal_str
        
    def edit_quantity(self, event):
        # 编辑数量
        region = self.quotation_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.quotation_tree.identify_column(event.x)
        if column != "#4":  # 只允许编辑数量列
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
        
        def save_edit(event):
            try:
                new_value = int(entry.get())
                if new_value < 1:
                    raise ValueError("数量不能小于1")
                    
                # 更新产品数量
                code = self.quotation_tree.item(item)['values'][0]
                product = next((p for p in self.quotation_items 
                              if p['code'] == code), None)
                if product:
                    product['quantity'] = new_value
                    
                # 更新表格
                self.update_quotation_table()
                self.calculate_total()
                
            except ValueError as e:
                messagebox.showerror("错误", f"无效的数量值: {str(e)}")
            finally:
                entry.destroy()
                
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationApp(root)
    root.mainloop()
