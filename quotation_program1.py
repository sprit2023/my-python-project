try:
    import pandas as pd
except ImportError:
    print("请先安装pandas库：pip install pandas")
    exit(1)

from datetime import datetime
import tkinter as tk

from tkinter import filedialog, messagebox, ttk

class QuotationGenerator:
    def __init__(self, excel_path, start_row=None, end_row=None):
        self.df = pd.read_excel(excel_path, header=None)
        self.start_row = start_row if start_row is not None else 11
        self.end_row = end_row if end_row is not None else len(self.df)-5
        self.products = self._extract_products()
        
    def _extract_products(self):
        # 提取产品信息
        products = []
        for i in range(self.start_row, self.end_row):
            config = self.df.iloc[i, 0]
            if pd.notna(config) and '：' not in str(config) and str(config) != '数量':
                try:
                    quantity = int(self.df.iloc[i, 1]) if pd.notna(self.df.iloc[i, 1]) else 0
                except ValueError:
                    quantity = 0
                
                # 规格型号
                spec = self.df.iloc[i, 2] if pd.notna(self.df.iloc[i, 2]) else ''
                
                # 直接从Excel读取含税单价
                try:
                    unit_price = float(self.df.iloc[i, 6]) if pd.notna(self.df.iloc[i, 6]) else 0
                except ValueError:
                    unit_price = 0
                
                products.append({
                    '物料编码': config,
                    '物料名称': spec,
                    '规格型号': self.df.iloc[i, 2] if pd.notna(self.df.iloc[i, 2]) else '',
                    '数量': quantity,
                    '单价': unit_price,
                    '小计': round(unit_price * quantity, 2),
                    '含税单价': unit_price
                })
        return products
    
    def generate_quotation(self):
        # 生成报价单
        print("="*50)
        print("报价单")
        print("="*50)
        print(f"日期: {datetime.now().strftime('%Y-%m-%d')}\n")
        
        print("{:<10} {:<8} {:<30} {:<10}".format(
            "产品", "数量", "描述", "单价"
        ))
        print("-"*58)
        
        total = 0
        for product in self.products:
            price = product['单价']
            total += price * product['数量']
            print("{:<10} {:<8} {:<30} {:<10.2f}".format(
                product['配置'], 
                product['数量'], 
                product['描述'], 
                price
            ))
            
        print("\n含税总价: {:.2f} 元".format(total * 1.13))
        print("="*50)

from xpinyin import Pinyin

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("服务器报价单生成器")
        self.root.geometry("1000x700")  # 增大窗口尺寸
        self.products = []  # 初始化产品列表
        
        # 预计算拼音首字母
        self.pinyin_cache = {}
        self.p = Pinyin()
        
        # 延迟处理相关
        self.after_id = None
        self.last_keyword = ""
        
        # 创建控件
        self.create_widgets()
        
    def create_widgets(self):
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="选择Excel文件")
        file_frame.pack(fill="x", padx=15, pady=10)  # 增加间距
        
        self.file_entry = ttk.Entry(file_frame, width=60)  # 增加输入框宽度
        self.file_entry.pack(side="left", padx=10, pady=10)
        
        browse_btn = ttk.Button(file_frame, text="浏览...", command=self.select_file)
        browse_btn.pack(side="left", padx=10, pady=10)
        
        # 规格型号输入区域
        product_frame = ttk.LabelFrame(self.root, text="规格型号")
        product_frame.pack(fill="x", padx=15, pady=10)
        
        self.product_entry = ttk.Entry(product_frame, width=60)
        self.product_entry.pack(side="left", padx=10, pady=10)
        
        add_btn = ttk.Button(product_frame, text="添加", command=self.add_product)
        add_btn.pack(side="left", padx=10, pady=10)
        
        # 自动补全功能
        self.product_list = []
        self.product_var = tk.StringVar()
        self.product_entry.bind('<KeyRelease>', self._on_key_release)
        
        # 选择列表
        self.listbox = tk.Listbox(self.root)
        self.listbox.bind('<<ListboxSelect>>', self._on_select)
        
        # 报价结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="报价结果")
        result_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        columns = ("物料编码", "物料名称", "规格型号", "数量", "含税单价", "小计")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        
        # 设置列宽
        col_widths = [150, 150, 250, 80, 120, 120]  # 调整列宽
        for col, width in zip(columns, col_widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width)
            
        # 使数量列可编辑
        self.tree.bind("<Double-1>", self._edit_cell)
        self.tree.bind("<Return>", self._edit_cell)
        
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 总价显示
        total_frame = ttk.Frame(self.root)
        total_frame.pack(fill="x", padx=15, pady=10)
        
        ttk.Label(total_frame, text="含税总价：", font=("Arial", 12)).pack(side="left")
        self.total_label = ttk.Label(total_frame, text="0.00 元", font=("Arial", 14, "bold"), foreground="red")
        self.total_label.pack(side="left", padx=10)
        
        # 操作按钮
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=15, pady=10)
        
        # 添加操作按钮
        copy_btn = ttk.Button(btn_frame, text="复制", command=self.copy_product)
        copy_btn.pack(side="right", padx=10)
        
        delete_btn = ttk.Button(btn_frame, text="删除", command=self.delete_product)
        delete_btn.pack(side="right", padx=10)
        
        generate_btn = ttk.Button(btn_frame, text="生成报价单", command=self.generate_quotation)
        generate_btn.pack(side="right", padx=10)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择报价文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            # 加载产品列表
            try:
                df = pd.read_excel(file_path, header=None)
                # 从第11行开始读取规格型号列（第3列），去除空值并去重
                self.product_list = df.iloc[11:, 2].dropna().unique().tolist()
                # 清空并重新绑定自动补全列表
                self.listbox.delete(0, tk.END)
                for product in self.product_list:
                    self.listbox.insert(tk.END, product)
            except Exception as e:
                messagebox.showwarning("警告", f"加载产品列表失败：{str(e)}")
                
        
    def _on_key_release(self, event):
        keyword = self.product_entry.get().strip()
        if not keyword:
            self.listbox.pack_forget()
            return
            
        # 取消之前的延迟调用
        if self.after_id:
            self.root.after_cancel(self.after_id)
            
        # 如果输入内容没有变化，直接返回
        if keyword == self.last_keyword:
            return
            
        # 设置200ms延迟处理
        self.after_id = self.root.after(200, self._process_input, keyword)
        self.last_keyword = keyword
        
    def _process_input(self, keyword):
        # 只处理长度大于等于2的输入
        if len(keyword) < 2:
            self.listbox.pack_forget()
            return
            
        matches = []
        
        # 1. 包含关键字匹配
        matches.extend(item for item in self.product_list 
                      if keyword.lower() in item.lower())
        
        # 2. 拼音首字母匹配（使用缓存）
        if len(keyword) <= 4:  # 只对短输入进行拼音匹配
            for item in self.product_list:
                if item not in self.pinyin_cache:
                    self.pinyin_cache[item] = ''.join(
                        [self.p.get_initials(char)[0] for char in item]
                    )
                if keyword.lower() == self.pinyin_cache[item].lower():
                    matches.append(item)
        
        # 3. 模糊匹配（仅对长度>=3的输入）
        if len(keyword) >= 3:
            matches.extend(
                item for item in self.product_list
                if self._fuzzy_match(item.lower(), keyword.lower())
            )
            
        # 去重并限制结果数量
        matches = list(dict.fromkeys(matches))[:15]
        
        # 更新UI
        self.listbox.delete(0, tk.END)
        for match in matches:
            self.listbox.insert(tk.END, match)
            
        if matches:
            self.listbox.pack(fill="x", padx=10, pady=5)
        else:
            self.listbox.pack_forget()
            
    def _fuzzy_match(self, text, pattern):
        """改进的模糊匹配算法，更接近Excel的筛选功能"""
        # 空模式匹配所有文本
        if not pattern:
            return True
            
        # 转换为小写进行不区分大小写的匹配
        text = text.lower()
        pattern = pattern.lower()
        
        # 支持通配符匹配
        if '*' in pattern or '?' in pattern:
            import fnmatch
            return fnmatch.fnmatch(text, pattern)
            
        # 支持部分匹配（任意位置包含）
        if pattern in text:
            return True
            
        # 支持拼音首字母匹配
        pinyin_initials = ''.join([self.p.get_initials(char)[0] for char in text])
        if pattern in pinyin_initials:
            return True
            
        # 支持连续字符匹配（类似Excel的筛选）
        pattern_index = 0
        for char in text:
            if char == pattern[pattern_index]:
                pattern_index += 1
                if pattern_index == len(pattern):
                    return True
                    
        return False
            
    def _on_select(self, event):
        if self.listbox.curselection():
            selected = self.listbox.get(self.listbox.curselection())
            self.product_entry.delete(0, tk.END)
            self.product_entry.insert(0, selected)
            self.listbox.pack_forget()
                
    def add_product(self):
        product_name = self.product_entry.get()
        if product_name and self.file_entry.get():
            try:
                df = pd.read_excel(self.file_entry.get(), header=None)
                product_info = df[df[2] == product_name].iloc[0]
                
                # 添加到产品列表
                self.products.append({
                    '物料编码': product_info[0] if pd.notna(product_info[0]) else f"MAT-{len(self.products)+1:04d}",
                    '物料名称': product_info[1] if pd.notna(product_info[1]) else '',
                    '规格型号': product_name,
                    '数量': 2,  # 初始数量改为2
                    '单价': float(product_info[5]) if pd.notna(product_info[5]) else 0,  # 从第6列读取未税单价
                    '含税单价': round(float(product_info[5]) * 1.13, 2) if pd.notna(product_info[5]) else 0,  # 含税单价 = 未税单价 * 1.13
                    '小计': round(float(product_info[5]) * 1.13 * 2, 2) if pd.notna(product_info[5]) else 0  # 小计 = 含税单价 * 数量
                })
                
                # 更新表格显示
                self.tree.insert("", "end", values=(
                    product_info[0] if pd.notna(product_info[0]) else '',  # 物料名称
                    product_info[1] if pd.notna(product_info[1]) else f"MAT-{len(self.products):04d}",  # 物料编码
                    product_name,  # 规格型号
                    2,  # 数量初始化为2
                    f"{float(product_info[3]):.2f}" if pd.notna(product_info[3]) else '0.00',  # 未税单价
                    f"{float(product_info[3]) * 1.13:.2f}" if pd.notna(product_info[3]) else '0.00',  # 含税单价
                    f"{float(product_info[3]) * 1.13 * 1:.2f}" if pd.notna(product_info[3]) else '0.00',  # 小计
                    # 验证价格计算
                    # 未税单价: Excel原始值
                    # 含税单价: 未税单价 * 1.13
                    # 小计: 含税单价 * 数量
                ))
                
                # 更新总价
                total = sum(p['小计'] for p in self.products)
                self.total_label.config(text=f"{total:.2f} 元")
                
                # 清空输入框
                self.product_entry.delete(0, tk.END)
                self.listbox.pack_forget()
                
            except Exception as e:
                messagebox.showerror("错误", f"添加产品失败：{str(e)}")
            
    def _edit_cell(self, event):
        """处理表格单元格编辑"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.tree.identify_column(event.x)
        if column != "#4":  # 只允许编辑数量列
            return
            
        item = self.tree.identify_row(event.y)
        column_box = self.tree.bbox(item, column)
        
        # 创建编辑框
        entry = ttk.Entry(self.tree, width=column_box[2])
        entry.place(x=column_box[0], y=column_box[1], 
                   width=column_box[2], height=column_box[3])
        
        # 获取当前值
        current_value = self.tree.set(item, column)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()
        
        def save_edit(event):
            try:
                # 获取新值
                new_value = int(entry.get())
                if new_value < 0:
                    raise ValueError("数量不能为负数")
                    
                # 更新表格
                self.tree.set(item, column, new_value)
                
                # 更新产品列表
                item_id = self.tree.index(item)
                self.products[item_id]['数量'] = new_value
                self.products[item_id]['小计'] = self.products[item_id]['含税单价'] * new_value
                
                # 更新表格中的小计
                self.tree.set(item, "#6", f"{self.products[item_id]['小计']:.2f}")
                
                # 重新计算总价
                total = sum(p['小计'] for p in self.products)
                self.total_label.config(text=f"{total:.2f} 元")
                
            except ValueError as e:
                messagebox.showerror("错误", f"无效的数量值: {str(e)}")
            finally:
                entry.destroy()
                
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: entry.destroy())
        
    def copy_product(self):
        """复制选中的产品"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要复制的产品")
            return
            
        for item in selected:
            # 获取选中产品的信息
            values = self.tree.item(item, 'values')
            product = {
                '物料编码': values[0],
                '物料名称': values[1],
                '规格型号': values[2],
                '数量': 1,  # 复制后数量重置为1
                '含税单价': float(values[4]),
                '小计': float(values[4])  # 小计 = 含税单价 * 数量（1）
            }
            
            # 添加到产品列表
            self.products.append(product)
            
            # 添加到表格
            self.tree.insert("", "end", values=(
                product['物料编码'],
                product['物料名称'],
                product['规格型号'],
                product['数量'],
                f"{product['含税单价']:.2f}",
                f"{product['小计']:.2f}"
            ))
            
        # 重新计算总价
        total = sum(p['小计'] for p in self.products)
        self.total_label.config(text=f"{total:.2f} 元")
        
    def delete_product(self):
        """删除选中的产品"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的产品")
            return
            
        # 从产品列表中删除
        for item in selected:
            index = self.tree.index(item)
            del self.products[index]
            
        # 从表格中删除
        self.tree.delete(*selected)
        
        # 重新计算总价
        total = sum(p['小计'] for p in self.products)
        self.total_label.config(text=f"{total:.2f} 元")
        
    def generate_quotation(self):
        file_path = self.file_entry.get()
        if not file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
            
        try:
            # 清空现有数据
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.total_label.config(text="0.00 元")
            
            # 生成报价单
            generator = QuotationGenerator(file_path)
            total = 0
            for product in generator.products:
                # 只添加用户选择的产品
                if product['规格型号'] in [p['规格型号'] for p in self.products]:
                    self.tree.insert("", "end", values=(
                        product['物料名称'],
                        product['物料编码'],
                        product['规格型号'],
                        product['数量'],
                        f"{float(product['单价']):.2f}",  # 单价
                        f"{float(product['含税单价']):.2f}",  # 含税单价
                        f"{float(product['含税单价']) * product['数量']:.2f}"  # 小计
                    ))
                    total += product['含税单价'] * product['数量']
            
            # 显示总价
            self.total_label.config(text=f"{total:.2f} 元")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成报价单失败：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    print("Tkinter root window created")  # 调试信息
    app = QuotationApp(root)
    print("Application initialized")  # 调试信息
    root.mainloop()