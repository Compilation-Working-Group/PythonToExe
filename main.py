import tkinter as tk
from tkinter import ttk, messagebox
import math
import re

class ScientificCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("科学计算器")
        self.root.geometry("500x700")
        self.root.resizable(False, False)
        
        # 设置主题
        self.dark_mode = False
        self.setup_colors()
        
        # 历史记录
        self.history = []
        self.max_history = 10
        
        # 表达式变量
        self.expression = ""
        self.result_var = tk.StringVar()
        self.result_var.set("0")
        
        # 创建界面
        self.setup_ui()
        
        # 绑定键盘事件
        self.root.bind('<Key>', self.key_press)
        
    def setup_colors(self):
        """设置颜色主题"""
        if self.dark_mode:
            # 深色主题
            self.bg_color = "#2e2e2e"
            self.btn_color = "#3c3c3c"
            self.btn_text = "#ffffff"
            self.display_bg = "#1e1e1e"
            self.display_text = "#ffffff"
            self.history_bg = "#252525"
            self.history_text = "#cccccc"
            self.special_btn = "#ff9500"
            self.special_text = "#ffffff"
            self.func_btn = "#505050"
        else:
            # 浅色主题
            self.bg_color = "#f0f0f0"
            self.btn_color = "#ffffff"
            self.btn_text = "#000000"
            self.display_bg = "#ffffff"
            self.display_text = "#000000"
            self.history_bg = "#e8e8e8"
            self.history_text = "#333333"
            self.special_btn = "#ff9500"
            self.special_text = "#ffffff"
            self.func_btn = "#e0e0e0"
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 结果显示区域
        display_frame = tk.Frame(main_frame, bg=self.display_bg, height=80)
        display_frame.pack(fill=tk.X, pady=(0, 10))
        display_frame.pack_propagate(False)
        
        # 历史记录显示
        self.history_label = tk.Label(
            display_frame, 
            text="", 
            anchor=tk.E, 
            bg=self.display_bg, 
            fg=self.history_text,
            font=("Arial", 10)
        )
        self.history_label.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        # 结果显示
        result_label = tk.Label(
            display_frame, 
            textvariable=self.result_var, 
            anchor=tk.E, 
            bg=self.display_bg, 
            fg=self.display_text,
            font=("Arial", 24, "bold")
        )
        result_label.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # 历史记录区域
        history_frame = tk.Frame(main_frame, bg=self.history_bg, height=100)
        history_frame.pack(fill=tk.X, pady=(0, 10))
        history_frame.pack_propagate(False)
        
        history_title = tk.Label(
            history_frame, 
            text="历史记录", 
            bg=self.history_bg, 
            fg=self.history_text,
            font=("Arial", 10, "bold")
        )
        history_title.pack(anchor=tk.W, padx=10, pady=(5, 0))
        
        # 历史记录列表
        self.history_listbox = tk.Listbox(
            history_frame, 
            bg=self.history_bg, 
            fg=self.history_text,
            font=("Arial", 9),
            borderwidth=0,
            highlightthickness=0,
            selectbackground=self.special_btn,
            selectforeground=self.special_text,
            height=5
        )
        self.history_listbox.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        
        # 历史记录滚动条
        history_scrollbar = tk.Scrollbar(self.history_listbox, orient=tk.VERTICAL)
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_listbox.config(yscrollcommand=history_scrollbar.set)
        history_scrollbar.config(command=self.history_listbox.yview)
        
        # 按钮区域
        buttons_frame = tk.Frame(main_frame, bg=self.bg_color)
        buttons_frame.pack(fill=tk.BOTH, expand=True)
        
        # 按钮布局
        buttons = [
            # 第一行
            [('C', self.clear_all, self.special_btn), ('CE', self.clear_entry, self.special_btn), 
             ('⌫', self.backspace, self.special_btn), ('÷', lambda: self.add_to_expression('/'), self.special_btn),
             ('sin', lambda: self.add_function('sin('), self.func_btn), ('cos', lambda: self.add_function('cos('), self.func_btn)],
            
            # 第二行
            [('7', lambda: self.add_to_expression('7'), self.btn_color), ('8', lambda: self.add_to_expression('8'), self.btn_color), 
             ('9', lambda: self.add_to_expression('9'), self.btn_color), ('×', lambda: self.add_to_expression('*'), self.special_btn),
             ('tan', lambda: self.add_function('tan('), self.func_btn), ('log', lambda: self.add_function('log('), self.func_btn)],
            
            # 第三行
            [('4', lambda: self.add_to_expression('4'), self.btn_color), ('5', lambda: self.add_to_expression('5'), self.btn_color), 
             ('6', lambda: self.add_to_expression('6'), self.btn_color), ('-', lambda: self.add_to_expression('-'), self.special_btn),
             ('√', lambda: self.add_function('sqrt('), self.func_btn), ('x²', lambda: self.add_to_expression('**2'), self.func_btn)],
            
            # 第四行
            [('1', lambda: self.add_to_expression('1'), self.btn_color), ('2', lambda: self.add_to_expression('2'), self.btn_color), 
             ('3', lambda: self.add_to_expression('3'), self.btn_color), ('+', lambda: self.add_to_expression('+'), self.special_btn),
             ('π', lambda: self.add_to_expression(str(math.pi)), self.func_btn), ('e', lambda: self.add_to_expression(str(math.e)), self.func_btn)],
            
            # 第五行
            [('0', lambda: self.add_to_expression('0'), self.btn_color), ('.', lambda: self.add_to_expression('.'), self.btn_color), 
             ('(', lambda: self.add_to_expression('('), self.btn_color), (')', lambda: self.add_to_expression(')'), self.btn_color),
             ('=', self.calculate, self.special_btn), ('x^y', lambda: self.add_to_expression('**'), self.func_btn)]
        ]
        
        # 创建按钮
        for i, row in enumerate(buttons):
            for j, (text, command, color) in enumerate(row):
                btn = tk.Button(
                    buttons_frame, 
                    text=text, 
                    command=command,
                    bg=color,
                    fg=self.btn_text if color == self.btn_color else self.special_text,
                    font=("Arial", 14, "bold"),
                    relief=tk.FLAT,
                    height=2,
                    width=5 if text not in ['sin', 'cos', 'tan', 'log', 'x^y'] else 6
                )
                btn.grid(row=i, column=j, padx=2, pady=2, sticky="nsew")
                
                # 鼠标悬停效果
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#d0d0d0" if b.cget("bg") == self.btn_color else "#ffaa33"))
                btn.bind("<Leave>", lambda e, b=btn, c=color: b.config(bg=c))
        
        # 设置按钮区域网格权重
        for i in range(6):
            buttons_frame.grid_columnconfigure(i, weight=1)
        for i in range(5):
            buttons_frame.grid_rowconfigure(i, weight=1)
        
        # 主题切换按钮
        theme_btn = tk.Button(
            main_frame, 
            text="🌙 深色模式" if not self.dark_mode else "☀️ 浅色模式", 
            command=self.toggle_theme,
            bg=self.special_btn,
            fg=self.special_text,
            font=("Arial", 10),
            relief=tk.FLAT
        )
        theme_btn.pack(fill=tk.X, pady=(10, 0))
        
        # 历史记录操作按钮
        history_btn_frame = tk.Frame(main_frame, bg=self.bg_color)
        history_btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        clear_history_btn = tk.Button(
            history_btn_frame, 
            text="清空历史", 
            command=self.clear_history,
            bg=self.func_btn,
            fg=self.btn_text,
            font=("Arial", 9),
            relief=tk.FLAT,
            width=10
        )
        clear_history_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        use_history_btn = tk.Button(
            history_btn_frame, 
            text="使用选中历史", 
            command=self.use_history,
            bg=self.func_btn,
            fg=self.btn_text,
            font=("Arial", 9),
            relief=tk.FLAT,
            width=12
        )
        use_history_btn.pack(side=tk.LEFT)
    
    def add_to_expression(self, value):
        """向表达式中添加值"""
        if self.result_var.get() == "0" and value not in '/*-+':
            self.expression = value
        else:
            self.expression += value
        
        self.result_var.set(self.expression)
    
    def add_function(self, func):
        """添加函数到表达式"""
        # 如果当前显示的是结果，则清空表达式
        if self.result_var.get() == "0" or self.is_result_displayed():
            self.expression = ""
        
        self.expression += func
        self.result_var.set(self.expression)
    
    def clear_all(self):
        """清除所有"""
        self.expression = ""
        self.result_var.set("0")
    
    def clear_entry(self):
        """清除当前输入"""
        self.expression = ""
        self.result_var.set("0")
    
    def backspace(self):
        """退格删除"""
        if self.expression:
            self.expression = self.expression[:-1]
            self.result_var.set(self.expression if self.expression else "0")
    
    def calculate(self):
        """计算表达式"""
        if not self.expression:
            return
        
        try:
            # 将表达式的数学符号转换为Python可识别的符号
            expr = self.expression.replace('×', '*').replace('÷', '/')
            
            # 处理数学函数
            expr = expr.replace('sqrt', 'math.sqrt')
            expr = expr.replace('sin', 'math.sin')
            expr = expr.replace('cos', 'math.cos')
            expr = expr.replace('tan', 'math.tan')
            expr = expr.replace('log', 'math.log10')
            
            # 计算表达式
            result = eval(expr, {"__builtins__": None}, {"math": math})
            
            # 处理浮点数精度
            if isinstance(result, float):
                # 如果是整数，则显示为整数
                if result.is_integer():
                    result = int(result)
                else:
                    # 限制小数位数为10位
                    result = round(result, 10)
            
            # 保存到历史记录
            history_item = f"{self.expression} = {result}"
            self.history.insert(0, history_item)
            if len(self.history) > self.max_history:
                self.history.pop()
            
            # 更新历史记录显示
            self.update_history()
            
            # 显示结果
            self.result_var.set(str(result))
            self.expression = str(result)
            
        except ZeroDivisionError:
            messagebox.showerror("错误", "除以零错误！")
            self.clear_entry()
        except ValueError as e:
            messagebox.showerror("错误", f"数学错误: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"无效表达式: {str(e)}")
    
    def is_result_displayed(self):
        """检查当前显示的是否是计算结果"""
        # 简单检查：如果表达式为空但结果显示不为0，或者表达式与结果相同
        if not self.expression and self.result_var.get() != "0":
            return True
        
        # 检查结果是否只包含数字和小数点
        result = self.result_var.get()
        if re.match(r'^[-+]?[0-9]*\.?[0-9]+$', result):
            return True
        
        return False
    
    def update_history(self):
        """更新历史记录显示"""
        self.history_listbox.delete(0, tk.END)
        for item in self.history:
            self.history_listbox.insert(tk.END, item)
    
    def clear_history(self):
        """清空历史记录"""
        self.history = []
        self.update_history()
    
    def use_history(self):
        """使用选中的历史记录"""
        selection = self.history_listbox.curselection()
        if selection:
            item = self.history_listbox.get(selection[0])
            # 提取表达式部分（等号之前的部分）
            if '=' in item:
                expr = item.split('=')[0].strip()
                self.expression = expr
                self.result_var.set(expr)
    
    def toggle_theme(self):
        """切换主题"""
        self.dark_mode = not self.dark_mode
        self.setup_colors()
        
        # 重新创建界面
        for widget in self.root.winfo_children():
            widget.destroy()
        
        self.setup_ui()
    
    def key_press(self, event):
        """处理键盘事件"""
        key = event.char
        
        # 数字和运算符
        if key in '0123456789':
            self.add_to_expression(key)
        elif key in '+-*/':
            # 将*和/转换为计算器上的符号
            if key == '*':
                self.add_to_expression('×')
            elif key == '/':
                self.add_to_expression('÷')
            else:
                self.add_to_expression(key)
        elif key == '.':
            self.add_to_expression('.')
        elif key == '(' or key == ')':
            self.add_to_expression(key)
        elif key == '\r':  # 回车键
            self.calculate()
        elif key == '\x08':  # 退格键
            self.backspace()
        elif key == '\x1b':  # ESC键
            self.clear_all()
        elif key == 'c' or key == 'C':
            self.clear_entry()

def main():
    root = tk.Tk()
    app = ScientificCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
