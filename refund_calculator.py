import tkinter as tk
from tkinter import ttk, messagebox
import math
import json
import os
import threading
import re
import time

import pystray
from PIL import Image, ImageDraw
import keyboard

CONFIG_FILE = "config.json"


class ToastMessage:
    """气泡提示类，支持淡入淡出效果"""
    
    def __init__(self, parent):
        self.parent = parent
        self.toast_window = None
        self.alpha = 0.0
        
    def show(self, message="已复制", duration=1500):
        """显示气泡提示"""
        try:
            # 如果已有提示窗口，先关闭
            if self.toast_window:
                self.hide()
            
            # 创建气泡窗口
            self.toast_window = tk.Toplevel(self.parent)
            self.toast_window.overrideredirect(True)
            self.toast_window.configure(bg="#333333", highlightthickness=0)
            self.toast_window.attributes("-alpha", 0.0)
            self.toast_window.attributes("-topmost", True)
            
            # 设置窗口位置（在父窗口中央偏上）
            self.parent.update()  # 确保窗口尺寸已更新
            
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # 气泡尺寸
            toast_width = 100
            toast_height = 50
            
            # 计算位置（窗口中央偏上）
            x = parent_x + (parent_width - toast_width) // 2
            y = parent_y + (parent_height - toast_height) // 3  # 偏上位置
            
            self.toast_window.geometry(f"{toast_width}x{toast_height}+{x}+{y}")
            
            # 创建提示内容
            label = tk.Label(self.toast_window, text=message, 
                           bg="#333333", fg="white", font=("Arial", 14, "bold"),
                           padx=10, pady=5)
            label.pack(expand=True, fill="both")
            
            # 开始淡入动画
            self.fade_in()
            
            # 设置自动关闭
            self.toast_window.after(duration, self.fade_out)
            
        except Exception as e:
            print(f"显示气泡提示失败: {e}")
    
    def fade_in(self):
        """淡入动画"""
        try:
            if self.toast_window and self.alpha < 1.0:
                self.alpha += 0.1
                self.toast_window.attributes("-alpha", self.alpha)
                self.toast_window.after(30, self.fade_in)
        except Exception as e:
            print(f"淡入动画失败: {e}")
    
    def fade_out(self):
        """淡出动画"""
        try:
            if self.toast_window and self.alpha > 0.0:
                self.alpha -= 0.1
                self.toast_window.attributes("-alpha", self.alpha)
                self.toast_window.after(30, self.fade_out)
            else:
                self.hide()
        except Exception as e:
            print(f"淡出动画失败: {e}")
    
    def hide(self):
        """隐藏气泡提示"""
        try:
            if self.toast_window:
                self.toast_window.destroy()
                self.toast_window = None
                self.alpha = 0.0
        except Exception as e:
            print(f"隐藏气泡提示失败: {e}")


class RefundCalculator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("750x600")  # 增大窗口尺寸以适应更多内容
        self.root.overrideredirect(True)
        self.root.configure(bg="#f4f6f8")

        self.is_topmost = False

        # 创建气泡提示
        self.toast = ToastMessage(self.root)

        # 加载配置
        self.load_config()

        # 创建界面
        self.create_titlebar()
        self.create_ui()
        self.create_tray()
        self.register_hotkey()

        # 绑定窗口事件
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    # ---------------- 标题栏 ----------------
    def create_titlebar(self):
        """创建自定义标题栏"""
        try:
            bar = tk.Frame(self.root, bg="#2c3e50", height=36)
            bar.pack(fill="x")

            tk.Label(bar, text="山药赔付计算器", fg="white", bg="#2c3e50", 
                    font=("Arial", 10, "bold")).pack(side="left", padx=10)

            btns = tk.Frame(bar, bg="#2c3e50")
            btns.pack(side="right")

            # 置顶按钮
            self.pin_btn = tk.Button(btns, text="📌", bg="#2c3e50", fg="white", 
                                   bd=0, font=("Arial", 12), width=3, height=1,
                                   command=self.toggle_top)
            self.pin_btn.pack(side="left", padx=2)

            # 最小化按钮
            tk.Button(btns, text="─", bg="#2c3e50", fg="white", bd=0, 
                     font=("Arial", 12), width=3, height=1,
                     command=self.minimize).pack(side="left", padx=2)

            # 关闭按钮
            tk.Button(btns, text="✕", bg="#2c3e50", fg="white", bd=0, 
                     font=("Arial", 12), width=3, height=1,
                     command=self.hide_window).pack(side="left", padx=2)

            # 绑定标题栏拖动
            bar.bind("<Button-1>", self.start_move)
            bar.bind("<B1-Motion>", self.move)
            
        except Exception as e:
            print(f"创建标题栏失败: {e}")

    def start_move(self, event):
        """开始拖动窗口"""
        self.x = event.x
        self.y = event.y

    def move(self, event):
        """拖动窗口"""
        try:
            deltax = event.x - self.x
            deltay = event.y - self.y
            x = self.root.winfo_x() + deltax
            y = self.root.winfo_y() + deltay
            self.root.geometry(f"+{x}+{y}")
        except Exception as e:
            print(f"拖动窗口失败: {e}")

    def toggle_top(self):
        """切换窗口置顶状态"""
        try:
            self.is_topmost = not self.is_topmost
            self.root.attributes("-topmost", self.is_topmost)
            self.pin_btn.config(text="📍" if self.is_topmost else "📌")
        except Exception as e:
            print(f"切换置顶状态失败: {e}")

    def minimize(self):
        """最小化窗口"""
        try:
            self.root.iconify()
        except Exception as e:
            print(f"最小化失败: {e}")

    def hide_window(self):
        """隐藏窗口到托盘"""
        try:
            self.root.withdraw()
        except Exception as e:
            print(f"隐藏窗口失败: {e}")

    def quit_app(self):
        """退出应用程序"""
        try:
            if hasattr(self, 'tray'):
                self.tray.stop()
            keyboard.unhook_all()
            self.root.destroy()
        except Exception as e:
            print(f"退出应用失败: {e}")

    # ---------------- 用户界面 ----------------
    def create_ui(self):
        """创建主界面"""
        try:
            # 顶部区域 - 实付金额
            top_frame = tk.Frame(self.root, bg="white", padx=15, pady=10)
            top_frame.pack(fill="x", padx=10, pady=10)

            tk.Label(top_frame, text="实付金额（元）", bg="white", 
                    font=("Arial", 12)).pack(anchor="w")
            
            self.amount_var = tk.StringVar()
            amount_entry = tk.Entry(top_frame, textvariable=self.amount_var, 
                                   font=("Arial", 12))
            amount_entry.pack(fill="x", pady=5)
            amount_entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))
            amount_entry.bind("<KeyRelease>", lambda e: self.schedule_update())

            # 金额验证标签
            self.amount_warning = tk.Label(top_frame, text="", fg="red", 
                                          bg="white", font=("Arial", 10))
            self.amount_warning.pack(anchor="w")

            # 标签页区域
            self.tabs = ttk.Notebook(self.root)
            self.tabs.pack(fill="both", expand=True, padx=10, pady=5)

            # 创建三个标签页
            self.create_weight_tab()
            self.create_rotten_tab()
            self.create_ratio_tab()

            # 绑定标签页切换事件
            self.tabs.bind("<<NotebookTabChanged>>", lambda e: self.schedule_update())

            # 底部区域 - 结果和话术
            self.create_bottom_section()
            
        except Exception as e:
            print(f"创建界面失败: {e}")

    def create_weight_tab(self):
        """创建少重赔付标签页"""
        try:
            tab = tk.Frame(self.tabs, padx=10, pady=10)
            self.tabs.add(tab, text="少重赔付")

            # 购买规格选择
            spec_frame = tk.Frame(tab)
            spec_frame.pack(fill="x", pady=5)
            
            tk.Label(spec_frame, text="购买规格：", font=("Arial", 12)).pack(side="left")
            
            self.spec_var = tk.StringVar(value="5斤")
            spec_combo = ttk.Combobox(spec_frame, textvariable=self.spec_var,
                                     values=["3斤", "5斤", "7斤", "10斤"], 
                                     state="readonly", font=("Arial", 12))
            spec_combo.pack(side="left", padx=10)
            spec_combo.bind("<<ComboboxSelected>>", lambda e: self.schedule_update())

            # 收到重量输入
            weight_frame = tk.Frame(tab)
            weight_frame.pack(fill="x", pady=5)
            
            tk.Label(weight_frame, text="收到重量：", font=("Arial", 12)).pack(side="left")
            
            self.weight_var = tk.StringVar()
            weight_entry = tk.Entry(weight_frame, textvariable=self.weight_var, 
                                  font=("Arial", 12))
            weight_entry.pack(side="left", padx=10, fill="x", expand=True)
            weight_entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))
            
            # 单位识别结果显示
            self.unit_recognition_label = tk.Label(weight_frame, text="", 
                                                  fg="blue", font=("Arial", 10))
            self.unit_recognition_label.pack(side="left", padx=10)

            # 绑定重量变化事件
            self.weight_var.trace("w", lambda *a: self.schedule_update())
            
        except Exception as e:
            print(f"创建少重标签页失败: {e}")

    def create_rotten_tab(self):
        """创建腐烂/变质标签页"""
        try:
            tab = tk.Frame(self.tabs, padx=10, pady=10)
            self.tabs.add(tab, text="腐烂/变质")

            # 总根数输入
            total_frame = tk.Frame(tab)
            total_frame.pack(fill="x", pady=5)
            
            tk.Label(total_frame, text="总根数：", font=("Arial", 12)).pack(side="left")
            
            self.total_roots_var = tk.StringVar()
            total_entry = tk.Entry(total_frame, textvariable=self.total_roots_var, 
                                  font=("Arial", 12))
            total_entry.pack(side="left", padx=10, fill="x", expand=True)
            total_entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))

            # 坏根数输入
            bad_frame = tk.Frame(tab)
            bad_frame.pack(fill="x", pady=5)
            
            tk.Label(bad_frame, text="坏根数：", font=("Arial", 12)).pack(side="left")
            
            self.bad_roots_var = tk.StringVar()
            bad_entry = tk.Entry(bad_frame, textvariable=self.bad_roots_var, 
                               font=("Arial", 12))
            bad_entry.pack(side="left", padx=10, fill="x", expand=True)
            bad_entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))

            # 绑定根数变化事件
            self.total_roots_var.trace("w", lambda *a: self.schedule_update())
            self.bad_roots_var.trace("w", lambda *a: self.schedule_update())
            
        except Exception as e:
            print(f"创建腐烂标签页失败: {e}")

    def create_ratio_tab(self):
        """创建比例赔付标签页"""
        try:
            tab = tk.Frame(self.tabs, padx=10, pady=10)
            self.tabs.add(tab, text="比例赔付")

            # 损坏比例输入
            ratio_frame = tk.Frame(tab)
            ratio_frame.pack(fill="x", pady=5)
            
            tk.Label(ratio_frame, text="损坏比例：", font=("Arial", 12)).pack(side="left")
            
            self.damage_ratio_var = tk.StringVar()
            ratio_entry = tk.Entry(ratio_frame, textvariable=self.damage_ratio_var, 
                                 font=("Arial", 12))
            ratio_entry.pack(side="left", padx=10, fill="x", expand=True)
            ratio_entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))
            
            tk.Label(ratio_frame, text="%（1-100之间的数字）", 
                    font=("Arial", 10)).pack(side="left", padx=10)

            # 绑定比例变化事件
            self.damage_ratio_var.trace("w", lambda *a: self.schedule_update())
            
        except Exception as e:
            print(f"创建比例标签页失败: {e}")

    def create_bottom_section(self):
        """创建底部结果和话术区域"""
        try:
            bottom_frame = tk.Frame(self.root, bg="white", padx=15, pady=10)
            bottom_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # 结果显示
            self.result_var = tk.StringVar(value="-")
            result_label = tk.Label(bottom_frame, textvariable=self.result_var, 
                                   fg="#27ae60", bg="white", font=("Arial", 12, "bold"))
            result_label.pack(anchor="w", pady=(0, 10))

            # 话术区域
            tk.Label(bottom_frame, text="话术（点击自动复制）", bg="white", 
                    font=("Arial", 12)).pack(anchor="w")
            
            # 话术文本框和滚动条
            text_frame = tk.Frame(bottom_frame)
            text_frame.pack(fill="both", expand=True, pady=5)
            
            self.speech_text = tk.Text(text_frame, height=6, wrap=tk.WORD, 
                                      font=("Arial", 11))
            
            # 添加滚动条
            scrollbar = tk.Scrollbar(text_frame, orient="vertical", 
                                   command=self.speech_text.yview)
            self.speech_text.configure(yscrollcommand=scrollbar.set)
            
            self.speech_text.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # 绑定点击复制事件
            self.speech_text.bind("<Button-1>", self.copy_speech)

            # 按钮区域
            button_frame = tk.Frame(bottom_frame)
            button_frame.pack(fill="x", pady=10)
            
            # 设置按钮
            tk.Button(button_frame, text="设置", command=self.open_settings, 
                     font=("Arial", 10), width=10).pack(side="right", padx=5)
            
        except Exception as e:
            print(f"创建底部区域失败: {e}")

    # ---------------- 工具函数 ----------------
    def ceil_half(self, value):
        """向上取0.5"""
        try:
            return math.ceil(value * 2) / 2
        except:
            return 0.0

    def parse_weight(self, text, spec_weight):
        """解析重量输入，支持智能单位识别"""
        try:
            # 提取数字
            numbers = re.findall(r"\d+\.?\d*", text)
            if not numbers:
                return 0, ""
                
            num = float(numbers[0])
            original_text = text.strip()
            
            # 单位识别
            if "kg" in text.lower() or "公斤" in text:
                weight_g = num * 1000
                unit = "公斤"
            elif "斤" in text:
                weight_g = num * 500
                unit = "斤"
            elif "g" in text.lower():
                weight_g = num
                unit = "克"
            else:
                # 智能识别：尝试不同单位，选择最接近规格的
                options = [
                    (num, "克"),
                    (num * 500, "斤"), 
                    (num * 1000, "公斤")
                ]
                # 选择与规格重量最接近的单位
                best_match = min(options, key=lambda x: abs(x[0] - spec_weight))
                weight_g, unit = best_match
            
            return weight_g, unit
            
        except Exception as e:
            print(f"解析重量失败: {e}")
            return 0, ""

    def validate_input(self, value, input_type="number"):
        """输入验证"""
        try:
            if not value.strip():
                return False, "输入不能为空"
                
            if input_type == "number":
                num = float(value)
                if num <= 0:
                    return False, "请输入大于0的数字"
                return True, ""
                    
            elif input_type == "integer":
                if not value.isdigit():
                    return False, "请输入整数"
                num = int(value)
                if num <= 0:
                    return False, "请输入大于0的整数"
                return True, ""
                    
            elif input_type == "percentage":
                num = float(value)
                if num <= 0 or num > 100:
                    return False, "请输入1-100之间的数字"
                return True, ""
                    
        except ValueError:
            return False, "请输入有效的数字"
        except Exception as e:
            return False, f"输入验证失败: {e}"

    # ---------------- 核心计算逻辑 ----------------
    def schedule_update(self):
        """调度更新，添加延迟避免频繁计算"""
        if hasattr(self, '_update_job'):
            self.root.after_cancel(self._update_job)
        self._update_job = self.root.after(300, self.update_all)  # 300ms延迟

    def update_all(self):
        """更新所有计算结果"""
        try:
            # 验证实付金额
            amount_text = self.amount_var.get().strip()
            if not amount_text:
                self.result_var.set("-")
                self.speech_text.delete("1.0", tk.END)
                self.amount_warning.config(text="")
                return
                
            is_valid, message = self.validate_input(amount_text, "number")
            if not is_valid:
                self.amount_warning.config(text=message)
                self.result_var.set("-")
                self.speech_text.delete("1.0", tk.END)
                return
            else:
                self.amount_warning.config(text="")
            
            amount = float(amount_text)
            
            # 获取当前标签页
            current_tab = self.tabs.index(self.tabs.select())
            
            if current_tab == 0:  # 少重赔付
                self.calculate_weight_refund(amount)
            elif current_tab == 1:  # 腐烂/变质
                self.calculate_rotten_refund(amount)
            else:  # 比例赔付
                self.calculate_ratio_refund(amount)
                
        except Exception as e:
            print(f"更新计算失败: {e}")
            self.result_var.set("计算错误")
            self.speech_text.delete("1.0", tk.END)
            self.speech_text.insert("1.0", "计算过程中出现错误，请检查输入数据。")

    def calculate_weight_refund(self, amount):
        """计算少重赔付"""
        try:
            # 规格映射
            spec_map = {"3斤": 1500, "5斤": 2500, "7斤": 3500, "10斤": 5000}
            spec_text = self.spec_var.get()
            spec_weight = spec_map.get(spec_text, 2500)
            
            # 解析收到重量
            weight_text = self.weight_var.get().strip()
            if not weight_text:
                self.result_var.set("-")
                self.speech_text.delete("1.0", tk.END)
                self.unit_recognition_label.config(text="")
                return
            
            received_weight, unit = self.parse_weight(weight_text, spec_weight)
            
            # 更新单位识别显示
            if unit:
                self.unit_recognition_label.config(text=f"识别为: {weight_text} → {received_weight}克")
            else:
                self.unit_recognition_label.config(text="")
            
            # 计算缺失重量
            missing_weight = spec_weight - received_weight
            
            # 判断是否在误差范围内
            if missing_weight <= 50:
                speech = """亲，这边跟您说明一下哈～

我们发货都是按【净重（可食用部分）】严格称重打包的，保证您收到的都是能吃的部分。

不过山药属于生鲜产品，在运输过程中会有轻微自然失水，一般在几十克左右（差不多一个山药头的重量），这个属于正常范围内的误差哦～

实际您收到的整体是达标的，可以放心食用哈～我们这边也是一直按足斤足两标准发货的 🙏"""
                self.result_var.set("正常误差范围内（无需赔付）")
                self.set_speech(speech)
                return
            
            # 计算赔付金额
            refund_amount = (missing_weight / spec_weight) * amount
            final_amount = self.ceil_half(refund_amount)
            
            speech = f"""亲，这边已经帮您核算过了～

您反馈的重量情况确实存在偏差，我们这边是按比例给您做补偿的：

应赔金额是：{refund_amount:.2f}元  
我们统一是按平台规则【向上取0.5】进行处理，最终为您申请：{final_amount:.2f}元

这边已经帮您申请好了，请您放心查收～

也感谢您的理解，我们后续也会继续优化包装，尽量减少运输损耗 🙏"""
            
            self.result_var.set(f"应赔 {refund_amount:.2f} → 实赔 {final_amount:.2f}")
            self.set_speech(speech)
            
        except Exception as e:
            print(f"少重赔付计算失败: {e}")

    def calculate_rotten_refund(self, amount):
        """计算腐烂/变质赔付"""
        try:
            # 验证输入
            total_text = self.total_roots_var.get().strip()
            bad_text = self.bad_roots_var.get().strip()
            
            if not total_text or not bad_text:
                self.result_var.set("-")
                self.speech_text.delete("1.0", tk.END)
                return
            
            is_valid_total, msg_total = self.validate_input(total_text, "integer")
            is_valid_bad, msg_bad = self.validate_input(bad_text, "integer")
            
            if not is_valid_total or not is_valid_bad:
                self.result_var.set("输入错误")
                self.speech_text.delete("1.0", tk.END)
                self.speech_text.insert("1.0", f"请检查输入: {msg_total if not is_valid_total else msg_bad}")
                return
            
            total_roots = int(total_text)
            bad_roots = int(bad_text)
            
            if bad_roots > total_roots:
                self.result_var.set("输入错误")
                self.speech_text.delete("1.0", tk.END)
                self.speech_text.insert("1.0", "坏根数不能大于总根数")
                return
            
            # 计算赔付比例和金额
            ratio = bad_roots / total_roots
            refund_amount = amount * ratio
            final_amount = self.ceil_half(refund_amount)
            
            speech = f"""亲，这边跟您说明一下我们的售后标准哈～

生鲜类商品主要是按【是否影响食用】来判定的：

👉 如果是表皮磕碰、轻微损伤  
这种不影响内部品质，削皮后是可以正常食用的，不属于质量问题范围哦～

👉 如果出现内部发黑、腐烂、异味等情况  
这种是可以给您做售后的 👍

根据您反馈情况，这边核算：

腐烂比例：{ratio:.0%}  
应赔金额：{refund_amount:.2f}元  
最终补偿：{final_amount:.2f}元

如果方便的话，可以拍一下切开后的内部情况，我们可以帮您进一步确认处理～"""
            
            self.result_var.set(f"{ratio:.0%} → {final_amount:.2f}")
            self.set_speech(speech)
            
        except Exception as e:
            print(f"腐烂赔付计算失败: {e}")

    def calculate_ratio_refund(self, amount):
        """计算比例赔付"""
        try:
            # 验证输入
            ratio_text = self.damage_ratio_var.get().strip()
            if not ratio_text:
                self.result_var.set("-")
                self.speech_text.delete("1.0", tk.END)
                return
            
            is_valid, message = self.validate_input(ratio_text, "percentage")
            if not is_valid:
                self.result_var.set("输入错误")
                self.speech_text.delete("1.0", tk.END)
                self.speech_text.insert("1.0", message)
                return
            
            ratio = float(ratio_text) / 100
            refund_amount = amount * ratio
            final_amount = self.ceil_half(refund_amount)
            
            speech = f"""亲，实在不好意思给您带来不好的体验了 🙏

根据您反馈的情况，这边已经帮您按比例核算：

赔付比例：{ratio:.0%}  
应赔金额：{refund_amount:.2f}元  
最终为您申请：{final_amount:.2f}元（按平台规则向上取整）

这边已经为您处理完成，您注意查收一下～

我们后续也会继续优化发货品质，争取给您更好的体验 🌿"""
            
            self.result_var.set(f"{final_amount:.2f}")
            self.set_speech(speech)
            
        except Exception as e:
            print(f"比例赔付计算失败: {e}")

    def set_speech(self, text):
        """设置话术内容"""
        try:
            self.speech_text.delete("1.0", tk.END)
            self.speech_text.insert("1.0", text)
        except Exception as e:
            print(f"设置话术失败: {e}")

    def copy_speech(self, event):
        """复制话术到剪贴板"""
        try:
            text = self.speech_text.get("1.0", tk.END).strip()
            if text:
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                # 使用气泡提示替代弹窗
                self.toast.show("已复制")
        except Exception as e:
            print(f"复制话术失败: {e}")
            # 复制失败时也使用气泡提示
            self.toast.show("复制失败")

    # ---------------- 设置功能 ----------------
    def open_settings(self):
        """打开设置窗口"""
        try:
            settings_window = tk.Toplevel(self.root)
            settings_window.title("设置")
            settings_window.geometry("300x200")
            settings_window.resizable(False, False)
            
            # 置顶窗口
            settings_window.transient(self.root)
            settings_window.grab_set()
            
            # 快捷键设置
            tk.Label(settings_window, text="全局快捷键", 
                    font=("Arial", 12, "bold")).pack(pady=10)
            
            hotkey_frame = tk.Frame(settings_window)
            hotkey_frame.pack(pady=10)
            
            hotkey_var = tk.StringVar(value=self.config.get("hotkey", "ctrl+shift+r"))
            hotkey_entry = tk.Entry(hotkey_frame, textvariable=hotkey_var, 
                                   font=("Arial", 11), width=20)
            hotkey_entry.pack(pady=5)
            
            tk.Label(hotkey_frame, text="格式如: ctrl+shift+r", 
                    font=("Arial", 9), fg="gray").pack()
            
            def save_settings():
                try:
                    # 卸载旧热键
                    keyboard.unhook_all()
                    
                    # 保存新配置
                    self.config["hotkey"] = hotkey_var.get()
                    self.save_config()
                    
                    # 注册新热键
                    self.register_hotkey()
                    
                    messagebox.showinfo("提示", "设置已保存并生效")
                    settings_window.destroy()
                    
                except Exception as e:
                    print(f"保存设置失败: {e}")
                    messagebox.showerror("错误", "保存失败，请检查快捷键格式")
            
            # 保存按钮
            tk.Button(settings_window, text="保存", command=save_settings,
                     font=("Arial", 11), width=10).pack(pady=10)
            
        except Exception as e:
            print(f"打开设置窗口失败: {e}")

    # ---------------- 托盘功能 ----------------
    def create_tray(self):
        """创建系统托盘图标"""
        try:
            # 创建托盘图标
            image = Image.new("RGB", (64, 64), "#27ae60")
            draw = ImageDraw.Draw(image)
            draw.text((22, 20), "￥", fill="white", font_size=20)
            
            # 创建菜单
            menu = pystray.Menu(
                pystray.MenuItem("显示", lambda: self.root.deiconify()),
                pystray.MenuItem("退出", self.quit_app)
            )
            
            # 创建托盘图标
            self.tray = pystray.Icon("refund_calculator", image, "退款计算器", menu)
            
            # 在后台线程中运行托盘
            threading.Thread(target=self.tray.run, daemon=True).start()
            
        except Exception as e:
            print(f"创建托盘失败: {e}")

    # ---------------- 热键功能 ----------------
    def register_hotkey(self):
        """注册全局热键"""
        try:
            hotkey = self.config.get("hotkey", "ctrl+shift+r")
            keyboard.add_hotkey(hotkey, self.toggle_window)
        except Exception as e:
            print(f"注册热键失败: {e}")

    def toggle_window(self):
        """切换窗口显示/隐藏"""
        try:
            if self.root.state() == "withdrawn":
                self.root.deiconify()
                self.root.lift()
            else:
                self.root.withdraw()
        except Exception as e:
            print(f"切换窗口失败: {e}")

    # ---------------- 配置管理 ----------------
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            else:
                self.config = {"hotkey": "ctrl+shift+r"}
        except Exception as e:
            print(f"加载配置失败: {e}")
            self.config = {"hotkey": "ctrl+shift+r"}

    def save_config(self):
        """保存配置文件"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    # ---------------- 应用程序运行 ----------------
    def run(self):
        """运行应用程序"""
        try:
            self.root.mainloop()
        except Exception as e:
            print(f"应用程序运行失败: {e}")


if __name__ == "__main__":
    # 检查依赖库
    try:
        app = RefundCalculator()
        app.run()
    except ImportError as e:
        print(f"缺少依赖库: {e}")
        print("请运行以下命令安装依赖:")
        print("conda activate js")
        print("pip install pystray keyboard pillow")
    except Exception as e:
        print(f"应用程序启动失败: {e}")