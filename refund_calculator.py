import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os

DEBUG_MODE = True

# ==================== 默认退款分类数据 ====================
DEFAULT_CATEGORIES = [
    {"name": "轻度腐烂", "desc": "表面小面积软点、水渍斑，削掉可食用", "ratio": 0.30, "scope": "单根局部软点", "group": "腐烂", "level": 1, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "中度腐烂", "desc": "局部腐烂面积＜1/3，轻微渗水或霉点", "ratio": 0.60, "scope": "单根腐烂小于1/3", "group": "腐烂", "level": 2, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "重度腐烂", "desc": "大面积腐烂、发黑发霉超过1/2", "ratio": 1.00, "scope": "单根腐烂超过一半", "group": "腐烂", "level": 3, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "轻度发芽", "desc": "芽长＜1cm，山药体未变软", "ratio": 0.10, "scope": "芽点刚萌动", "group": "发芽", "level": 1, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "中度发芽", "desc": "芽长1-3cm，根体开始变软", "ratio": 0.30, "scope": "芽长1到3厘米", "group": "发芽", "level": 2, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "重度发芽", "desc": "芽长＞3cm，根体明显干缩", "ratio": 0.60, "scope": "芽长超过3厘米", "group": "发芽", "level": 3, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "轻度断裂", "desc": "断成两截，断面整齐新鲜", "ratio": 0.20, "scope": "断裂1到2根", "group": "断裂", "level": 1, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "中度断裂", "desc": "断成两截以上，断面轻微氧化", "ratio": 0.40, "scope": "断裂2到3根", "group": "断裂", "level": 2, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
    {"name": "重度断裂", "desc": "断成多截、断面发黑", "ratio": 0.60, "scope": "断裂3根及以上", "group": "断裂", "level": 3, "final_increase": 5,
     "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
     "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"},
]

DEFAULT_GROUPS = ["腐烂", "发芽", "断裂"]

CONFIG_FILE = "refund_categories.json"
LOG_FILE = "refund_log.json"
GROUPS_FILE = "refund_groups.json"


# ==================== 数据管理 ====================
def load_groups():
    if os.path.exists(GROUPS_FILE):
        try:
            with open(GROUPS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return list(DEFAULT_GROUPS)


def save_groups(groups):
    with open(GROUPS_FILE, "w", encoding="utf-8") as f:
        json.dump(groups, f, ensure_ascii=False, indent=2)


def load_categories():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    # 默认数据深拷贝
    return [dict(c) for c in DEFAULT_CATEGORIES]


def save_categories(categories):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(categories, f, ensure_ascii=False, indent=2)


def load_log():
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return []


def save_log(logs):
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)


def add_log_entry(reason_name, amount, ratio, level):
    logs = load_log()
    logs.append({
        "reason": reason_name,
        "amount": amount,
        "ratio": ratio,
        "level": level,
        "disagree_count": 1
    })
    save_log(logs)


def render_template(template, **kwargs):
    for key, value in kwargs.items():
        template = template.replace(f"{{{key}}}", str(value))
    return template


# ==================== 主应用类 ====================
class RefundApp:
    def __init__(self, root):
        self.root = root
        self.root.title("生鲜山药售后 · 话术生成器")
        self.root.geometry("900x680")
        self.root.minsize(800, 600)

        self.categories = load_categories()
        self.groups = load_groups()
        self.bubble_buttons = []
        self.amount_var = tk.StringVar()
        self.history_text = None
        self.current_category = None
        self.current_ratio = 0
        self.current_level = 0

        self.build_ui()
        self.refresh_bubbles()
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 900
        window_height = 680
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # ---------- 主界面 ----------
    def build_ui(self):
        # 顶部说明
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)
        ttk.Label(top_frame, text="🛠️ 山药售后问题快速话术生成", font=("微软雅黑", 14, "bold")).pack(anchor=tk.W)
        ttk.Label(top_frame, text="鼠标悬停气泡查看退款范畴 · 点击气泡自动生成话术", foreground="gray").pack(anchor=tk.W)

        # 中间：气泡容器 + 右侧设置按钮
        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 左侧气泡区域（带滚动条）
        left_frame = ttk.LabelFrame(middle_frame, text="退款原因分类", padding=8)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(left_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.bubble_canvas = canvas
        self.bubble_frame = ttk.Frame(canvas)

        self.bubble_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.bubble_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧设置面板
        right_frame = ttk.Frame(middle_frame, width=160)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))
        right_frame.pack_propagate(False)

        ttk.Button(right_frame, text="⚙️ 设置管理", command=self.open_settings).pack(fill=tk.X, pady=7, ipadx=1)
        ttk.Button(right_frame, text="🏷️ 分类管理", command=self.open_group_manager).pack(fill=tk.X, pady=7, ipadx=1)
        ttk.Separator(right_frame).pack(fill=tk.X, pady=6)

        ttk.Label(right_frame, text="实付金额 (元)：").pack(anchor=tk.W)
        self.amount_entry = ttk.Entry(right_frame, textvariable=self.amount_var, font=("微软雅黑", 11))
        self.amount_entry.pack(fill=tk.X, pady=2)
        self.amount_entry.bind("<FocusIn>", lambda e: self.amount_entry.select_range(0, tk.END))

        ttk.Separator(right_frame).pack(fill=tk.X, pady=8)

        ttk.Label(right_frame, text="生成话术预览：").pack(anchor=tk.W)
        self.history_text = tk.Text(right_frame, height=14, width=28, font=("微软雅黑", 9), wrap=tk.WORD,
                                    relief=tk.SUNKEN, borderwidth=1)
        self.history_text.pack(fill=tk.BOTH, expand=True, pady=2)
        self.disagree_btn = ttk.Button(right_frame, text="顾客不同意 → 升级方案", command=self.on_disagree)
        self.disagree_btn.pack(fill=tk.X, pady=4)

        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=3)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------- 刷新气泡按钮 ----------
    def refresh_bubbles(self):
        for btn in self.bubble_buttons:
            btn.destroy()
        self.bubble_buttons.clear()

        for widget in self.bubble_frame.winfo_children():
            widget.destroy()

        row = 0
        for group_name in self.groups:
            group_cats = [c for c in self.categories if c.get("group", "") == group_name]
            group_cats.sort(key=lambda x: x.get("level", 0))

            for i, cat in enumerate(group_cats):
                name = cat["name"]
                desc = cat["desc"]
                ratio = cat.get("ratio", 0)

                tooltip_text = f"{name}\n参考比例：{int(ratio*100)}%"

                btn = ttk.Button(self.bubble_frame, text=name, width=14,
                                 command=lambda c=cat: self.on_bubble_click(c))
                btn.grid(row=row, column=i % 3, padx=4, pady=3, sticky="ew")
                self.bubble_buttons.append(btn)
                self.create_tooltip(btn, tooltip_text)

            if group_cats:
                row += 1

        self.bubble_frame.update_idletasks()
        self.bubble_canvas.configure(scrollregion=self.bubble_canvas.bbox("all"))

    def create_tooltip(self, widget, text):
        """简陋但可用的悬停提示"""
        tooltip = None

        def on_enter(event):
            nonlocal tooltip
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            tooltip = tk.Toplevel(widget)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(tooltip, text=text, background="#ffffcc", relief=tk.SOLID, borderwidth=1,
                              font=("微软雅黑", 9), padding=6, wraplength=220)
            label.pack()

        def on_leave(event):
            nonlocal tooltip
            if tooltip:
                tooltip.destroy()
                tooltip = None

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    # ---------- 气泡点击事件 ----------
    def on_bubble_click(self, category):
        try:
            amount = float(self.amount_var.get())
        except ValueError:
            messagebox.showwarning("输入错误", "请输入有效的实付金额（数字）")
            return

        name = category["name"]
        ratio = category.get("ratio", 0)
        scope = category.get("scope", "")
        group = category.get("group", "")
        level = category.get("level", 1)

        self.current_category = category
        self.current_ratio = ratio
        self.current_level = level
        self.current_amount = amount

        refund_money = round(amount * ratio, 2)
        template = category.get("template", "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？")
        script = render_template(template, name=name, desc=desc, ratio=int(ratio*100), money=refund_money, final_ratio=int(ratio*100))

        self.history_text.delete("1.0", tk.END)
        self.history_text.insert(tk.END, script)
        self.root.clipboard_clear()
        self.root.clipboard_append(script)
        self.status_var.set(f"已生成【{name}】话术，赔付 {refund_money} 元（已自动复制）")

    def on_disagree(self):
        if not self.current_category:
            messagebox.showinfo("提示", "请先点击气泡生成话术")
            return

        name = self.current_category["name"]
        group = self.current_category.get("group", "")
        current_level = self.current_level
        final_increase = self.current_category.get("final_increase", 5)
        amount = self.current_amount

        same_group_cats = [c for c in self.categories if c.get("group", "") == group]
        same_group_cats.sort(key=lambda x: x.get("level", 0))

        next_cat = None
        is_final = False

        for cat in same_group_cats:
            if cat.get("level", 0) > current_level:
                next_cat = cat
                break

        if next_cat is None:
            current_ratio = self.current_ratio
            final_ratio = min(current_ratio + final_increase / 100, 1.0)
            is_final = True
        else:
            final_ratio = next_cat.get("ratio", self.current_ratio)
            next_cat = cat

        refund_money = round(amount * final_ratio, 2)

        if is_final:
            template_upgrade = self.current_category.get("template_upgrade", "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？")
            script = render_template(template_upgrade, name=name, desc=self.current_category.get("desc", ""),
                                     ratio=int(self.current_ratio*100), money=refund_money, final_ratio=int(final_ratio*100))
        else:
            template = next_cat.get("template", "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？")
            desc = next_cat.get("desc", "")
            script = render_template(template, name=next_cat["name"], desc=desc, ratio=int(final_ratio*100), money=refund_money, final_ratio=int(final_ratio*100))
            self.current_category = next_cat

        self.current_ratio = final_ratio
        self.current_level = current_level + 1 if not is_final else 99

        add_log_entry(name, amount, self.current_ratio, self.current_level)

        self.history_text.delete("1.0", tk.END)
        self.history_text.insert(tk.END, script)
        self.root.clipboard_clear()
        self.root.clipboard_append(script)
        self.status_var.set(f"已升级方案，【{name}】赔付 {refund_money} 元（已自动复制）")

    # ---------- 设置窗口 ----------
    def open_settings(self):
        SettingsDialog(self.root, self.categories, self.groups, self.on_categories_updated)

    def open_group_manager(self):
        GroupManagerDialog(self.root, self.on_groups_updated)

    def on_groups_updated(self, new_groups):
        self.groups = new_groups
        save_groups(self.groups)
        self.refresh_bubbles()
        self.status_var.set("退款分类已更新")

    def on_categories_updated(self, new_categories):
        self.categories = new_categories
        save_categories(self.categories)
        self.refresh_bubbles()
        self.status_var.set("退款分类已更新")


# ==================== 设置对话框 ====================
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, categories, groups, callback):
        super().__init__(parent)
        self.title("退款方案管理 · 编辑/添加/删除")
        self.geometry("1100x650")
        self.minsize(900, 550)

        self.categories = [dict(c) for c in categories]
        self.groups = list(groups)
        self.callback = callback
        self.tree = None
        self.detail_frame = None
        self.entry_vars = {}

        self.withdraw()
        self.fix_duplicate_levels()
        self.build_tree()
        self.build_detail_panel()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 1100
        window_height = 650
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.deiconify()

    def fix_duplicate_levels(self):
        KEYWORD_ORDER = {"轻": 1, "中": 2, "重": 3, "轻": 1, "中": 2, "重": 3}

        for group in self.groups:
            cats_in_group = [c for c in self.categories if c.get("group", "") == group]
            cats_in_group.sort(key=lambda c: KEYWORD_ORDER.get(c.get("name", "")[0], 99) if c.get("name", "") else 99)

            for idx, cat in enumerate(cats_in_group):
                cat["level"] = idx + 1

        if hasattr(self, "callback"):
            self.callback(self.categories)

    def build_tree(self):
        frame = ttk.LabelFrame(self, text="现有分类列表", padding=8)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)

        columns = ("name", "ratio", "level")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("name", text="退款原因")
        self.tree.heading("ratio", text="赔付比例(%)")
        self.tree.heading("level", text="等级")
        self.tree.column("name", width=120)
        self.tree.column("ratio", width=80)
        self.tree.column("level", width=50)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.refresh_tree()

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=4)
        ttk.Button(btn_frame, text="➕ 新增方案", command=self.add_category).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="🗑️ 删除选中", command=self.delete_category).pack(side=tk.RIGHT, padx=3)

    def build_detail_panel(self):
        panel = ttk.LabelFrame(self, text="编辑选中分类详情", padding=12)
        panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=8, pady=8, expand=False)

        self.entry_vars["name"] = tk.StringVar()
        self.entry_vars["ratio"] = tk.StringVar()
        self.entry_vars["desc"] = tk.StringVar()
        self.entry_vars["group"] = tk.StringVar()
        self.entry_vars["level"] = tk.StringVar()
        self.entry_vars["final_increase"] = tk.StringVar()

        self.status_label = ttk.Label(panel, text="", foreground="green", font=("微软雅黑", 9))

        ttk.Label(panel, text="退款原因名称：").pack(anchor=tk.W, pady=(6, 1))
        name_entry = ttk.Entry(panel, textvariable=self.entry_vars["name"], font=("微软雅黑", 10), width=32)
        name_entry.pack(fill=tk.X, pady=1)
        name_entry.bind("<KeyRelease>", lambda e: self.auto_save())

        ttk.Label(panel, text="赔付比例 (百分比数字，如30)：").pack(anchor=tk.W, pady=(6, 1))
        ratio_entry = ttk.Entry(panel, textvariable=self.entry_vars["ratio"], font=("微软雅黑", 10), width=32)
        ratio_entry.pack(fill=tk.X, pady=1)
        ratio_entry.bind("<KeyRelease>", lambda e: self.auto_save())

        ttk.Label(panel, text="所属分类：").pack(anchor=tk.W, pady=(6, 1))
        self.group_combo = ttk.Combobox(panel, textvariable=self.entry_vars["group"], font=("微软雅黑", 10), width=30, state="readonly")
        self.group_combo.pack(fill=tk.X, pady=1)
        self.group_combo.bind("<<ComboboxSelected>>", lambda e: self.on_group_changed())

        ttk.Label(panel, text="方案等级：").pack(anchor=tk.W, pady=(6, 1))
        self.level_combo = ttk.Combobox(panel, textvariable=self.entry_vars["level"], font=("微软雅黑", 10), width=30, state="readonly")
        self.level_combo.pack(fill=tk.X, pady=1)
        self.level_combo.bind("<<ComboboxSelected>>", lambda e: self.on_level_changed())

        self.top_level_label = ttk.Label(panel, text="🏆 最高赔偿方案", foreground="red", font=("微软雅黑", 10, "bold"))

        self.final_label = ttk.Label(panel, text="最终方案增加百分比：")
        self.final_label.pack(anchor=tk.W, pady=(6, 1))
        self.final_entry = ttk.Entry(panel, textvariable=self.entry_vars["final_increase"], font=("微软雅黑", 10), width=32)
        self.final_entry.pack(fill=tk.X, pady=1)
        self.final_entry.bind("<KeyRelease>", lambda e: self.auto_save())

        ttk.Label(panel, text="详细描述 (悬停补充)：").pack(anchor=tk.W, pady=(6, 1))
        desc_entry = ttk.Entry(panel, textvariable=self.entry_vars["desc"], font=("微软雅黑", 10), width=32)
        desc_entry.pack(fill=tk.X, pady=1)
        desc_entry.bind("<KeyRelease>", lambda e: self.auto_save())

        self.status_label.pack(anchor=tk.W, pady=(10, 0))

        btn_frame = ttk.Frame(panel)
        btn_frame.pack(fill=tk.X, pady=12)
        ttk.Button(btn_frame, text="💾 保存修改", command=self.save_current).pack(side=tk.LEFT, padx=7, ipadx=1)
        ttk.Button(btn_frame, text="📝 话术模板", command=self.open_template_editor).pack(side=tk.LEFT, padx=7, ipadx=1)
        ttk.Button(btn_frame, text="❌ 关闭", command=self.on_close).pack(side=tk.RIGHT, padx=7, ipadx=1)

        self.clear_detail_fields()

    def refresh_tree(self):
        if not self.tree:
            return
        for row in self.tree.get_children():
            self.tree.delete(row)
        for i, cat in enumerate(self.categories):
            ratio_percent = int(cat.get("ratio", 0) * 100)
            level = cat.get("level", 1)
            self.tree.insert("", tk.END, iid=str(i), values=(cat["name"], f"{ratio_percent}%", f"方案{level}"))

    def on_tree_select(self, event):
        selection = self.tree.selection()
        if not selection:
            self.clear_detail_fields()
            return
        idx = int(selection[0])
        if 0 <= idx < len(self.categories):
            cat = self.categories[idx]
            self.current_edit_idx = idx
            self.entry_vars["name"].set(cat.get("name", ""))
            self.entry_vars["ratio"].set(str(int(cat.get("ratio", 0) * 100)))
            self.entry_vars["desc"].set(cat.get("desc", ""))
            self.entry_vars["group"].set(cat.get("group", ""))
            self.entry_vars["final_increase"].set(str(cat.get("final_increase", 5)))
            self.update_comboboxes()
            current_level = cat.get("level", 1)
            self.entry_vars["level"].set(str(current_level))
            self.update_top_level_visibility()

    def update_top_level_visibility(self):
        selected_group = self.entry_vars["group"].get().strip()
        current_level = int(self.entry_vars["level"].get() or 1)

        max_level_in_group = max([c.get("level", 0) for c in self.categories if c.get("group", "") == selected_group], default=0)

        is_top_level = (current_level >= max_level_in_group)

        if is_top_level:
            self.top_level_label.pack(anchor=tk.W, pady=(6, 1))
            self.final_label.pack(anchor=tk.W, pady=(6, 1))
            self.final_entry.pack(fill=tk.X, pady=1)
        else:
            self.top_level_label.pack_forget()
            self.final_label.pack_forget()
            self.final_entry.pack_forget()

    def update_comboboxes(self):
        self.group_combo["values"] = self.groups
        self.update_levels_for_group()

    def update_levels_for_group(self):
        selected_group = self.entry_vars["group"].get().strip()
        count_in_group = sum(1 for c in self.categories if c.get("group", "") == selected_group)
        max_level = max(count_in_group, 1)

        level_values = [str(i) for i in range(1, max_level + 1)]
        self.level_combo["values"] = level_values

    def on_group_changed(self):
        self.update_levels_for_group()
        self.auto_save()
        self.update_top_level_visibility()

    def on_level_changed(self):
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.categories):
            return

        current_cat = self.categories[idx]
        selected_group = current_cat.get("group", "")
        new_level = int(self.entry_vars["level"].get())
        old_level = current_cat.get("level", 1)

        for i, cat in enumerate(self.categories):
            if i != idx and cat.get("group", "") == selected_group and cat.get("level", 0) == new_level:
                cat["level"] = old_level
                break

        self.auto_save()
        self.update_top_level_visibility()

    def clear_detail_fields(self):
        for var in self.entry_vars.values():
            var.set("")
        self.update_comboboxes()

    def auto_save(self):
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        try:
            ratio_percent = float(self.entry_vars["ratio"].get())
            if not (0 <= ratio_percent <= 100):
                return
            ratio_val = ratio_percent / 100
        except:
            return

        name = self.entry_vars["name"].get().strip()
        if not name:
            return

        self.categories[idx]["name"] = name
        self.categories[idx]["ratio"] = ratio_val
        self.categories[idx]["desc"] = self.entry_vars["desc"].get().strip()
        self.categories[idx]["group"] = self.entry_vars["group"].get().strip()
        self.categories[idx]["level"] = int(self.entry_vars["level"].get() or 1)
        self.categories[idx]["final_increase"] = float(self.entry_vars["final_increase"].get() or 5)

        self.refresh_tree()
        self.callback(self.categories)
        self.tree.selection_set(str(idx))
        self.status_label.config(text="已保存")
        self.after_id = self.after(2000, lambda: self.status_label.config(text=""))

    def save_current(self):
        self.auto_save()

    def add_category(self):
        selected_group = self.entry_vars["group"].get().strip() or (self.groups[0] if self.groups else "腐烂")

        same_group_cats = [c for c in self.categories if c.get("group", "") == selected_group]
        max_level = max([c.get("level", 0) for c in same_group_cats], default=0)
        new_level = max_level + 1

        new_cat = {
            "name": "新方案",
            "desc": "",
            "ratio": 0.2,
            "group": selected_group,
            "level": new_level,
            "final_increase": 5,
            "template": "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？",
            "template_upgrade": "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？",
        }
        self.categories.append(new_cat)
        self.refresh_tree()
        self.callback(self.categories)
        children = self.tree.get_children()
        if children:
            last_idx = len(children) - 1
            self.tree.selection_set(children[last_idx])
            self.tree.focus(children[last_idx])
            self.on_tree_select(None)

    def delete_category(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("未选择", "请先选中要删除的分类")
            return
        idx = int(selection[0])
        if messagebox.askyesno("确认删除", f"确定删除【{self.categories[idx]['name']}】吗？"):
            del self.categories[idx]
            self.refresh_tree()
            self.callback(self.categories)
            self.clear_detail_fields()

    def open_template_editor(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("未选择", "请先在左侧列表选中一个分类")
            return
        idx = int(selection[0])
        TemplateEditorDialog(self, idx, self.categories[idx], self.on_template_updated)

    def on_template_updated(self, idx, template, template_upgrade):
        self.categories[idx]["template"] = template
        self.categories[idx]["template_upgrade"] = template_upgrade
        messagebox.showinfo("成功", "话术模板已更新")

    def on_close(self):
        # 最终保存并回调
        self.callback(self.categories)
        self.destroy()


# ==================== 话术模板编辑对话框 ====================
class TemplateEditorDialog(tk.Toplevel):
    def __init__(self, parent, idx, category, callback):
        super().__init__(parent)
        self.title(f"话术模板编辑 · {category.get('name', '')}")
        self.geometry("700x550")
        self.resizable(True, True)

        self.idx = idx
        self.category = category
        self.callback = callback

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def build_ui(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="可用的变量占位符：{name}退款原因、{desc}详细描述、{ratio}赔付比例、{money}赔付金额、{final_ratio}最终赔付比例",
                  font=("微软雅黑", 9), foreground="gray").pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(main_frame, text="方案话术模板（初始方案使用）：", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        self.template_text = tk.Text(main_frame, height=6, font=("微软雅黑", 10), wrap=tk.WORD, relief=tk.SOLID, borderwidth=1)
        self.template_text.pack(fill=tk.X, pady=(0, 5))

        placeholder_frame1 = ttk.Frame(main_frame)
        placeholder_frame1.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(placeholder_frame1, text="快速插入：", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=(0, 5))
        placeholders1 = [("{name}", "退款原因"), ("{desc}", "详细描述"), ("{ratio}", "赔付比例"), ("{money}", "赔付金额")]
        for ph, label in placeholders1:
            ttk.Button(placeholder_frame1, text=label, width=8,
                      command=lambda p=ph: self.template_text.insert(tk.END, p)).pack(side=tk.LEFT, padx=2)

        ttk.Label(main_frame, text="升级话术模板（顾客不同意时使用）：", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        self.template_upgrade_text = tk.Text(main_frame, height=6, font=("微软雅黑", 10), wrap=tk.WORD, relief=tk.SOLID, borderwidth=1)
        self.template_upgrade_text.pack(fill=tk.X, pady=(0, 5))

        placeholder_frame2 = ttk.Frame(main_frame)
        placeholder_frame2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(placeholder_frame2, text="快速插入：", font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=(0, 5))
        placeholders2 = [("{name}", "退款原因"), ("{desc}", "详细描述"), ("{ratio}", "赔付比例"), ("{money}", "赔付金额"), ("{final_ratio}", "最终赔付比例")]
        for ph, label in placeholders2:
            ttk.Button(placeholder_frame2, text=label, width=10,
                      command=lambda p=ph: self.template_upgrade_text.insert(tk.END, p)).pack(side=tk.LEFT, padx=2)

        self.template_text.insert("1.0", self.category.get("template", ""))
        self.template_upgrade_text.insert("1.0", self.category.get("template_upgrade", ""))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="💾 保存", command=self.on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="恢复默认", command=self.on_reset).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ 取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=5)

    def on_save(self):
        template = self.template_text.get("1.0", tk.END).strip()
        template_upgrade = self.template_upgrade_text.get("1.0", tk.END).strip()
        self.callback(self.idx, template, template_upgrade)
        self.destroy()

    def on_reset(self):
        default_template = "亲，非常抱歉！根据您反馈的【{name}】（{desc}），我们按单根/问题部分赔付{ratio}%，为您直接补偿 {money} 元。金额立即到账，无需退货，您看可以吗？"
        default_upgrade = "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付 {money} 元。金额立即到账，无需退货，这是我们最大的诚意了，您看可以吗？"
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", default_template)
        self.template_upgrade_text.delete("1.0", tk.END)
        self.template_upgrade_text.insert("1.0", default_upgrade)

    def on_cancel(self):
        self.destroy()


# ==================== 退款大类管理对话框 ====================
class GroupManagerDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.withdraw()
        self.title("分类管理")

        self.groups = load_groups()
        self.callback = callback
        self.group_vars = {}

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 500
        window_height = 500
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.deiconify()

    def create_tooltip(self, widget, text):
        tooltip = None
        def on_enter(event):
            nonlocal tooltip
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            tooltip = tk.Toplevel(widget)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(tooltip, text=text, background="#ffffcc", relief=tk.SOLID, borderwidth=1,
                              font=("微软雅黑", 9), padding=6, wraplength=220)
            label.pack()
        def on_leave(event):
            nonlocal tooltip
            if tooltip:
                tooltip.destroy()
                tooltip = None
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def build_ui(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="已设置的分类（可新增/编辑/删除）：", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))

        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.group_listbox = tk.Listbox(list_frame, font=("微软雅黑", 11), height=8)
        self.group_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.group_listbox.yview)
        self.group_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.refresh_list()

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        self.new_group_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.new_group_var, font=("微软雅黑", 11)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(input_frame, text="新增分类", command=self.add_group).pack(side=tk.RIGHT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        btn_del_grp = ttk.Button(btn_frame, text="删除选中", command=self.delete_group)
        btn_del_grp.pack(side=tk.LEFT, padx=8, ipadx=1)
        if DEBUG_MODE:
            self.create_tooltip(btn_del_grp, "[分类窗口_删除选中]")
        ttk.Button(btn_frame, text="关闭", command=self.on_close).pack(side=tk.RIGHT, padx=8, ipadx=1)

    def refresh_list(self):
        self.group_listbox.delete(0, tk.END)
        for g in self.groups:
            self.group_listbox.insert(tk.END, g)

    def add_group(self):
        new_group = self.new_group_var.get().strip()
        if not new_group:
            messagebox.showwarning("提示", "请输入分类名称")
            return
        if new_group in self.groups:
            messagebox.showwarning("提示", "该分类已存在")
            return
        self.groups.append(new_group)
        self.groups.sort()
        self.new_group_var.set("")
        self.refresh_list()
        save_groups(self.groups)
        self.callback(self.groups)

    def delete_group(self):
        selection = self.group_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选中要删除的分类")
            return
        idx = selection[0]
        group_name = self.groups[idx]
        if messagebox.askyesno("确认删除", f"确定删除分类【{group_name}】吗？\n注意：该分类下的所有方案也会被删除！"):
            del self.groups[idx]
            self.refresh_list()
            save_groups(self.groups)
            self.callback(self.groups)

    def on_close(self):
        self.destroy()


# ==================== 主程序入口 ====================
if __name__ == "__main__":
    root = tk.Tk()
    app = RefundApp(root)
    root.mainloop()