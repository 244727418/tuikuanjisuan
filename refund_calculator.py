import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog
import json
import os
import io
import csv
import sys
from datetime import datetime

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    import cairosvg
except ImportError:
    cairosvg = None

DEBUG_MODE = True
GROUP_COLOR_PALETTE = [
    "#2f80ed",
    "#27ae60",
    "#f2994a",
    "#eb5757",
    "#9b51e0",
    "#00a8a8",
    "#d9468f",
    "#6b7280",
]

# ==================== 默认退款分类数据 ====================
DEFAULT_CATEGORIES = [{'name': '轻度腐烂',
  'desc': '表面小面积软点、水渍斑，削掉可食用',
  'ratio': 0.3,
  'scope': '单根局部软点',
  'group': '腐烂',
  'level': 1,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '中度腐烂',
  'desc': '局部腐烂面积＜1/3，轻微渗水或霉点',
  'ratio': 0.6,
  'scope': '单根腐烂小于1/3',
  'group': '腐烂',
  'level': 2,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '重度腐烂',
  'desc': '大面积腐烂、发黑发霉超过1/2',
  'ratio': 0.8,
  'scope': '单根腐烂超过一半',
  'group': '腐烂',
  'level': 3,
  'final_increase': 10.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '轻度发芽',
  'desc': '芽长＜1cm，山药体未变软',
  'ratio': 0.1,
  'scope': '芽点刚萌动',
  'group': '发芽',
  'level': 1,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '中度发芽',
  'desc': '芽长1-3cm，根体开始变软',
  'ratio': 0.3,
  'scope': '芽长1到3厘米',
  'group': '发芽',
  'level': 2,
  'final_increase': 5,
  'per_root_enabled': False,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '重度发芽',
  'desc': '芽长＞3cm，根体明显干缩',
  'ratio': 0.6,
  'scope': '芽长超过3厘米',
  'group': '发芽',
  'level': 3,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '轻度断裂',
  'desc': '断成两截，断面整齐新鲜',
  'ratio': 0.2,
  'scope': '断裂1到2根',
  'group': '断裂',
  'level': 1,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '中度断裂',
  'desc': '断成两截以上，断面轻微氧化',
  'ratio': 0.4,
  'scope': '断裂2到3根',
  'group': '断裂',
  'level': 2,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '重度断裂',
  'desc': '断成多截、断面发黑',
  'ratio': 0.6,
  'scope': '断裂3根及以上',
  'group': '断裂',
  'level': 3,
  'final_increase': 5.0,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '轻微挤压',
  'desc': '',
  'ratio': 0.2,
  'scope': '请修改范畴',
  'group': '挤压',
  'level': 1,
  'final_increase': 5,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？'},
 {'name': '外观轻度瑕疵',
  'desc': '外观轻微磕碰瑕疵',
  'ratio': 0.1,
  'group': '外观瑕疵',
  'level': 1,
  'final_increase': 5,
  'per_root_enabled': False,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'scope': '外观有轻微磕碰瑕疵'},
 {'name': '外观中度瑕疵',
  'desc': '外观中等程度瑕疵',
  'ratio': 0.2,
  'group': '外观瑕疵',
  'level': 2,
  'final_increase': 5,
  'per_root_enabled': False,
  'template': '亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'template_upgrade': '亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？',
  'scope': '外观中等程度瑕疵'}]

DEFAULT_GROUPS = [{'name': '发芽', 'per_root_enabled': False},
 {'name': '外观瑕疵', 'per_root_enabled': False},
 {'name': '挤压', 'per_root_enabled': False},
 {'name': '断裂', 'per_root_enabled': False},
 {'name': '腐烂', 'per_root_enabled': True}]

DATA_FILE = "refund_data.json"
DATA_VERSION = 1
CONFIG_FILE = "refund_categories.json"
LOG_FILE = "refund_log.json"
GROUPS_FILE = "refund_groups.json"
APP_CONFIG_FILE = "config.json"
PER_ROOT_ICON_FILE = os.path.join("assets", "per_root.png")
PER_ROOT_SVG_ICON_FILE = os.path.join("assets", "per_root.svg")
DEFAULT_TEMPLATE = "亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？"
DEFAULT_UPGRADE_TEMPLATE = "亲，非常抱歉！您说方案一（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {money} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？"
FONT_FAMILY = "微软雅黑"
APP_BG = "#f4f7fb"
PANEL_BG = "#ffffff"
TEXT_BG = "#fbfdff"
MUTED_FG = "#6b7280"
PRIMARY_FG = "#1f2937"
ACCENT = "#2563eb"
BUTTON_VARIANTS = {
    "primary": {"fill": "#2563eb", "hover": "#1d4ed8", "press": "#1e40af", "text": "#ffffff", "outline": "#1d4ed8"},
    "success": {"fill": "#16a34a", "hover": "#15803d", "press": "#166534", "text": "#ffffff", "outline": "#15803d"},
    "danger": {"fill": "#dc2626", "hover": "#b91c1c", "press": "#991b1b", "text": "#ffffff", "outline": "#b91c1c"},
    "warning": {"fill": "#f59e0b", "hover": "#d97706", "press": "#b45309", "text": "#111827", "outline": "#d97706"},
    "secondary": {"fill": "#e8f0fb", "hover": "#dbeafe", "press": "#bfdbfe", "text": "#1e3a8a", "outline": "#bfdbfe"},
    "ghost": {"fill": "#ffffff", "hover": "#f3f6fb", "press": "#e5eaf3", "text": "#374151", "outline": "#d6deea"},
}


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path):
    base_dir = getattr(sys, "_MEIPASS", app_dir())
    return os.path.join(base_dir, relative_path)


def user_data_path(filename):
    return os.path.join(app_dir(), filename)


def show_toast(parent, message, kind="info"):
    window = parent.winfo_toplevel()
    old_toast = getattr(window, "_toast_window", None)
    if old_toast and old_toast.winfo_exists():
        old_toast.destroy()

    toast = tk.Toplevel(window)
    window._toast_window = toast
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)

    colors = {
        "error": ("#fff1f0", "#cf1322"),
        "warning": ("#fffbe6", "#ad6800"),
        "success": ("#f6ffed", "#237804"),
        "info": ("#e6f4ff", "#0958d9"),
    }
    background, foreground = colors.get(kind, colors["info"])
    label = ttk.Label(
        toast,
        text=message,
        background=background,
        foreground=foreground,
        relief=tk.SOLID,
        borderwidth=1,
        padding=(12, 8),
        wraplength=260,
        font=("微软雅黑", 10),
    )
    label.pack()

    window.update_idletasks()
    toast.update_idletasks()
    width = toast.winfo_reqwidth()
    height = toast.winfo_reqheight()
    x = window.winfo_rootx() + max((window.winfo_width() - width) // 2, 0)
    y = window.winfo_rooty() + 36
    toast.geometry(f"{width}x{height}+{x}+{y}")

    def set_alpha(value):
        if toast.winfo_exists():
            try:
                toast.attributes("-alpha", value)
            except tk.TclError:
                pass

    def fade_in(step=0):
        if not toast.winfo_exists():
            return
        set_alpha(min(step / 5, 1))
        if step < 5:
            toast.after(40, lambda: fade_in(step + 1))
        else:
            toast.after(200, fade_out)

    def fade_out(step=5):
        if not toast.winfo_exists():
            return
        set_alpha(max(step / 5, 0))
        if step > 0:
            toast.after(40, lambda: fade_out(step - 1))
        else:
            toast.destroy()
            if getattr(window, "_toast_window", None) is toast:
                window._toast_window = None

    set_alpha(0)
    fade_in()


def load_per_root_icon(size=18):
    png_icon = resource_path(PER_ROOT_ICON_FILE)
    svg_icon = resource_path(PER_ROOT_SVG_ICON_FILE)

    if Image is not None and ImageTk is not None and os.path.exists(png_icon):
        try:
            image = Image.open(png_icon).convert("RGBA")
            image.thumbnail((size, size), Image.LANCZOS)
            return ImageTk.PhotoImage(image)
        except Exception:
            pass

    if Image is not None and ImageTk is not None and cairosvg is not None and os.path.exists(svg_icon):
        try:
            png_bytes = cairosvg.svg2png(url=svg_icon, output_width=size, output_height=size)
            image = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
            return ImageTk.PhotoImage(image)
        except Exception:
            pass

    return None


# ==================== 数据管理 ====================
def default_group_color(index):
    return GROUP_COLOR_PALETTE[index % len(GROUP_COLOR_PALETTE)]


def normalize_group(group, index=0):
    if isinstance(group, str):
        return {
            "name": group,
            "per_root_enabled": False,
            "color": default_group_color(index),
        }
    return {
        "name": group.get("name", ""),
        "per_root_enabled": bool(group.get("per_root_enabled", False)),
        "color": group.get("color") or default_group_color(index),
    }


def read_json_file(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return default


def default_data():
    return {
        "version": DATA_VERSION,
        "categories": [dict(c) for c in DEFAULT_CATEGORIES],
        "groups": [normalize_group(g, i) for i, g in enumerate(DEFAULT_GROUPS)],
        "logs": [],
        "app_config": {},
    }


def normalize_data(data):
    fallback = default_data()
    if not isinstance(data, dict):
        data = {}

    categories = data.get("categories")
    if not isinstance(categories, list):
        categories = fallback["categories"]

    groups = data.get("groups")
    if not isinstance(groups, list):
        groups = fallback["groups"]
    groups = [
        normalized
        for i, group in enumerate(groups)
        for normalized in [normalize_group(group, i)]
        if normalized.get("name")
    ] or fallback["groups"]

    logs = data.get("logs")
    if not isinstance(logs, list):
        logs = []

    app_config = data.get("app_config")
    if not isinstance(app_config, dict):
        app_config = {}

    return {
        "version": data.get("version", DATA_VERSION),
        "categories": categories,
        "groups": groups,
        "logs": logs,
        "app_config": app_config,
    }


def load_legacy_data():
    data = default_data()

    legacy_categories = read_json_file(user_data_path(CONFIG_FILE), None)
    if isinstance(legacy_categories, list):
        data["categories"] = legacy_categories

    legacy_groups = read_json_file(user_data_path(GROUPS_FILE), None)
    if isinstance(legacy_groups, list):
        data["groups"] = legacy_groups

    legacy_logs = read_json_file(user_data_path(LOG_FILE), None)
    if isinstance(legacy_logs, list):
        data["logs"] = legacy_logs

    legacy_config = read_json_file(user_data_path(APP_CONFIG_FILE), None)
    if isinstance(legacy_config, dict):
        data["app_config"] = legacy_config

    return normalize_data(data)


def load_data():
    writable_data = user_data_path(DATA_FILE)
    bundled_data = resource_path(DATA_FILE)
    if os.path.exists(writable_data):
        return normalize_data(read_json_file(writable_data, {}))
    if os.path.exists(bundled_data):
        return normalize_data(read_json_file(bundled_data, {}))
    return load_legacy_data()


def save_data(data):
    normalized = normalize_data(data)
    with open(user_data_path(DATA_FILE), "w", encoding="utf-8") as f:
        json.dump(normalized, f, ensure_ascii=False, indent=2)


def update_data_section(section, value):
    data = load_data()
    data[section] = value
    save_data(data)


def load_app_config():
    return load_data().get("app_config", {})


def save_app_config(config):
    update_data_section("app_config", config if isinstance(config, dict) else {})


def load_recent_colors(groups=None):
    config = load_app_config()
    colors = [
        color
        for color in config.get("recent_colors", [])
        if isinstance(color, str) and color.startswith("#")
    ]
    if not colors and groups:
        colors = [g.get("color") for g in groups if g.get("color")]
    colors.extend(GROUP_COLOR_PALETTE)
    unique_colors = []
    for color in colors:
        if color and color not in unique_colors:
            unique_colors.append(color)
    return unique_colors[:8]


def save_recent_colors(colors):
    config = load_app_config()
    config["recent_colors"] = colors[:8]
    save_app_config(config)


def load_groups():
    return load_data().get("groups", [])


def save_groups(groups):
    update_data_section("groups", groups)


def load_categories():
    return load_data().get("categories", [])


def save_categories(categories):
    update_data_section("categories", categories)


def load_log():
    return load_data().get("logs", [])


def save_log(logs):
    update_data_section("logs", logs)


def add_log_entry(reason_name, amount, ratio, level, refund_money=None, total_roots="", bad_roots="", use_per_root=False, disagree_count=1):
    logs = load_log()
    logs.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "reason": reason_name,
        "amount": amount,
        "ratio": ratio,
        "level": level,
        "refund_money": refund_money,
        "total_roots": total_roots,
        "bad_roots": bad_roots,
        "use_per_root": use_per_root,
        "disagree_count": disagree_count,
        "event": "赔付不同意"
    })
    save_log(logs)


def render_template(template, **kwargs):
    for key, value in kwargs.items():
        template = template.replace(f"{{{key}}}", str(value))
    return template


def get_group_names(groups):
    return [g.get("name", "") for g in groups if g.get("name")]


def get_group_color(groups, group_name):
    for i, group in enumerate(groups):
        if group.get("name", "") == group_name:
            return group.get("color") or default_group_color(i)
    return default_group_color(0)


def is_group_per_root_enabled(groups, group_name):
    for group in groups:
        if group.get("name", "") == group_name:
            return group.get("per_root_enabled", False)
    return False


class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, color="#2f80ed", width=118, height=36):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bd=0, bg=parent.cget("bg"))
        self.text = text
        self.command = command
        self.color = color
        self.height = height
        self.fill = "#ffffff"
        self.hover_fill = "#f5f9ff"
        self.press_fill = "#eaf3ff"
        self.current_fill = self.fill
        self.configure(cursor="hand2")
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.draw()

    def rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def draw(self):
        self.delete("all")
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), self.height)
        self.rounded_rect(2, 2, width - 2, height - 2, 10, fill=self.current_fill, outline=self.color, width=1.5)
        self.create_text(
            width // 2,
            height // 2,
            text=self.text,
            fill=self.color,
            font=(FONT_FAMILY, 11, "bold"),
        )

    def on_enter(self, event):
        self.current_fill = self.hover_fill
        self.draw()

    def on_leave(self, event):
        self.current_fill = self.fill
        self.draw()

    def on_press(self, event):
        self.current_fill = self.press_fill
        self.draw()

    def on_release(self, event):
        self.current_fill = self.hover_fill
        self.draw()
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()


# ==================== 主应用类 ====================
class RoundedActionButton(tk.Canvas):
    def __init__(self, parent, text, command, variant="secondary", width=None, height=34, radius=13):
        bg = parent.cget("bg") if "bg" in parent.keys() else APP_BG
        measured_width = max(74, len(text) * 15 + 28)
        super().__init__(
            parent,
            width=width or measured_width,
            height=height,
            highlightthickness=0,
            bd=0,
            bg=bg,
        )
        self.text = text
        self.command = command
        self.height = height
        self.radius = radius
        self.palette = BUTTON_VARIANTS.get(variant, BUTTON_VARIANTS["secondary"])
        self.current_fill = self.palette["fill"]
        self.configure(cursor="hand2")
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.draw()

    def rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def draw(self):
        self.delete("all")
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), self.height)
        self.rounded_rect(
            2,
            2,
            width - 2,
            height - 2,
            self.radius,
            fill=self.current_fill,
            outline=self.palette["outline"],
            width=1,
        )
        self.create_text(
            width // 2,
            height // 2,
            text=self.text,
            fill=self.palette["text"],
            font=(FONT_FAMILY, 10, "bold"),
        )

    def on_enter(self, event):
        self.current_fill = self.palette["hover"]
        self.draw()

    def on_leave(self, event):
        self.current_fill = self.palette["fill"]
        self.draw()

    def on_press(self, event):
        self.current_fill = self.palette["press"]
        self.draw()

    def on_release(self, event):
        self.current_fill = self.palette["hover"]
        self.draw()
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()


class RoundedColorSwatch(tk.Canvas):
    def __init__(self, parent, color, command, size=30):
        bg = parent.cget("bg") if "bg" in parent.keys() else APP_BG
        super().__init__(parent, width=size, height=size, highlightthickness=0, bd=0, bg=bg)
        self.color = color
        self.command = command
        self.size = size
        self.selected = False
        self.configure(cursor="hand2")
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<ButtonRelease-1>", self.on_release)
        self.draw()

    def draw(self):
        self.delete("all")
        width = max(self.winfo_width(), 1)
        height = max(self.winfo_height(), self.size)
        outline = "#111827" if self.selected else "#d6deea"
        outline_width = 3 if self.selected else 1
        self.create_polygon(
            8, 2,
            width - 8, 2,
            width - 2, 8,
            width - 2, height - 8,
            width - 8, height - 2,
            8, height - 2,
            2, height - 8,
            2, 8,
            smooth=True,
            fill=self.color,
            outline=outline,
            width=outline_width,
        )

    def set_selected(self, selected):
        self.selected = selected
        self.draw()

    def on_release(self, event):
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self.command()


def action_button(parent, text, command, variant="secondary", width=None, height=34):
    return RoundedActionButton(parent, text=text, command=command, variant=variant, width=width, height=height)


class RefundApp:
    def __init__(self, root):
        self.root = root
        self.root.title("生鲜山药售后 · 话术生成器")
        self.root.geometry("680x890")
        self.root.minsize(640, 800)
        self.root.configure(bg=APP_BG)
        self.configure_styles()

        self.categories = load_categories()
        self.groups = load_groups()
        self.bubble_buttons = []
        self.amount_var = tk.StringVar(value="19.90")
        self.history_text = None
        self.current_category = None
        self.current_ratio = 0
        self.current_level = 0
        self.current_root_money = 0
        self.current_total_roots = ""
        self.current_bad_roots = ""
        self.current_use_per_root = False
        self.current_disagree_count = 0
        self.settings_dialog = None
        self.group_manager_dialog = None
        self.log_viewer_dialog = None
        self.per_root_icon = load_per_root_icon()

        self.build_ui()
        self.refresh_bubbles()
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 680
        window_height = 890
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def configure_styles(self):
        self.root.option_add("*Font", f"{FONT_FAMILY} 10")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=(FONT_FAMILY, 10), background=APP_BG, foreground=PRIMARY_FG)
        style.configure("TFrame", background=APP_BG)
        style.configure("Panel.TFrame", background=APP_BG)
        style.configure("TLabel", background=APP_BG, foreground=PRIMARY_FG)
        style.configure("Muted.TLabel", background=APP_BG, foreground=MUTED_FG)
        style.configure("Header.TLabel", background=APP_BG, foreground="#111827", font=(FONT_FAMILY, 16, "bold"))
        style.configure("Section.TLabel", background=APP_BG, foreground=PRIMARY_FG, font=(FONT_FAMILY, 10, "bold"))
        style.configure("TButton", font=(FONT_FAMILY, 10), padding=(8, 5))
        style.configure("TEntry", padding=(4, 3))
        style.configure("TCombobox", padding=(4, 3))
        style.configure("TLabelFrame", background=APP_BG, bordercolor="#d8e0ea", relief=tk.SOLID)
        style.configure("TLabelFrame.Label", background=APP_BG, foreground="#111827", font=(FONT_FAMILY, 10, "bold"))
        style.configure("Treeview", font=(FONT_FAMILY, 10), rowheight=28, background=PANEL_BG, fieldbackground=PANEL_BG, foreground=PRIMARY_FG)
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 10, "bold"), background="#eaf1fb", foreground="#111827")
        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", "#111827")])

    # ---------- 主界面 ----------
    def build_ui(self):
        # 顶部说明
        top_frame = ttk.Frame(self.root, padding=(14, 12, 14, 8))
        top_frame.pack(fill=tk.X)
        ttk.Label(top_frame, text="🛠️ 山药售后问题快速话术生成", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(top_frame, text="鼠标悬停气泡查看退款范畴 · 点击气泡自动生成话术", style="Muted.TLabel").pack(anchor=tk.W, pady=(3, 0))

        # 中间：气泡容器 + 右侧设置按钮
        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 10))

        # 右侧设置面板（先放置固定宽度）
        right_frame = ttk.Frame(middle_frame, width=190, style="Panel.TFrame", padding=(10, 8))
        right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        right_frame.pack_propagate(False)

        # 左侧气泡区域（自适应内容）
        left_frame = ttk.LabelFrame(middle_frame, text="退款方案选择", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12))

        canvas = tk.Canvas(left_frame, borderwidth=0, highlightthickness=0, bg=APP_BG)
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.bubble_canvas = canvas
        self.bubble_frame = tk.Frame(canvas, bg=APP_BG)

        self.bubble_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.bubble_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        action_button(right_frame, text="⚙️ 设置管理", command=self.open_settings, variant="primary", width=170).pack(fill=tk.X, pady=4)
        action_button(right_frame, text="🏷️ 分类管理", command=self.open_group_manager, variant="secondary", width=170).pack(fill=tk.X, pady=4)
        action_button(right_frame, text="📊 查看记录日志", command=self.open_log_viewer, variant="secondary", width=170).pack(fill=tk.X, pady=4)
        ttk.Separator(right_frame).pack(fill=tk.X, pady=4)

        ttk.Label(right_frame, text="实付金额 (元)：", style="Section.TLabel").pack(anchor=tk.W)
        self.amount_entry = ttk.Entry(right_frame, textvariable=self.amount_var, font=(FONT_FAMILY, 11))
        self.amount_entry.pack(fill=tk.X, pady=(4, 2))
        self.amount_entry.bind("<FocusIn>", lambda e: self.amount_entry.select_range(0, tk.END))

        self.root_count_frame = ttk.LabelFrame(right_frame, text="按根数赔付", padding=6)
        self.root_count_frame.pack(fill=tk.X, pady=(6, 2))
        ttk.Label(self.root_count_frame, text="用户收到根数（可选）：").pack(anchor=tk.W)
        self.total_roots_var = tk.StringVar()
        self.total_roots_entry = ttk.Entry(self.root_count_frame, textvariable=self.total_roots_var, font=(FONT_FAMILY, 10))
        self.total_roots_entry.pack(fill=tk.X, pady=(1, 3))
        ttk.Label(self.root_count_frame, text="损坏根数（可选）：").pack(anchor=tk.W)
        self.bad_roots_var = tk.StringVar()
        self.bad_roots_entry = ttk.Entry(self.root_count_frame, textvariable=self.bad_roots_var, font=(FONT_FAMILY, 10))
        self.bad_roots_entry.pack(fill=tk.X, pady=(1, 0))
        ttk.Label(self.root_count_frame, text="填写完整且分类已勾选时生效", style="Muted.TLabel", wraplength=155).pack(anchor=tk.W, pady=(3, 0))

        ttk.Separator(right_frame).pack(fill=tk.X, pady=5)

        ttk.Label(right_frame, text="生成话术预览：", style="Section.TLabel").pack(anchor=tk.W)
        self.history_text = tk.Text(right_frame, height=11, width=28, font=(FONT_FAMILY, 10), wrap=tk.WORD,
                                    relief=tk.SOLID, borderwidth=1, bg=TEXT_BG, fg=PRIMARY_FG,
                                    insertbackground=PRIMARY_FG, padx=8, pady=8)
        self.history_text.pack(fill=tk.BOTH, expand=True, pady=(3, 4))
        action_button(right_frame, text="导出表格", command=self.export_refund_table, variant="success", width=170).pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self.disagree_btn = action_button(right_frame, text="顾客不同意 → 升级方案", command=self.on_disagree, variant="warning", width=170)
        self.disagree_btn.pack(fill=tk.X, pady=(2, 4))

        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W, padding=(10, 5), style="Muted.TLabel")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------- 刷新气泡按钮 ----------
    def refresh_bubbles(self):
        for btn in self.bubble_buttons:
            btn.destroy()
        self.bubble_buttons.clear()

        for widget in self.bubble_frame.winfo_children():
            widget.destroy()

        group_names = get_group_names(self.groups)
        group_names.sort(key=lambda name: 0 if is_group_per_root_enabled(self.groups, name) else 1)

        row = 0
        for group_name in group_names:
            group_cats = [c for c in self.categories if c.get("group", "") == group_name]
            group_cats.sort(key=lambda x: x.get("level", 0))
            if not group_cats:
                continue
            group_color = get_group_color(self.groups, group_name)

            title_frame = tk.Frame(self.bubble_frame, bg=group_color)
            title_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=3, pady=(6 if row else 0, 2))
            tk.Label(
                title_frame,
                text=group_name,
                font=(FONT_FAMILY, 10, "bold"),
                foreground="#111827",
                background=group_color,
                padx=8,
                pady=3,
            ).pack(side=tk.LEFT)
            if is_group_per_root_enabled(self.groups, group_name):
                per_root_label = tk.Label(
                    title_frame,
                    text="按根数",
                    image=self.per_root_icon,
                    compound=tk.LEFT if self.per_root_icon else tk.NONE,
                    font=(FONT_FAMILY, 8),
                    foreground="#111827",
                    background="#ffffff",
                    padx=6,
                    pady=1,
                )
                per_root_label.pack(side=tk.RIGHT, padx=6, pady=3)
            row += 1

            for i, cat in enumerate(group_cats):
                name = cat["name"]
                desc = cat.get("desc", "")
                scope = cat.get("scope", "").strip() or desc or "未设置"
                ratio = cat.get("ratio", 0)

                tooltip_text = f"{name}\n退款范畴：{scope}\n参考比例：{int(ratio*100)}%"

                btn = RoundedButton(
                    self.bubble_frame,
                    text=name,
                    color=group_color,
                    command=lambda c=cat: self.on_bubble_click(c),
                )
                btn.grid(row=row + (i // 3), column=i % 3, padx=3, pady=3, sticky="ew", ipady=3)
                self.bubble_buttons.append(btn)
                self.create_tooltip(btn, tooltip_text)

            row += (len(group_cats) + 2) // 3

        for col in range(3):
            self.bubble_frame.grid_columnconfigure(col, weight=1, uniform="bubble")

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
                              font=(FONT_FAMILY, 9), padding=6, wraplength=220)
            label.pack()

        def on_leave(event):
            nonlocal tooltip
            if tooltip:
                tooltip.destroy()
                tooltip = None

        widget.bind("<Enter>", on_enter, add="+")
        widget.bind("<Leave>", on_leave, add="+")

    def get_root_count_values(self):
        total_roots = self.total_roots_var.get().strip()
        bad_roots = self.bad_roots_var.get().strip()
        if not total_roots and not bad_roots:
            return None
        if not total_roots or not bad_roots:
            show_toast(self.root, "如需按根数赔付，请同时填写用户收到根数和损坏根数", "warning")
            return None
        try:
            total = int(total_roots)
            bad = int(bad_roots)
        except ValueError:
            show_toast(self.root, "根数必须是整数", "warning")
            return None
        if total <= 0:
            show_toast(self.root, "用户收到根数必须大于 0", "warning")
            return None
        if bad <= 0:
            show_toast(self.root, "损坏根数必须大于 0", "warning")
            return None
        if bad > total:
            show_toast(self.root, "损坏根数不能大于用户收到根数", "warning")
            return None
        return total, bad

    # ---------- 气泡点击事件 ----------
    def on_bubble_click(self, category):
        try:
            amount = float(self.amount_var.get())
        except ValueError:
            show_toast(self.root, "请输入有效的实付金额（数字）", "warning")
            return

        name = category["name"]
        desc = category.get("desc", "")
        ratio = category.get("ratio", 0)
        scope = category.get("scope", "")
        group = category.get("group", "")
        level = category.get("level", 1)
        per_root_enabled = is_group_per_root_enabled(self.groups, group)

        self.current_category = category
        self.current_ratio = ratio
        self.current_level = level
        self.current_amount = amount

        refund_money = round(amount * ratio, 2)
        root_money = refund_money
        total = 0
        bad = 0
        use_per_root = False

        if per_root_enabled:
            root_counts = self.get_root_count_values()
            if root_counts:
                total, bad = root_counts
                root_money = round(amount * ratio * (bad / total), 2)
                use_per_root = True
            elif self.total_roots_var.get().strip() or self.bad_roots_var.get().strip():
                return

        self.current_root_money = root_money
        self.current_total_roots = total if use_per_root else ""
        self.current_bad_roots = bad if use_per_root else ""
        self.current_use_per_root = use_per_root
        self.current_disagree_count = 0

        template = category.get("template") or DEFAULT_TEMPLATE

        if use_per_root:
            template_per_root = category.get("template_per_root", "")
            if template_per_root:
                template = template_per_root

        if use_per_root:
            script = render_template(template, name=name, desc=desc, ratio=int(ratio*100),
                                    money=root_money, root_money=root_money,
                                    total_roots=total, bad_roots=bad, final_ratio=int(ratio*100))
        else:
            script = render_template(template, name=name, desc=desc, ratio=int(ratio*100),
                                    money=refund_money, final_ratio=int(ratio*100))

        self.history_text.delete("1.0", tk.END)
        self.history_text.insert(tk.END, script)
        self.root.clipboard_clear()
        self.root.clipboard_append(script)
        self.status_var.set(f"已生成【{name}】话术，赔付 {root_money} 元（已自动复制）")

    def on_disagree(self):
        if not self.current_category:
            show_toast(self.root, "请先点击气泡生成话术", "info")
            return

        name = self.current_category["name"]
        group = self.current_category.get("group", "")
        current_level = self.current_level
        final_increase = self.current_category.get("final_increase", 5)
        amount = self.current_amount
        self.current_disagree_count += 1
        add_log_entry(
            name,
            amount,
            self.current_ratio,
            current_level,
            refund_money=self.current_root_money,
            total_roots=self.current_total_roots,
            bad_roots=self.current_bad_roots,
            use_per_root=self.current_use_per_root,
            disagree_count=self.current_disagree_count,
        )

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

        refund_money = round(amount * final_ratio, 2)
        root_money = refund_money

        current_group = self.current_category.get("group", "")
        per_root_enabled = is_group_per_root_enabled(self.groups, current_group)

        total = 0
        bad = 0
        use_per_root = False
        if per_root_enabled:
            root_counts = self.get_root_count_values()
            if root_counts:
                total, bad = root_counts
                root_money = round(amount * final_ratio * (bad / total), 2)
                use_per_root = True
            elif self.total_roots_var.get().strip() or self.bad_roots_var.get().strip():
                return

        if is_final:
            template_upgrade = self.current_category.get("template_upgrade") or DEFAULT_UPGRADE_TEMPLATE
            if use_per_root:
                script = render_template(template_upgrade, name=name, desc=self.current_category.get("desc", ""),
                                         ratio=int(self.current_ratio*100), money=root_money, root_money=root_money,
                                         total_roots=total, bad_roots=bad, final_ratio=int(final_ratio*100))
            else:
                script = render_template(template_upgrade, name=name, desc=self.current_category.get("desc", ""),
                                         ratio=int(self.current_ratio*100), money=refund_money, final_ratio=int(final_ratio*100))
        else:
            template = next_cat.get("template") or DEFAULT_TEMPLATE
            desc = next_cat.get("desc", "")
            if use_per_root:
                script = render_template(template, name=next_cat["name"], desc=desc, ratio=int(final_ratio*100),
                                        money=root_money, root_money=root_money,
                                        total_roots=total, bad_roots=bad, final_ratio=int(final_ratio*100))
            else:
                script = render_template(template, name=next_cat["name"], desc=desc, ratio=int(final_ratio*100),
                                        money=refund_money, final_ratio=int(final_ratio*100))
            self.current_category = next_cat

        self.current_ratio = final_ratio
        self.current_root_money = root_money
        self.current_total_roots = total if use_per_root else ""
        self.current_bad_roots = bad if use_per_root else ""
        self.current_use_per_root = use_per_root
        self.current_level = current_level + 1 if not is_final else 99

        self.history_text.delete("1.0", tk.END)
        self.history_text.insert(tk.END, script)
        self.root.clipboard_clear()
        self.root.clipboard_append(script)
        display_name = self.current_category.get("name", name)
        self.status_var.set(f"已升级方案，【{display_name}】赔付 {root_money} 元（已自动复制）")

    # ---------- 设置窗口 ----------
    def focus_existing_window(self, dialog):
        if not dialog or not dialog.winfo_exists():
            return False
        try:
            if dialog.state() == "withdrawn":
                dialog.deiconify()
            else:
                dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
        except tk.TclError:
            return False
        return True

    def open_settings(self):
        selected_name = self.current_category.get("name", "") if self.current_category else ""
        if self.focus_existing_window(self.settings_dialog):
            if selected_name and hasattr(self.settings_dialog, "select_category_by_name"):
                self.settings_dialog.select_category_by_name(selected_name)
            return
        self.settings_dialog = SettingsDialog(
            self.root,
            self.categories,
            self.groups,
            self.on_categories_updated,
            selected_name,
            on_close_callback=lambda: setattr(self, "settings_dialog", None),
        )

    def open_group_manager(self):
        if self.focus_existing_window(self.group_manager_dialog):
            return
        self.group_manager_dialog = GroupManagerDialog(
            self.root,
            self.categories,
            self.on_groups_updated,
            on_close_callback=lambda: setattr(self, "group_manager_dialog", None),
        )

    def open_log_viewer(self):
        if self.focus_existing_window(self.log_viewer_dialog):
            self.log_viewer_dialog.reload_logs()
            return
        self.log_viewer_dialog = LogViewerDialog(
            self.root,
            on_close_callback=lambda: setattr(self, "log_viewer_dialog", None),
        )

    def export_refund_table(self):
        if not self.categories:
            show_toast(self.root, "暂无退款方案可导出", "info")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"退款方案表_{timestamp}.csv"
        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出退款方案表格",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV 表格", "*.csv"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        group_names = get_group_names(self.groups)
        group_order = {name: i for i, name in enumerate(group_names)}
        sorted_categories = sorted(
            self.categories,
            key=lambda cat: (
                group_order.get(cat.get("group", ""), len(group_order)),
                cat.get("level", 0),
                cat.get("name", ""),
            ),
        )

        headers = [
            "退款类型",
            "方案",
            "赔付比例",
            "赔付方式",
            "方案等级",
            "退款范围",
            "详细描述",
            "最终方案增加百分比",
            "最终赔付比例",
            "初始话术",
            "升级话术",
            "按根数话术",
        ]

        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for cat in sorted_categories:
                    group = cat.get("group", "")
                    ratio = cat.get("ratio", 0)
                    try:
                        ratio_percent = int(round(float(ratio) * 100))
                        ratio_text = f"{ratio_percent}%"
                    except (TypeError, ValueError):
                        ratio_text = ""
                    final_increase = cat.get("final_increase", 5)
                    try:
                        final_ratio = min(float(ratio) * 100 + float(final_increase), 100)
                        final_ratio_text = f"{int(round(final_ratio))}%"
                    except (TypeError, ValueError):
                        final_ratio_text = ""
                    writer.writerow([
                        group,
                        cat.get("name", ""),
                        ratio_text,
                        "按根数赔付" if is_group_per_root_enabled(self.groups, group) else "按订单金额比例赔付",
                        cat.get("level", ""),
                        cat.get("scope", ""),
                        cat.get("desc", ""),
                        f"{final_increase}%",
                        final_ratio_text,
                        cat.get("template") or DEFAULT_TEMPLATE,
                        cat.get("template_upgrade") or DEFAULT_UPGRADE_TEMPLATE,
                        cat.get("template_per_root", ""),
                    ])
        except Exception as exc:
            messagebox.showerror("导出失败", f"导出表格失败：{exc}", parent=self.root)
            return

        self.status_var.set(f"已导出退款方案表格：{file_path}")
        show_toast(self.root, "退款方案表格已导出", "success")

    def on_groups_updated(self, new_groups, new_categories=None):
        self.groups = new_groups
        save_groups(self.groups)
        if new_categories is not None:
            self.categories = new_categories
            save_categories(self.categories)
        self.refresh_bubbles()
        self.status_var.set("退款分类已更新")

    def on_categories_updated(self, new_categories):
        self.categories = new_categories
        save_categories(self.categories)
        self.refresh_bubbles()
        self.status_var.set("退款分类已更新")


# ==================== 设置对话框 ====================
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, categories, groups, callback, selected_name="", on_close_callback=None):
        super().__init__(parent)
        self.title("退款方案管理 · 编辑/添加/删除")
        self.geometry("1100x750")
        self.minsize(900, 650)
        self.configure(bg=APP_BG)
        self.selected_name = selected_name
        self.on_close_callback = on_close_callback

        self.categories = [dict(c) for c in categories]
        self.groups = list(groups)
        self.callback = callback
        self.tree = None
        self.detail_frame = None
        self.entry_vars = {}
        self.template_editor_dialog = None
        self.detail_entries = []
        self.per_root_icon = load_per_root_icon()

        self.withdraw()
        self.fix_duplicate_levels()
        self.build_tree()
        self.build_detail_panel()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Button-1>", self.on_detail_blank_click, add="+")

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 1100
        window_height = 750
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.deiconify()

        if self.selected_name:
            self.select_category_by_name(self.selected_name)

    def fix_duplicate_levels(self):
        KEYWORD_ORDER = {"轻": 1, "中": 2, "重": 3, "轻": 1, "中": 2, "重": 3}

        for group_name in get_group_names(self.groups):
            cats_in_group = [c for c in self.categories if c.get("group", "") == group_name]
            cats_in_group.sort(key=lambda c: KEYWORD_ORDER.get(c.get("name", "")[0], 99) if c.get("name", "") else 99)

            for idx, cat in enumerate(cats_in_group):
                cat["level"] = idx + 1

        if hasattr(self, "callback"):
            self.callback(self.categories)

    def build_tree(self):
        frame = ttk.LabelFrame(self, text="现有分类列表", padding=8)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)

        style = ttk.Style(self)
        style.configure("CategoryManager.Treeview", font=(FONT_FAMILY, 11), rowheight=30)
        style.configure("CategoryManager.Treeview.Heading", font=(FONT_FAMILY, 11, "bold"))

        columns = ("group", "pay_type", "name", "ratio", "level")
        self.tree = ttk.Treeview(frame, columns=columns, show="tree headings", selectmode="browse", style="CategoryManager.Treeview")
        self.tree.heading("#0", text="")
        self.tree.column("#0", width=30, minwidth=30, stretch=False, anchor=tk.CENTER)
        self.tree.heading("group", text="分类")
        self.tree.heading("pay_type", text="赔付方式")
        self.tree.heading("name", text="退款原因")
        self.tree.heading("ratio", text="赔付比例(%)")
        self.tree.heading("level", text="等级")
        self.tree.column("group", width=110, anchor=tk.CENTER)
        self.tree.column("pay_type", width=90, anchor=tk.CENTER)
        self.tree.column("name", width=160, anchor=tk.CENTER)
        self.tree.column("ratio", width=95, anchor=tk.CENTER)
        self.tree.column("level", width=70, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.refresh_tree()

    def build_detail_panel(self):
        panel = ttk.LabelFrame(self, text="编辑选中分类详情", padding=12)
        panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=8, pady=8, expand=False)

        self.entry_vars["name"] = tk.StringVar()
        self.entry_vars["ratio"] = tk.StringVar()
        self.entry_vars["scope"] = tk.StringVar()
        self.entry_vars["desc"] = tk.StringVar()
        self.entry_vars["group"] = tk.StringVar()
        self.entry_vars["level"] = tk.StringVar()
        self.entry_vars["final_increase"] = tk.StringVar()

        self.status_label = ttk.Label(panel, text="", foreground="green", font=(FONT_FAMILY, 9))

        ttk.Label(panel, text="退款原因名称：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        name_entry = ttk.Entry(panel, textvariable=self.entry_vars["name"], font=(FONT_FAMILY, 10), width=32)
        name_entry.pack(fill=tk.X, pady=1)
        self.register_detail_entry(name_entry)

        ttk.Label(panel, text="赔付比例 (百分比数字，如30)：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        ratio_entry = ttk.Entry(panel, textvariable=self.entry_vars["ratio"], font=(FONT_FAMILY, 10), width=32)
        ratio_entry.pack(fill=tk.X, pady=1)
        self.register_detail_entry(ratio_entry)

        ttk.Label(panel, text="所属分类：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        self.group_combo = ttk.Combobox(panel, textvariable=self.entry_vars["group"], font=(FONT_FAMILY, 10), width=30, state="readonly")
        self.group_combo.pack(fill=tk.X, pady=1)
        self.group_combo.bind("<<ComboboxSelected>>", lambda e: self.on_group_changed())

        ttk.Label(panel, text="方案等级：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        self.level_combo = ttk.Combobox(panel, textvariable=self.entry_vars["level"], font=(FONT_FAMILY, 10), width=30, state="readonly")
        self.level_combo.pack(fill=tk.X, pady=1)
        self.level_combo.bind("<<ComboboxSelected>>", lambda e: self.on_level_changed())

        self.top_level_label = ttk.Label(panel, text="🏆 最高赔偿方案", foreground="red", font=(FONT_FAMILY, 10, "bold"))

        self.final_label = ttk.Label(panel, text="最终方案增加百分比：", style="Section.TLabel")
        self.final_label.pack(anchor=tk.W, pady=(6, 1))
        self.final_entry = ttk.Entry(panel, textvariable=self.entry_vars["final_increase"], font=(FONT_FAMILY, 10), width=32)
        self.final_entry.pack(fill=tk.X, pady=1)
        self.register_detail_entry(self.final_entry)

        ttk.Label(panel, text="退款范畴（悬停显示）：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        scope_entry = ttk.Entry(panel, textvariable=self.entry_vars["scope"], font=(FONT_FAMILY, 10), width=32)
        scope_entry.pack(fill=tk.X, pady=1)
        self.register_detail_entry(scope_entry)

        ttk.Label(panel, text="详细描述（话术显示）：", style="Section.TLabel").pack(anchor=tk.W, pady=(6, 1))
        desc_entry = ttk.Entry(panel, textvariable=self.entry_vars["desc"], font=(FONT_FAMILY, 10), width=32)
        desc_entry.pack(fill=tk.X, pady=1)
        self.register_detail_entry(desc_entry)

        self.status_label.pack(anchor=tk.W, pady=(10, 0))

        btn_frame = ttk.Frame(panel)
        btn_frame.pack(fill=tk.X, pady=12)
        action_button(btn_frame, text="💾 保存修改", command=self.save_current, variant="primary").pack(side=tk.LEFT, padx=7)
        action_button(btn_frame, text="📝 话术模板", command=self.open_template_editor, variant="secondary").pack(side=tk.LEFT, padx=7)
        action_button(btn_frame, text="❌ 关闭", command=self.on_close, variant="ghost").pack(side=tk.RIGHT, padx=7)

        manage_btn_frame = ttk.Frame(panel)
        manage_btn_frame.pack(fill=tk.X, pady=(0, 12))
        action_button(manage_btn_frame, text="➕ 新增方案", command=self.add_category, variant="success").pack(side=tk.LEFT, padx=7)
        action_button(manage_btn_frame, text="🗑️ 删除选中", command=self.delete_category, variant="danger").pack(side=tk.LEFT, padx=7)

        self.clear_detail_fields()
        self.bind_detail_blank_focus(panel)

    def register_detail_entry(self, entry):
        self.detail_entries.append(entry)
        entry.bind("<FocusOut>", lambda e: self.auto_save())
        entry.bind("<Return>", lambda e: self.auto_save())

    def on_detail_blank_click(self, event):
        widget = event.widget
        if widget in self.detail_entries:
            return
        widget_class = widget.winfo_class()
        if widget_class in {"TCombobox", "Treeview", "TButton", "Button", "Canvas"}:
            return
        if self.focus_get() in self.detail_entries:
            self.focus_set()

    def bind_detail_blank_focus(self, widget):
        widget_class = widget.winfo_class()
        if widget not in self.detail_entries and widget_class not in {"TCombobox", "Treeview", "TButton", "Button", "Canvas"}:
            widget.bind("<Button-1>", self.on_detail_blank_click, add="+")
        for child in widget.winfo_children():
            self.bind_detail_blank_focus(child)

    def refresh_tree(self):
        if not self.tree:
            return
        for row in self.tree.get_children():
            self.tree.delete(row)

        group_names = get_group_names(self.groups)
        group_names.sort(key=lambda name: 0 if is_group_per_root_enabled(self.groups, name) else 1)
        group_order = {name: i for i, name in enumerate(group_names)}
        sorted_items = sorted(
            enumerate(self.categories),
            key=lambda item: (
                group_order.get(item[1].get("group", ""), len(group_order)),
                item[1].get("level", 0),
                item[1].get("name", ""),
                item[0],
            ),
        )

        for i, cat in sorted_items:
            group_name = cat.get("group", "")
            is_per_root = is_group_per_root_enabled(self.groups, group_name)
            pay_type = "按根数" if is_per_root else ""
            ratio_percent = int(cat.get("ratio", 0) * 100)
            level = cat.get("level", 1)
            group_color = get_group_color(self.groups, group_name)
            tag_name = f"category_group_{i}"
            self.tree.tag_configure(tag_name, foreground=group_color)
            self.tree.insert(
                "",
                tk.END,
                iid=str(i),
                text="●" if is_per_root and not self.per_root_icon else "",
                image=self.per_root_icon if is_per_root and self.per_root_icon else "",
                values=(group_name, pay_type, cat["name"], f"{ratio_percent}%", f"方案{level}"),
                tags=(tag_name,),
            )

    def select_category_by_name(self, name):
        for i, cat in enumerate(self.categories):
            if cat.get("name", "") == name:
                self.tree.selection_set(str(i))
                self.tree.focus(str(i))
                self.on_tree_select(None)
                break

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
            self.entry_vars["scope"].set(cat.get("scope", ""))
            self.entry_vars["desc"].set(cat.get("desc", ""))
            self.entry_vars["group"].set(cat.get("group", ""))
            self.entry_vars["final_increase"].set(str(int(cat.get("final_increase", 5))))
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
        self.group_combo["values"] = get_group_names(self.groups)
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
        self.categories[idx]["scope"] = self.entry_vars["scope"].get().strip()
        self.categories[idx]["desc"] = self.entry_vars["desc"].get().strip()
        self.categories[idx]["group"] = self.entry_vars["group"].get().strip()
        self.categories[idx]["level"] = int(self.entry_vars["level"].get() or 1)
        self.categories[idx]["final_increase"] = int(self.entry_vars["final_increase"].get() or 5)

        self.refresh_tree()
        self.callback(self.categories)
        self.tree.selection_set(str(idx))
        self.status_label.config(text="已保存")
        self.after_id = self.after(2000, lambda: self.status_label.config(text=""))

    def save_current(self):
        self.auto_save()

    def add_category(self):
        name = self.entry_vars["name"].get().strip()
        selected_group = self.entry_vars["group"].get().strip()
        level_text = self.entry_vars["level"].get().strip()

        if not name or not selected_group or not level_text:
            show_toast(self, "请先填写退款原因、所属分类和方案等级", "warning")
            return

        if any(c.get("name", "").strip() == name for c in self.categories):
            show_toast(self, "退款方案已存在", "warning")
            return

        try:
            ratio_percent = float(self.entry_vars["ratio"].get().strip())
            if not (0 <= ratio_percent <= 100):
                show_toast(self, "请输入有效的赔付比例", "warning")
                return
            ratio_val = ratio_percent / 100
        except ValueError:
            show_toast(self, "请输入有效的赔付比例", "warning")
            return

        try:
            new_level = int(level_text)
        except ValueError:
            show_toast(self, "请先填写退款原因、所属分类和方案等级", "warning")
            return

        try:
            final_increase = int(self.entry_vars["final_increase"].get().strip() or 5)
        except ValueError:
            show_toast(self, "请输入有效的最终方案增加百分比", "warning")
            return

        new_cat = {
            "name": name,
            "desc": self.entry_vars["desc"].get().strip(),
            "ratio": ratio_val,
            "scope": self.entry_vars["scope"].get().strip(),
            "group": selected_group,
            "level": new_level,
            "final_increase": final_increase,
            "per_root_enabled": False,
            "template": DEFAULT_TEMPLATE,
            "template_upgrade": DEFAULT_UPGRADE_TEMPLATE,
        }
        self.categories.append(new_cat)
        self.refresh_tree()
        self.callback(self.categories)
        new_idx = len(self.categories) - 1
        self.tree.selection_set(str(new_idx))
        self.tree.focus(str(new_idx))
        self.on_tree_select(None)
        show_toast(self, "已新增退款方案", "success")

    def delete_category(self):
        selection = self.tree.selection()
        if not selection:
            show_toast(self, "请先选中要删除的分类", "warning")
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
            show_toast(self, "请先在左侧列表选中一个分类", "warning")
            return
        if self.template_editor_dialog and self.template_editor_dialog.winfo_exists():
            self.template_editor_dialog.deiconify()
            self.template_editor_dialog.lift()
            self.template_editor_dialog.focus_force()
            return
        idx = int(selection[0])
        self.template_editor_dialog = TemplateEditorDialog(
            self,
            idx,
            self.categories[idx],
            self.groups,
            self.on_template_updated,
            on_close_callback=lambda: setattr(self, "template_editor_dialog", None),
        )

    def on_template_updated(self, idx, template, template_upgrade, template_per_root=""):
        self.categories[idx]["template"] = template
        self.categories[idx]["template_upgrade"] = template_upgrade
        self.categories[idx]["template_per_root"] = template_per_root
        show_toast(self, "话术模板已更新", "success")

    def on_close(self):
        # 最终保存并回调
        self.callback(self.categories)
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()


# ==================== 话术模板编辑对话框 ====================
class TemplateEditorDialog(tk.Toplevel):
    def __init__(self, parent, idx, category, groups, callback, on_close_callback=None):
        super().__init__(parent)
        self.title(f"话术模板编辑 · {category.get('name', '')}")
        self.geometry("860x760")
        self.minsize(820, 700)
        self.configure(bg=APP_BG)
        self.resizable(True, True)

        self.idx = idx
        self.category = category
        self.groups = groups
        self.per_root_enabled = is_group_per_root_enabled(groups, category.get("group", ""))
        self.callback = callback
        self.on_close_callback = on_close_callback

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def build_ui(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="可用的变量占位符：{name}退款原因、{desc}详细描述、{ratio}赔付比例、{money}赔付金额、{final_ratio}最终赔付比例",
                  style="Muted.TLabel").pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(main_frame, text="方案话术模板（初始方案使用）：", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 5))
        self.template_text = tk.Text(main_frame, height=6, font=(FONT_FAMILY, 10), wrap=tk.WORD, relief=tk.SOLID, borderwidth=1,
                                     bg=TEXT_BG, fg=PRIMARY_FG, insertbackground=PRIMARY_FG, padx=8, pady=8)
        self.template_text.pack(fill=tk.X, pady=(0, 5))

        placeholder_frame1 = ttk.Frame(main_frame)
        placeholder_frame1.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(placeholder_frame1, text="快速插入：", style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        placeholders1 = [("{name}", "退款原因"), ("{desc}", "详细描述"), ("{ratio}", "赔付比例"), ("{money}", "赔付金额")]
        for ph, label in placeholders1:
            action_button(placeholder_frame1, text=label, width=78, height=30,
                          command=lambda p=ph: self.template_text.insert(tk.END, p)).pack(side=tk.LEFT, padx=2)

        ttk.Label(main_frame, text="升级话术模板（顾客不同意时使用）：", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 5))
        self.template_upgrade_text = tk.Text(main_frame, height=6, font=(FONT_FAMILY, 10), wrap=tk.WORD, relief=tk.SOLID, borderwidth=1,
                                             bg=TEXT_BG, fg=PRIMARY_FG, insertbackground=PRIMARY_FG, padx=8, pady=8)
        self.template_upgrade_text.pack(fill=tk.X, pady=(0, 5))

        placeholder_frame2 = ttk.Frame(main_frame)
        placeholder_frame2.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(placeholder_frame2, text="快速插入：", style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        placeholders2 = [("{name}", "退款原因"), ("{desc}", "详细描述"), ("{ratio}", "赔付比例"), ("{money}", "赔付金额"), ("{final_ratio}", "最终赔付比例")]
        for ph, label in placeholders2:
            action_button(placeholder_frame2, text=label, width=88, height=30,
                          command=lambda p=ph: self.template_upgrade_text.insert(tk.END, p)).pack(side=tk.LEFT, padx=2)

        self.per_root_frame = ttk.LabelFrame(main_frame, text="按根数赔付话术模板", padding=5)
        self.per_root_frame.pack(fill=tk.X, pady=(10, 5))

        ttk.Label(self.per_root_frame, text="可用的变量占位符：{money}普通赔付金额、{root_money}按根数赔付金额、{total_roots}收到根数、{bad_roots}有问题根数",
                  style="Muted.TLabel").pack(anchor=tk.W, pady=(0, 5))

        self.template_per_root_text = tk.Text(self.per_root_frame, height=4, font=(FONT_FAMILY, 10), wrap=tk.WORD, relief=tk.SOLID, borderwidth=1,
                                             bg=TEXT_BG, fg=PRIMARY_FG, insertbackground=PRIMARY_FG, padx=8, pady=8)
        self.template_per_root_text.pack(fill=tk.X, pady=(0, 5))

        placeholder_frame3 = ttk.Frame(self.per_root_frame)
        placeholder_frame3.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(placeholder_frame3, text="快速插入：", style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        placeholders3 = [("{money}", "普通金额"), ("{root_money}", "按根数金额"), ("{total_roots}", "收到根数"), ("{bad_roots}", "有问题根数")]
        for ph, label in placeholders3:
            action_button(placeholder_frame3, text=label, width=88, height=30,
                          command=lambda p=ph: self.template_per_root_text.insert(tk.END, p)).pack(side=tk.LEFT, padx=2)

        if not self.category.get("template"):
            default_template, _ = self.generate_default_templates()
            self.template_text.insert("1.0", default_template)
        else:
            self.template_text.insert("1.0", self.category.get("template"))

        if not self.category.get("template_upgrade"):
            _, default_upgrade = self.generate_default_templates()
            self.template_upgrade_text.insert("1.0", default_upgrade)
        else:
            self.template_upgrade_text.insert("1.0", self.category.get("template_upgrade"))

        if self.per_root_enabled:
            self.template_per_root_text.insert("1.0", self.category.get("template_per_root") or "")
        else:
            self.per_root_frame.pack_forget()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        action_button(btn_frame, text="💾 保存", command=self.on_save, variant="primary").pack(side=tk.LEFT, padx=5)
        action_button(btn_frame, text="恢复默认", command=self.on_reset, variant="secondary").pack(side=tk.LEFT, padx=5)
        action_button(btn_frame, text="❌ 取消", command=self.on_cancel, variant="ghost").pack(side=tk.RIGHT, padx=5)

    def generate_default_templates(self):
        name = self.category.get("name", "")
        desc = self.category.get("desc", "")
        ratio = int(self.category.get("ratio", 0) * 100)
        level = self.category.get("level", 1)
        final_increase = self.category.get("final_increase", 5)
        final_ratio = min(ratio + final_increase, 100)

        default_template = f"亲，非常抱歉！根据您反馈的【{name}】（{desc}），这边帮您申请赔偿 {{money}} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？"
        default_upgrade = f"亲，非常抱歉！您说方案{level}（赔付{ratio}%）无法接受，我们非常重视您的反馈，决定升级补偿方案，按{final_ratio}%赔付。这边帮您申请赔偿 {{money}} 元，不用自己申请哈。客服帮您申请是秒到账的，您看可以吗？"

        return default_template, default_upgrade

    def on_save(self):
        template = self.template_text.get("1.0", tk.END).strip()
        template_upgrade = self.template_upgrade_text.get("1.0", tk.END).strip()
        template_per_root = ""
        if self.per_root_enabled:
            template_per_root = self.template_per_root_text.get("1.0", tk.END).strip()
        self.callback(self.idx, template, template_upgrade, template_per_root)
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()

    def on_reset(self):
        default_template, default_upgrade = self.generate_default_templates()
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", default_template)
        self.template_upgrade_text.delete("1.0", tk.END)
        self.template_upgrade_text.insert("1.0", default_upgrade)

    def on_cancel(self):
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()


# ==================== 退款大类管理对话框 ====================
class GroupManagerDialog(tk.Toplevel):
    def __init__(self, parent, categories, callback, on_close_callback=None):
        super().__init__(parent)
        self.withdraw()
        self.title("分类管理")
        self.configure(bg=APP_BG)

        self.groups = load_groups()
        self.categories = [dict(c) for c in categories]
        self.callback = callback
        self.on_close_callback = on_close_callback
        self.group_vars = {}
        self.name_edit_entry = None
        self.recent_colors = load_recent_colors(self.groups)

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 620
        window_height = 600
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
                              font=(FONT_FAMILY, 9), padding=6, wraplength=220)
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

        ttk.Label(main_frame, text="已设置的分类（可新增/编辑/删除）：", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 10))

        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        columns = ("name", "per_root", "color")
        self.group_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse", height=10)
        self.group_tree.heading("name", text="分类名称")
        self.group_tree.heading("per_root", text="是否按根数赔付")
        self.group_tree.heading("color", text="标签颜色")
        self.group_tree.column("name", width=260, anchor=tk.CENTER)
        self.group_tree.column("per_root", width=130, anchor=tk.CENTER)
        self.group_tree.column("color", width=110, anchor=tk.CENTER)
        self.group_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.group_tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.group_tree.bind("<Double-1>", self.on_tree_double_click)
        self.group_tree.bind("<<TreeviewSelect>>", self.on_group_select)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.group_tree.yview)
        self.group_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.refresh_list()

        self.color_frame = ttk.LabelFrame(main_frame, text="最近使用颜色", padding=6)
        self.color_frame.pack(fill=tk.X, pady=(0, 10))
        self.color_buttons = []
        self.render_recent_color_buttons()

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        self.new_group_var = tk.StringVar()
        input_frame.columnconfigure(0, weight=1)
        ttk.Entry(input_frame, textvariable=self.new_group_var, font=(FONT_FAMILY, 11)).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        action_button(input_frame, text="新增分类", command=self.add_group, variant="success", width=88).grid(row=0, column=1, sticky="e")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        btn_del_grp = action_button(btn_frame, text="删除选中", command=self.delete_group, variant="danger")
        btn_del_grp.pack(side=tk.LEFT, padx=8)
        if DEBUG_MODE:
            self.create_tooltip(btn_del_grp, "[分类窗口_删除选中]")
        action_button(btn_frame, text="关闭", command=self.on_close, variant="ghost").pack(side=tk.RIGHT, padx=8)

    def render_recent_color_buttons(self):
        for widget in self.color_frame.winfo_children():
            widget.destroy()
        self.color_buttons = []
        for color in self.recent_colors:
            swatch = RoundedColorSwatch(
                self.color_frame,
                color=color,
                command=lambda c=color: self.set_selected_group_color(c),
            )
            swatch.pack(side=tk.LEFT, padx=3)
            self.color_buttons.append(swatch)
        self.update_color_buttons()

    def on_tree_click(self, event):
        region = self.group_tree.identify("region", event.x, event.y)
        column = self.group_tree.identify_column(event.x)
        row_id = self.group_tree.identify_row(event.y)
        if region != "cell" or column != "#2" or not row_id:
            return
        idx = int(row_id)
        self.groups[idx]["per_root_enabled"] = not self.groups[idx].get("per_root_enabled", False)
        self.refresh_list()
        self.group_tree.selection_set(row_id)
        self.group_tree.focus(row_id)
        save_groups(self.groups)
        self.callback(self.groups, self.categories)

    def on_group_select(self, event=None):
        self.update_color_buttons()

    def update_color_buttons(self):
        selection = self.group_tree.selection()
        selected_color = ""
        if selection:
            idx = int(selection[0])
            if 0 <= idx < len(self.groups):
                selected_color = self.groups[idx].get("color", "")
        for button in getattr(self, "color_buttons", []):
            button.set_selected(button.color == selected_color)

    def set_selected_group_color(self, color):
        selection = self.group_tree.selection()
        if not selection:
            show_toast(self, "请先选中要设置颜色的分类", "warning")
            return
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.groups):
            return
        self.groups[idx]["color"] = color
        self.recent_colors = [color] + [c for c in self.recent_colors if c != color]
        self.recent_colors = self.recent_colors[:8]
        save_recent_colors(self.recent_colors)
        save_groups(self.groups)
        self.refresh_list()
        self.group_tree.selection_set(str(idx))
        self.group_tree.focus(str(idx))
        self.render_recent_color_buttons()
        self.update_color_buttons()
        self.callback(self.groups, self.categories)

    def open_group_color_picker(self, row_id):
        idx = int(row_id)
        if idx < 0 or idx >= len(self.groups):
            return
        current_color = self.groups[idx].get("color") or default_group_color(idx)
        _, color = colorchooser.askcolor(
            color=current_color,
            title=f"选择【{self.groups[idx]['name']}】标签颜色",
            parent=self,
        )
        if not color:
            return
        self.group_tree.selection_set(row_id)
        self.group_tree.focus(row_id)
        self.set_selected_group_color(color)

    def on_tree_double_click(self, event):
        region = self.group_tree.identify("region", event.x, event.y)
        column = self.group_tree.identify_column(event.x)
        row_id = self.group_tree.identify_row(event.y)
        if region != "cell" or not row_id:
            return
        if column == "#1":
            self.start_name_edit(row_id)
        elif column == "#3":
            self.open_group_color_picker(row_id)

    def start_name_edit(self, row_id):
        self.close_name_editor()
        bbox = self.group_tree.bbox(row_id, "#1")
        if not bbox:
            return
        x, y, width, height = bbox
        current_name = self.groups[int(row_id)]["name"]
        self.name_edit_var = tk.StringVar(value=current_name)
        self.name_edit_entry = ttk.Entry(self.group_tree, textvariable=self.name_edit_var, font=(FONT_FAMILY, 11))
        self.name_edit_entry.place(x=x, y=y, width=width, height=height)
        self.name_edit_entry.focus_set()
        self.name_edit_entry.select_range(0, tk.END)
        self.name_edit_entry.bind("<Return>", lambda e, rid=row_id: self.save_name_edit(rid))
        self.name_edit_entry.bind("<FocusOut>", lambda e, rid=row_id: self.save_name_edit(rid))
        self.name_edit_entry.bind("<Escape>", lambda e: self.close_name_editor())

    def save_name_edit(self, row_id):
        if not self.name_edit_entry:
            return
        idx = int(row_id)
        old_name = self.groups[idx]["name"]
        new_name = self.name_edit_var.get().strip()
        if not new_name:
            show_toast(self, "分类名称不能为空", "warning")
            self.close_name_editor()
            return
        if new_name != old_name and new_name in [g["name"] for i, g in enumerate(self.groups) if i != idx]:
            show_toast(self, "该分类已存在", "warning")
            self.close_name_editor()
            return
        if new_name != old_name:
            self.groups[idx]["name"] = new_name
            for category in self.categories:
                if category.get("group", "") == old_name:
                    category["group"] = new_name
            save_groups(self.groups)
            save_categories(self.categories)
            self.callback(self.groups, self.categories)
        self.close_name_editor()
        self.refresh_list()
        self.group_tree.selection_set(str(idx))
        self.group_tree.focus(str(idx))

    def close_name_editor(self):
        if self.name_edit_entry:
            self.name_edit_entry.destroy()
            self.name_edit_entry = None

    def refresh_list(self):
        self.close_name_editor()
        for row in self.group_tree.get_children():
            self.group_tree.delete(row)
        for i, g in enumerate(self.groups):
            checked = "☑" if g.get("per_root_enabled", False) else "☐"
            color = g.get("color") or default_group_color(i)
            tag_name = f"group_color_{i}"
            self.group_tree.tag_configure(tag_name, foreground=color)
            self.group_tree.insert("", tk.END, iid=str(i), values=(g["name"], checked, "■"), tags=(tag_name,))
        self.update_color_buttons()

    def add_group(self):
        new_group = self.new_group_var.get().strip()
        if not new_group:
            show_toast(self, "请输入分类名称", "warning")
            return
        existing_names = [g["name"] for g in self.groups]
        if new_group in existing_names:
            show_toast(self, "该分类已存在", "warning")
            return
        new_entry = {
            "name": new_group,
            "per_root_enabled": False,
            "color": default_group_color(len(self.groups)),
        }
        self.groups.append(new_entry)
        self.groups.sort(key=lambda x: x["name"])
        self.new_group_var.set("")
        self.refresh_list()
        save_groups(self.groups)
        self.callback(self.groups, self.categories)

    def delete_group(self):
        selection = self.group_tree.selection()
        if not selection:
            show_toast(self, "请先选中要删除的分类", "warning")
            return
        idx = int(selection[0])
        group_name = self.groups[idx]["name"]
        if messagebox.askyesno("确认删除", f"确定删除分类【{group_name}】吗？\n注意：该分类下的所有方案也会被删除！"):
            del self.groups[idx]
            self.categories = [c for c in self.categories if c.get("group", "") != group_name]
            self.refresh_list()
            save_groups(self.groups)
            self.callback(self.groups, self.categories)

    def on_close(self):
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()


# ==================== 赔付不同意记录窗口 ====================
class LogViewerDialog(tk.Toplevel):
    def __init__(self, parent, on_close_callback=None):
        super().__init__(parent)
        self.title("赔付不同意记录")
        self.geometry("900x520")
        self.minsize(760, 420)
        self.configure(bg=APP_BG)
        self.on_close_callback = on_close_callback
        self.logs = load_log()
        self.summary_vars = {
            "count": tk.StringVar(),
            "amount": tk.StringVar(),
            "top_reason": tk.StringVar(),
        }

        self.build_ui()
        self.refresh_tree()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 900
        window_height = 520
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def build_ui(self):
        main_frame = ttk.Frame(self, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        summary_frame = ttk.Frame(main_frame)
        summary_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(summary_frame, textvariable=self.summary_vars["count"], style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(summary_frame, textvariable=self.summary_vars["amount"], style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 18))
        ttk.Label(summary_frame, textvariable=self.summary_vars["top_reason"], style="Section.TLabel").pack(side=tk.LEFT)

        columns = ("time", "amount", "reason", "roots", "ratio", "refund", "count", "event")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", selectmode="extended")
        headings = {
            "time": "时间",
            "amount": "实付金额",
            "reason": "退款原因",
            "roots": "根数",
            "ratio": "赔付比例",
            "refund": "赔付金额",
            "count": "不同意次数",
            "event": "事件",
        }
        widths = {
            "time": 145,
            "amount": 75,
            "reason": 120,
            "roots": 105,
            "ratio": 75,
            "refund": 85,
            "count": 80,
            "event": 110,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor=tk.CENTER)
        self.tree.column("reason", anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 12))
        action_button(btn_frame, text="刷新", command=self.reload_logs, variant="secondary").pack(side=tk.LEFT, padx=5)
        action_button(btn_frame, text="分析", command=self.open_analysis, variant="primary").pack(side=tk.LEFT, padx=5)
        action_button(btn_frame, text="删除选中", command=self.delete_selected, variant="danger").pack(side=tk.LEFT, padx=5)
        action_button(btn_frame, text="关闭", command=self.on_close, variant="ghost").pack(side=tk.RIGHT, padx=5)

    def reload_logs(self):
        self.logs = load_log()
        self.refresh_tree()

    def update_summary(self):
        total_count = len(self.logs)
        total_amount = 0
        by_reason = {}
        for log in self.logs:
            try:
                refund_money = float(log.get("refund_money") or 0)
            except (TypeError, ValueError):
                refund_money = 0
            total_amount += refund_money
            reason = log.get("reason", "未知")
            by_reason[reason] = by_reason.get(reason, 0) + 1
        top_reason = max(by_reason.items(), key=lambda item: item[1])[0] if by_reason else "无"
        self.summary_vars["count"].set(f"记录数：{total_count}")
        self.summary_vars["amount"].set(f"累计被拒赔付：{round(total_amount, 2)} 元")
        self.summary_vars["top_reason"].set(f"最多原因：{top_reason}")

    def refresh_tree(self):
        self.update_summary()
        for row in self.tree.get_children():
            self.tree.delete(row)
        for original_idx in range(len(self.logs) - 1, -1, -1):
            log = self.logs[original_idx]
            total_roots = log.get("total_roots", "")
            bad_roots = log.get("bad_roots", "")
            roots = f"{bad_roots}/{total_roots}" if total_roots and bad_roots else "-"
            ratio = log.get("ratio", 0)
            try:
                ratio_text = f"{int(float(ratio) * 100)}%"
            except (TypeError, ValueError):
                ratio_text = "-"
            refund = log.get("refund_money")
            if refund is None:
                refund = round(float(log.get("amount", 0)) * float(log.get("ratio", 0)), 2)
            self.tree.insert("", tk.END, iid=str(original_idx), values=(
                log.get("time", "-"),
                log.get("amount", ""),
                log.get("reason", ""),
                roots,
                ratio_text,
                refund,
                log.get("disagree_count", 1),
                log.get("event", "赔付不同意"),
            ))

    def delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            show_toast(self, "请先选中要删除的日志记录", "warning")
            return
        count = len(selection)
        if not messagebox.askyesno("确认删除", f"确定删除选中的 {count} 条记录吗？", parent=self):
            return
        delete_indexes = sorted((int(item) for item in selection), reverse=True)
        for idx in delete_indexes:
            if 0 <= idx < len(self.logs):
                del self.logs[idx]
        save_log(self.logs)
        self.refresh_tree()
        show_toast(self, f"已删除 {count} 条记录", "success")

    def open_analysis(self):
        if not self.logs:
            show_toast(self, "暂无日志可分析", "info")
            return

        total = len(self.logs)
        stats = {}
        for log in self.logs:
            reason = log.get("reason", "未知")
            ratio = log.get("ratio", 0)
            try:
                ratio_percent = int(round(float(ratio) * 100))
            except (TypeError, ValueError):
                ratio_percent = 0
            key = (reason, ratio_percent)
            stats[key] = stats.get(key, 0) + 1

        dialog = tk.Toplevel(self)
        dialog.title("拒绝赔付分析")
        dialog.geometry("560x460")
        dialog.minsize(520, 380)
        dialog.configure(bg=APP_BG)
        dialog.transient(self)

        main_frame = ttk.Frame(dialog, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text=f"共 {total} 条拒绝记录，以下占比为各类型在拒绝记录中的占比。", style="Muted.TLabel").pack(anchor=tk.W, pady=(0, 10))

        columns = ("reason", "ratio", "count", "percent")
        analysis_tree = ttk.Treeview(main_frame, columns=columns, show="headings", selectmode="browse")
        analysis_tree.heading("reason", text="退款原因")
        analysis_tree.heading("ratio", text="赔付比例")
        analysis_tree.heading("count", text="拒绝次数")
        analysis_tree.heading("percent", text="占比")
        analysis_tree.column("reason", width=180, anchor=tk.CENTER)
        analysis_tree.column("ratio", width=100, anchor=tk.CENTER)
        analysis_tree.column("count", width=90, anchor=tk.CENTER)
        analysis_tree.column("percent", width=90, anchor=tk.CENTER)

        sorted_stats = sorted(stats.items(), key=lambda item: item[1], reverse=True)
        for i, ((reason, ratio_percent), count) in enumerate(sorted_stats):
            percent = count / total * 100
            analysis_tree.insert("", tk.END, iid=str(i), values=(reason, f"{ratio_percent}%", count, f"{percent:.1f}%"))

        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=analysis_tree.yview)
        analysis_tree.configure(yscrollcommand=scrollbar.set)
        analysis_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 12))
        action_button(btn_frame, text="关闭", command=dialog.destroy, variant="ghost").pack(side=tk.RIGHT, padx=5)

    def on_close(self):
        if self.on_close_callback:
            self.on_close_callback()
        self.destroy()


# ==================== 主程序入口 ====================
if __name__ == "__main__":
    root = tk.Tk()
    app = RefundApp(root)
    root.mainloop()
