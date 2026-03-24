
import csv
import math
import random
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class RandomPickerApp:
    """随机选号系统（课程设计版 + 桌面小人多形象版）"""

    CHARACTER_SPRITES = [
        ("char_violet.png", "薇尔莉特"),
        ("char_rudeus.png", "鲁迪乌斯"),
        ("char_asuna.png", "亚斯娜"),
        ("char_luoxiaohei_human.png","罗小黑"),
        ("char_frieren.png","芙莉莲"),
        ("char_gojo.png","五条悟"),
        ("char_misaka .png", "御坂美琴"),
        ("char_luoxiaohei.png", "罗小黑"),
        ("char_wangquan.png", "王权富贵"),
        ("char_fengbaobao.png", "冯宝宝"),
        ("char_itachi.png", "宇智波鼬"),
        ("char_luffy.png", "路飞"),
        ("kuma.png","樱岛麻衣"),
        ("char_hauizhu.png","东方淮竹"),
        ("char_violet_2.png","薇尔莉特"),
        ("char_Miyamizu_ Mitsuha.png","宫水三叶"),
    ]

    TRANSPARENT_KEY_COLOR = "#f5efe6"
    DESKTOP_IMAGE_MAX_HEIGHT = 420

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("随机选号系统")
        self.root.geometry("1180x760")
        self.root.minsize(1080, 700)

        self.students: list[str] = []
        self.history: list[dict] = []
        self.last_result: list[str] = []

        self.character_window: tk.Toplevel | None = None
        self.character_bubble: tk.Toplevel | None = None
        self.character_label: tk.Label | None = None
        self.character_hint_label: tk.Label | None = None
        self.character_images: list[tk.PhotoImage] = []
        self.character_image_paths: list[Path] = []
        self.character_names: list[str] = []
        self.character_pose_index = 0
        self.character_drag_offset = (0, 0)
        self.character_has_moved = False
        self.character_pending_switch_job = None
        self.character_last_message = ""

        self._build_style()
        self._build_ui()
        self._load_character_images()
        self._create_desktop_character()
        self._refresh_all_views()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ========================= 界面构建 =========================
    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 10))
        style.configure("Panel.TLabelframe", padding=10)
        style.configure("Panel.TLabelframe.Label", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="both", expand=True)

        header = ttk.Frame(top)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text="Python课程设计：随机选号系统", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="支持手动输入 / TXT、CSV导入 / 不重复随机抽取 / 结果导出 / 历史记录 / 桌面小人抽签播报 / 多角色形象切换（已扩展 13 个角色）",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        stats_frame = ttk.Frame(top)
        stats_frame.pack(fill="x", pady=(0, 10))
        self.total_var = tk.StringVar(value="当前名单人数：0")
        self.unique_var = tk.StringVar(value="有效去重后人数：0")
        self.history_var = tk.StringVar(value="抽取历史次数：0")
        ttk.Label(stats_frame, textvariable=self.total_var).pack(side="left", padx=(0, 20))
        ttk.Label(stats_frame, textvariable=self.unique_var).pack(side="left", padx=(0, 20))
        ttk.Label(stats_frame, textvariable=self.history_var).pack(side="left")

        content = ttk.Frame(top)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=5)
        content.columnconfigure(1, weight=4)
        content.rowconfigure(0, weight=1)

        left = ttk.Frame(content)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.rowconfigure(1, weight=1)
        left.rowconfigure(3, weight=1)
        left.columnconfigure(0, weight=1)

        right = ttk.Frame(content)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=2)
        right.rowconfigure(3, weight=3)
        right.columnconfigure(0, weight=1)

        # 左上：输入区域
        input_box = ttk.LabelFrame(left, text="1. 名单输入与导入", style="Panel.TLabelframe")
        input_box.grid(row=0, column=0, sticky="nsew", pady=(0, 8))

        ttk.Label(
            input_box,
            text="可直接输入姓名/学号，每行一个；也支持逗号、顿号、分号分隔。",
        ).pack(anchor="w", pady=(0, 6))

        self.input_text = tk.Text(input_box, height=10, font=("Consolas", 11), wrap="word")
        self.input_text.pack(fill="both", expand=True)

        input_btns = ttk.Frame(input_box)
        input_btns.pack(fill="x", pady=(8, 0))
        ttk.Button(input_btns, text="从输入框加载名单", command=self.load_from_text).pack(side="left", padx=(0, 6))
        ttk.Button(input_btns, text="导入 TXT/CSV 文件", command=self.import_file).pack(side="left", padx=(0, 6))
        ttk.Button(input_btns, text="追加导入", command=lambda: self.import_file(append=True)).pack(side="left", padx=(0, 6))
        ttk.Button(input_btns, text="清空输入框", command=lambda: self.input_text.delete("1.0", tk.END)).pack(side="left")

        # 左下：名单区域
        list_box = ttk.LabelFrame(left, text="2. 当前名单", style="Panel.TLabelframe")
        list_box.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        list_box.rowconfigure(0, weight=1)
        list_box.columnconfigure(0, weight=1)

        self.student_listbox = tk.Listbox(list_box, font=("Consolas", 11))
        self.student_listbox.grid(row=0, column=0, sticky="nsew")
        list_scroll = ttk.Scrollbar(list_box, orient="vertical", command=self.student_listbox.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.student_listbox.config(yscrollcommand=list_scroll.set)

        list_btns = ttk.Frame(list_box)
        list_btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(list_btns, text="删除选中项", command=self.delete_selected_student).pack(side="left", padx=(0, 6))
        ttk.Button(list_btns, text="清空名单", command=self.clear_students).pack(side="left", padx=(0, 6))
        ttk.Button(list_btns, text="导出名单", command=self.export_students).pack(side="left")

        # 右上：抽取控制
        control_box = ttk.LabelFrame(right, text="3. 抽取设置", style="Panel.TLabelframe")
        control_box.grid(row=0, column=0, sticky="nsew", pady=(0, 8))

        line1 = ttk.Frame(control_box)
        line1.pack(fill="x", pady=(0, 6))
        ttk.Label(line1, text="抽取人数：").pack(side="left")
        self.count_var = tk.IntVar(value=1)
        ttk.Spinbox(line1, from_=1, to=9999, width=8, textvariable=self.count_var).pack(side="left", padx=(0, 12))

        self.remove_after_pick_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            line1,
            text="抽取后从名单中移除（连续抽签不重复）",
            variable=self.remove_after_pick_var,
        ).pack(side="left")

        line2 = ttk.Frame(control_box)
        line2.pack(fill="x", pady=(0, 6))
        self.sort_result_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(line2, text="按姓名排序显示结果", variable=self.sort_result_var).pack(side="left")

        line3 = ttk.Frame(control_box)
        line3.pack(fill="x")
        ttk.Label(
            line3,
            text="提示：点击桌面上的动漫小人，也可以直接触发一次新的随机抽取。",
        ).pack(side="left")

        btn_line = ttk.Frame(control_box)
        btn_line.pack(fill="x", pady=(6, 0))
        ttk.Button(btn_line, text="开始随机抽取", style="Accent.TButton", command=self.pick_students).pack(side="left", padx=(0, 8))
        ttk.Button(btn_line, text="再次抽取", command=self.pick_students).pack(side="left", padx=(0, 8))
        ttk.Button(btn_line, text="导出本次结果", command=self.export_last_result).pack(side="left", padx=(0, 8))
        ttk.Button(btn_line, text="召回桌面小人", command=self.summon_desktop_character).pack(side="left", padx=(0, 8))
        ttk.Button(btn_line, text="手动切换角色", command=self.switch_character_pose).pack(side="left")

        # 右中：结果展示
        result_box = ttk.LabelFrame(right, text="4. 抽取结果", style="Panel.TLabelframe")
        result_box.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        result_box.rowconfigure(0, weight=1)
        result_box.columnconfigure(0, weight=1)

        self.result_text = tk.Text(result_box, font=("Consolas", 12), wrap="word", state="disabled")
        self.result_text.grid(row=0, column=0, sticky="nsew")
        result_scroll = ttk.Scrollbar(result_box, orient="vertical", command=self.result_text.yview)
        result_scroll.grid(row=0, column=1, sticky="ns")
        self.result_text.config(yscrollcommand=result_scroll.set)

        # 右下：历史记录
        history_box = ttk.LabelFrame(right, text="5. 历史记录", style="Panel.TLabelframe")
        history_box.grid(row=3, column=0, sticky="nsew")
        history_box.rowconfigure(0, weight=1)
        history_box.columnconfigure(0, weight=1)

        columns = ("time", "count", "result")
        self.history_tree = ttk.Treeview(history_box, columns=columns, show="headings", height=8)
        self.history_tree.heading("time", text="时间")
        self.history_tree.heading("count", text="人数")
        self.history_tree.heading("result", text="抽取结果")
        self.history_tree.column("time", width=160, anchor="center")
        self.history_tree.column("count", width=70, anchor="center")
        self.history_tree.column("result", width=360, anchor="w")
        self.history_tree.grid(row=0, column=0, sticky="nsew")

        history_scroll = ttk.Scrollbar(history_box, orient="vertical", command=self.history_tree.yview)
        history_scroll.grid(row=0, column=1, sticky="ns")
        self.history_tree.config(yscrollcommand=history_scroll.set)

        history_btns = ttk.Frame(history_box)
        history_btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(history_btns, text="导出历史记录", command=self.export_history).pack(side="left", padx=(0, 6))
        ttk.Button(history_btns, text="清空历史记录", command=self.clear_history).pack(side="left")

    # ========================= 桌面小人 =========================
    def _character_assets_root(self) -> Path:
        return Path(__file__).resolve().parent

    def _load_character_images(self):
        self.character_images.clear()
        self.character_image_paths.clear()
        self.character_names.clear()

        for filename, display_name in self.CHARACTER_SPRITES:
            image_path = self._character_assets_root() / filename
            if not image_path.exists():
                continue

            display_image = tk.PhotoImage(file=str(image_path))
            if display_image.height() > self.DESKTOP_IMAGE_MAX_HEIGHT:
                factor = max(1, math.ceil(display_image.height() / self.DESKTOP_IMAGE_MAX_HEIGHT))
                display_image = display_image.subsample(factor, factor)

            self.character_images.append(display_image)
            self.character_image_paths.append(image_path)
            self.character_names.append(display_name)

    def _create_desktop_character(self):
        self.character_window = tk.Toplevel(self.root)
        self.character_window.overrideredirect(True)
        self.character_window.attributes("-topmost", True)
        self.character_window.config(bg=self.TRANSPARENT_KEY_COLOR)

        try:
            self.character_window.wm_attributes("-transparentcolor", self.TRANSPARENT_KEY_COLOR)
        except tk.TclError:
            self.character_window.config(bg="#f8f4ec")

        container = tk.Frame(self.character_window, bg=self.character_window.cget("bg"), bd=0, highlightthickness=0)
        container.pack()

        self.character_label = tk.Label(
            container,
            bg=self.character_window.cget("bg"),
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self.character_label.pack()

        self.character_hint_label = tk.Label(
            container,
            text="",
            font=("Microsoft YaHei UI", 9),
            bg="#fff9ef",
            fg="#5b6170",
            padx=8,
            pady=3,
            cursor="hand2",
            justify="center",
        )
        self.character_hint_label.pack(pady=(2, 0))

        self._apply_current_character_pose()

        self.character_label.bind("<Button-1>", self._on_character_click)
        self.character_hint_label.bind("<Button-1>", self._on_character_click)

        self.character_label.bind("<ButtonPress-3>", self._start_drag_character)
        self.character_label.bind("<B3-Motion>", self._drag_character)
        self.character_hint_label.bind("<ButtonPress-3>", self._start_drag_character)
        self.character_hint_label.bind("<B3-Motion>", self._drag_character)

        self.summon_desktop_character()

    def _apply_current_character_pose(self):
        if not self.character_label:
            return

        if self.character_images:
            self.character_label.config(
                image=self.character_images[self.character_pose_index],
                text="",
                padx=0,
                pady=0,
                relief="flat",
            )
        else:
            self.character_label.config(
                text=self.current_character_name(),
                image="",
                font=("Microsoft YaHei UI", 12, "bold"),
                bg="#f8f4ec",
                fg="#294764",
                bd=1,
                relief="solid",
                padx=12,
                pady=20,
            )

        self._update_character_hint()

    def current_character_name(self) -> str:
        if self.character_names and 0 <= self.character_pose_index < len(self.character_names):
            return self.character_names[self.character_pose_index]
        return "桌面小人"

    def _update_character_hint(self):
        if not self.character_hint_label:
            return
        self.character_hint_label.config(
            text=f"当前角色：{self.current_character_name()}\n左键随机抽取｜右键拖动｜共 {len(self.character_names)} 个角色"
        )

    def summon_desktop_character(self):
        if not self.character_window or not self.character_window.winfo_exists():
            return
        self.character_window.deiconify()
        self.character_window.lift()
        self.root.update_idletasks()
        self.character_window.update_idletasks()

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        char_w = self.character_window.winfo_reqwidth()
        char_h = self.character_window.winfo_reqheight()

        x = max(0, screen_w - char_w - 30)
        y = max(0, screen_h - char_h - 80)
        self.character_window.geometry(f"+{x}+{y}")

    def _cancel_pending_character_switch(self):
        if self.character_pending_switch_job is None:
            return
        if self.character_window and self.character_window.winfo_exists():
            try:
                self.character_window.after_cancel(self.character_pending_switch_job)
            except tk.TclError:
                pass
        self.character_pending_switch_job = None

    def switch_character_pose(self, refresh_bubble: bool = True):
        if not self.character_images:
            return
        if len(self.character_images) == 1:
            self.character_pose_index = 0
        else:
            choices = [idx for idx in range(len(self.character_images)) if idx != self.character_pose_index]
            self.character_pose_index = random.choice(choices)
        self._apply_current_character_pose()

        if refresh_bubble and self.character_bubble and self.character_bubble.winfo_exists():
            message = self.character_last_message or self._build_character_message()
            self.show_character_message(message)

    def _schedule_character_pose_switch(self, delay_ms: int = 700):
        if not self.character_window or not self.character_window.winfo_exists():
            return
        if self.character_pending_switch_job is not None:
            try:
                self.character_window.after_cancel(self.character_pending_switch_job)
            except tk.TclError:
                pass
        self.character_pending_switch_job = self.character_window.after(delay_ms, self._delayed_pose_switch)

    def _delayed_pose_switch(self):
        self.character_pending_switch_job = None
        self.switch_character_pose(refresh_bubble=True)

    def _start_drag_character(self, event):
        self.character_drag_offset = (event.x_root, event.y_root)
        self.character_has_moved = False

    def _drag_character(self, event):
        if not self.character_window:
            return

        old_x, old_y = self.character_drag_offset
        dx = event.x_root - old_x
        dy = event.y_root - old_y

        new_x = self.character_window.winfo_x() + dx
        new_y = self.character_window.winfo_y() + dy
        self.character_window.geometry(f"+{new_x}+{new_y}")
        self.character_drag_offset = (event.x_root, event.y_root)
        self.character_has_moved = True

        if self.character_bubble and self.character_bubble.winfo_exists():
            self.character_bubble.destroy()

    def _on_character_click(self, _event=None):
        self.pick_students(from_character=True)

    def show_character_message(self, message: str | None = None):
        if not self.character_window or not self.character_window.winfo_exists():
            return

        if self.character_bubble and self.character_bubble.winfo_exists():
            self.character_bubble.destroy()

        if message is None:
            message = self._build_character_message()
        self.character_last_message = message

        self.character_bubble = tk.Toplevel(self.root)
        self.character_bubble.overrideredirect(True)
        self.character_bubble.attributes("-topmost", True)
        self.character_bubble.config(bg="#d8c7a0")

        body = tk.Frame(self.character_bubble, bg="#fffaf1", padx=14, pady=12)
        body.pack(fill="both", expand=True, padx=1, pady=1)

        tk.Label(
            body,
            text=self.current_character_name(),
            font=("Microsoft YaHei UI", 12, "bold"),
            bg="#fffaf1",
            fg="#2d4c67",
        ).pack(anchor="w")

        tk.Label(
            body,
            text=message,
            font=("Microsoft YaHei UI", 10),
            bg="#fffaf1",
            fg="#333333",
            justify="left",
            wraplength=320,
        ).pack(anchor="w", pady=(6, 0))

        tk.Label(
            body,
            text="点击气泡可关闭",
            font=("Microsoft YaHei UI", 8),
            bg="#fffaf1",
            fg="#7a7a7a",
        ).pack(anchor="e", pady=(8, 0))

        self.character_bubble.bind("<Button-1>", lambda _e: self.character_bubble.destroy())
        body.bind("<Button-1>", lambda _e: self.character_bubble.destroy())

        self.character_bubble.update_idletasks()
        bubble_w = self.character_bubble.winfo_width()
        bubble_h = self.character_bubble.winfo_height()
        char_x = self.character_window.winfo_x()
        char_y = self.character_window.winfo_y()
        char_w = self.character_window.winfo_width()

        x = max(10, char_x + char_w - bubble_w + 10)
        y = max(10, char_y - bubble_h - 12)
        self.character_bubble.geometry(f"+{x}+{y}")
        self.character_bubble.after(9000, self._safe_close_character_bubble)

    def _safe_close_character_bubble(self):
        if self.character_bubble and self.character_bubble.winfo_exists():
            self.character_bubble.destroy()

    def _build_character_message(self) -> str:
        if not self.last_result:
            return "现在还没有抽签结果。你可以先在主窗口点击“开始随机抽取”，也可以直接点我开始抽签。抽完后我会播报名单，并切换到下一位角色。"

        latest = "\n".join(f"{idx}. {name}" for idx, name in enumerate(self.last_result, start=1))
        count = len(self.last_result)
        openings = [
            "抽签已经完成，我来为你播报结果。",
            "新的抽取结果已经出来了。",
            "本轮名单如下，请确认。",
            "我已经完成随机抽取。",
        ]
        opening = random.choice(openings)
        return f"{opening}\n本次一共抽中了 {count} 人：\n{latest}"

    # ========================= 数据处理 =========================
    @staticmethod
    def normalize_name(raw: str) -> str:
        return " ".join(raw.strip().split())

    def parse_names(self, text: str) -> list[str]:
        for sep in [",", "，", "、", ";", "；", "\t"]:
            text = text.replace(sep, "\n")
        items = [self.normalize_name(item) for item in text.splitlines()]
        return [item for item in items if item]

    @staticmethod
    def looks_like_id(value: str) -> bool:
        cleaned = value.strip().replace(" ", "").replace("-", "")
        if not cleaned:
            return False
        return cleaned.isdigit()

    @staticmethod
    def is_header_text(value: str) -> bool:
        normalized = value.strip().lower().replace(" ", "")
        headers = {
            "学号", "姓名", "名字", "名称", "编号", "序号", "id", "name",
            "studentid", "studentname", "学生姓名", "学生学号", "stu_id", "stu_name",
        }
        return normalized in headers

    def detect_header_mapping(self, row: list[str]) -> tuple[int | None, int | None]:
        id_col = None
        name_col = None
        for idx, cell in enumerate(row):
            normalized = cell.strip().lower().replace(" ", "")
            if normalized in {"学号", "编号", "id", "studentid", "学生学号", "stu_id"}:
                id_col = idx
            elif normalized in {"姓名", "名字", "名称", "name", "studentname", "学生姓名", "stu_name"}:
                name_col = idx
        return id_col, name_col

    def format_student_record(self, student_id: str, student_name: str) -> str:
        student_id = self.normalize_name(student_id)
        student_name = self.normalize_name(student_name)
        if student_name and student_id:
            return f"{student_name}（{student_id}）"
        return student_name or student_id

    def parse_csv_rows(self, rows: list[list[str]]) -> list[str]:
        records: list[str] = []
        if not rows:
            return records

        normalized_rows = []
        for row in rows:
            cleaned_row = [self.normalize_name(cell) for cell in row]
            if any(cleaned_row):
                normalized_rows.append(cleaned_row)

        if not normalized_rows:
            return records

        start_index = 0
        id_col = None
        name_col = None

        first_row = normalized_rows[0]
        detected_id_col, detected_name_col = self.detect_header_mapping(first_row)
        non_empty_first = [cell for cell in first_row if cell]
        if non_empty_first and all(self.is_header_text(cell) for cell in non_empty_first):
            id_col, name_col = detected_id_col, detected_name_col
            start_index = 1

        for row in normalized_rows[start_index:]:
            non_empty = [cell for cell in row if cell]
            if not non_empty:
                continue

            if id_col is not None or name_col is not None:
                student_id = row[id_col] if id_col is not None and id_col < len(row) else ""
                student_name = row[name_col] if name_col is not None and name_col < len(row) else ""
                record = self.format_student_record(student_id, student_name)
                if record and not self.is_header_text(record):
                    records.append(record)
                continue

            if len(non_empty) >= 2:
                first, second = non_empty[0], non_empty[1]
                if self.looks_like_id(first) and not self.looks_like_id(second):
                    records.append(self.format_student_record(first, second))
                    continue
                if self.looks_like_id(second) and not self.looks_like_id(first):
                    records.append(self.format_student_record(second, first))
                    continue

            for cell in non_empty:
                if self.is_header_text(cell):
                    continue
                parsed = self.parse_names(cell)
                records.extend([item for item in parsed if not self.is_header_text(item)])

        return records

    def deduplicate_keep_order(self, names: list[str]) -> list[str]:
        seen = set()
        result = []
        for name in names:
            if name not in seen:
                seen.add(name)
                result.append(name)
        return result

    def load_from_text(self):
        raw = self.input_text.get("1.0", tk.END)
        names = self.parse_names(raw)
        if not names:
            messagebox.showwarning("提示", "输入框中没有可用名单，请先输入姓名或学号。")
            return
        self.students = self.deduplicate_keep_order(names)
        self._refresh_all_views()
        messagebox.showinfo("成功", f"名单加载完成，共导入 {len(names)} 条，去重后 {len(self.students)} 人。")

    def import_file(self, append: bool = False):
        file_path = filedialog.askopenfilename(
            title="选择名单文件",
            filetypes=[("文本和CSV文件", "*.txt *.csv"), ("文本文件", "*.txt"), ("CSV文件", "*.csv")],
        )
        if not file_path:
            return

        try:
            imported = self.read_names_from_file(file_path)
        except Exception as exc:
            messagebox.showerror("导入失败", f"文件读取失败：\n{exc}")
            return

        if not imported:
            messagebox.showwarning("提示", "文件中没有读取到有效名单。")
            return

        if append:
            merged = self.students + imported
            self.students = self.deduplicate_keep_order(merged)
            messagebox.showinfo("成功", f"追加导入成功，本次读取 {len(imported)} 人，当前共 {len(self.students)} 人。")
        else:
            self.students = self.deduplicate_keep_order(imported)
            messagebox.showinfo("成功", f"文件导入成功，共读取 {len(imported)} 人，去重后 {len(self.students)} 人。")
        self._refresh_all_views()

    def read_names_from_file(self, file_path: str) -> list[str]:
        path = Path(file_path)
        suffix = path.suffix.lower()
        names: list[str] = []

        if suffix == ".txt":
            for encoding in ("utf-8-sig", "utf-8", "gbk"):
                try:
                    text = path.read_text(encoding=encoding)
                    names = self.parse_names(text)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise UnicodeDecodeError("txt", b"", 0, 1, "无法识别文件编码")

        elif suffix == ".csv":
            for encoding in ("utf-8-sig", "utf-8", "gbk"):
                try:
                    with open(path, "r", encoding=encoding, newline="") as f:
                        reader = csv.reader(f)
                        rows = list(reader)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise UnicodeDecodeError("csv", b"", 0, 1, "无法识别文件编码")

            names = self.parse_csv_rows(rows)
        else:
            raise ValueError("仅支持 .txt 或 .csv 文件。")

        return self.deduplicate_keep_order(names)

    def delete_selected_student(self):
        selected_indices = list(self.student_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("提示", "请先在名单中选中要删除的项。")
            return
        selected_indices.reverse()
        for index in selected_indices:
            self.students.pop(index)
        self._refresh_all_views()

    def clear_students(self):
        if not self.students:
            messagebox.showinfo("提示", "当前名单已为空。")
            return
        if messagebox.askyesno("确认", "确定要清空当前名单吗？"):
            self.students.clear()
            self.last_result.clear()
            self._set_result_text("当前还没有抽取结果。")
            self._refresh_all_views()
            self.show_character_message("名单已经清空。重新导入名单后，点击我就可以继续随机抽取。")

    def pick_students(self, from_character: bool = False):
        if not self.students:
            if from_character:
                self.show_character_message("当前名单为空。请先在主窗口输入或导入名单，然后再点击我开始抽取。")
            else:
                messagebox.showwarning("提示", "当前名单为空，请先输入或导入名单。")
            return

        try:
            count = int(self.count_var.get())
        except (ValueError, tk.TclError):
            if from_character:
                self.show_character_message("抽取人数设置有误，请先把主窗口中的“抽取人数”改成正整数。")
            else:
                messagebox.showerror("错误", "抽取人数必须是正整数。")
            return

        if count <= 0:
            if from_character:
                self.show_character_message("抽取人数必须大于 0。请先修改主窗口里的抽取人数。")
            else:
                messagebox.showerror("错误", "抽取人数必须大于 0。")
            return
        if count > len(self.students):
            text = f"当前名单只有 {len(self.students)} 人，无法抽取 {count} 人。"
            if from_character:
                self.show_character_message(text)
            else:
                messagebox.showerror("错误", text)
            return

        result = random.sample(self.students, count)
        if self.sort_result_var.get():
            result = sorted(result)

        self.last_result = result[:]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history.append({
            "time": now,
            "count": count,
            "result": result[:],
        })

        display_lines = [
            f"抽取时间：{now}",
            f"抽取人数：{count}",
            "抽取结果：",
            "-" * 30,
        ]
        display_lines.extend([f"{idx}. {name}" for idx, name in enumerate(result, start=1)])
        self._set_result_text("\n".join(display_lines))

        if self.remove_after_pick_var.get():
            selected_set = set(result)
            self.students = [name for name in self.students if name not in selected_set]

        self._cancel_pending_character_switch()
        self._refresh_all_views()
        self.root.update_idletasks()

        # 先切角色，再显示气泡，保证图片、提示文字、气泡标题始终同步。
        self.switch_character_pose(refresh_bubble=False)
        self.show_character_message()

    # ========================= 导出功能 =========================
    def export_students(self):
        if not self.students:
            messagebox.showwarning("提示", "当前没有可导出的名单。")
            return
        file_path = filedialog.asksaveasfilename(
            title="导出名单",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv")],
        )
        if not file_path:
            return

        try:
            self._write_list_to_file(file_path, self.students)
            messagebox.showinfo("成功", "名单导出成功。")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def export_last_result(self):
        if not self.last_result:
            messagebox.showwarning("提示", "当前还没有抽取结果可导出。")
            return
        file_path = filedialog.asksaveasfilename(
            title="导出本次抽取结果",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv")],
        )
        if not file_path:
            return

        try:
            self._write_list_to_file(file_path, self.last_result)
            messagebox.showinfo("成功", "抽取结果导出成功。")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def export_history(self):
        if not self.history:
            messagebox.showwarning("提示", "当前没有历史记录可导出。")
            return
        file_path = filedialog.asksaveasfilename(
            title="导出历史记录",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("文本文件", "*.txt")],
        )
        if not file_path:
            return

        path = Path(file_path)
        try:
            if path.suffix.lower() == ".txt":
                lines = []
                for idx, item in enumerate(self.history, start=1):
                    lines.append(f"第{idx}次抽取")
                    lines.append(f"时间：{item['time']}")
                    lines.append(f"人数：{item['count']}")
                    lines.append("结果：" + "、".join(item["result"]))
                    lines.append("-" * 40)
                path.write_text("\n".join(lines), encoding="utf-8-sig")
            else:
                with open(path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["序号", "时间", "抽取人数", "抽取结果"])
                    for idx, item in enumerate(self.history, start=1):
                        writer.writerow([idx, item["time"], item["count"], "、".join(item["result"])])
            messagebox.showinfo("成功", "历史记录导出成功。")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _write_list_to_file(self, file_path: str, data: list[str]):
        path = Path(file_path)
        if path.suffix.lower() == ".csv":
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["序号", "姓名/学号"])
                for idx, item in enumerate(data, start=1):
                    writer.writerow([idx, item])
        else:
            content = "\n".join(f"{idx}. {item}" for idx, item in enumerate(data, start=1))
            path.write_text(content, encoding="utf-8-sig")

    # ========================= 历史与刷新 =========================
    def clear_history(self):
        if not self.history:
            messagebox.showinfo("提示", "历史记录已为空。")
            return
        if messagebox.askyesno("确认", "确定要清空全部历史记录吗？"):
            self.history.clear()
            self._refresh_all_views()

    def _set_result_text(self, text: str):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", text)
        self.result_text.config(state="disabled")

    def _refresh_all_views(self):
        self.student_listbox.delete(0, tk.END)
        for idx, name in enumerate(self.students, start=1):
            self.student_listbox.insert(tk.END, f"{idx:>3}. {name}")

        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for row in self.history:
            self.history_tree.insert("", tk.END, values=(row["time"], row["count"], "、".join(row["result"])))

        self.total_var.set(f"当前名单人数：{len(self.students)}")
        self.unique_var.set(f"有效去重后人数：{len(set(self.students))}")
        self.history_var.set(f"抽取历史次数：{len(self.history)}")

        if not self.last_result:
            self._set_result_text("当前还没有抽取结果。")

    def on_close(self):
        self._safe_close_character_bubble()
        if self.character_pending_switch_job is not None and self.character_window:
            try:
                self.character_window.after_cancel(self.character_pending_switch_job)
            except tk.TclError:
                pass
        if self.character_window and self.character_window.winfo_exists():
            self.character_window.destroy()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = RandomPickerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
