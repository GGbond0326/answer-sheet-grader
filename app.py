#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
答题卡批阅工具（铅笔涂卡）

依赖：
    pip install opencv-python numpy pillow
"""

from __future__ import annotations

import json
import os
import subprocess
import threading
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import cv2
    import numpy as np
    from PIL import Image, ImageTk
except Exception as import_error:
    raise SystemExit(
        "缺少依赖，请先安装：pip install opencv-python numpy pillow\n"
        f"详细错误：{import_error}"
    )

try:
    PIL_RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    PIL_RESAMPLE_LANCZOS = Image.LANCZOS


LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@dataclass
class QuestionConfig:
    qid: int
    rect: Tuple[int, int, int, int]
    qtype: str = "single"  # single / multiple
    option_count: int = 4
    correct: List[str] = field(default_factory=list)

    def normalize(self) -> None:
        self.option_count = max(2, min(8, int(self.option_count)))
        self.qtype = "multiple" if self.qtype == "multiple" else "single"
        normalized: List[str] = []
        for ch in self.correct:
            up = str(ch).strip().upper()
            if up in LETTERS[: self.option_count] and up not in normalized:
                normalized.append(up)
        self.correct = normalized


@dataclass
class ScoreRules:
    single_correct: float = 2.0
    single_wrong: float = 0.0
    single_blank: float = 0.0
    multi_full: float = 3.0
    multi_partial: float = 2.0
    multi_wrong: float = 0.0
    multi_blank: float = 0.0


class VoiceBroadcaster:
    """Windows 下用 System.Speech 做语音播报。"""

    def __init__(self) -> None:
        self._lock = threading.Lock()

    def speak(self, text: str, enabled: bool = True) -> None:
        if not enabled:
            return
        thread = threading.Thread(target=self._speak_worker, args=(text,), daemon=True)
        thread.start()

    def _speak_worker(self, text: str) -> None:
        with self._lock:
            safe_text = text.replace("'", "''")
            ps_cmd = (
                "Add-Type -AssemblyName System.Speech; "
                "$s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                "$s.Volume=100; "
                "$s.Rate=2; "
                f"$s.Speak('{safe_text}')"
            )
            try:
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_cmd],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except Exception:
                pass


class AnswerSheetGraderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("答题卡智能批阅（铅笔涂卡）")
        self.geometry("1600x920")
        self.minsize(1320, 760)

        self.template_path: str = ""
        self.template_bgr: Optional[np.ndarray] = None
        self.template_gray: Optional[np.ndarray] = None
        self.template_edge: Optional[np.ndarray] = None
        self.template_binary: Optional[np.ndarray] = None
        self.template_kp = None
        self.template_des = None
        self.template_anchor_quad: Optional[np.ndarray] = None
        self.template_markers: List[Tuple[int, int]] = []
        self.questions: List[QuestionConfig] = []
        self.tree_item_to_index: Dict[str, int] = {}
        self.tree: Optional[ttk.Treeview] = None

        self.canvas_scale: float = 1.0
        self.canvas_offset_x: int = 0
        self.canvas_offset_y: int = 0
        self.canvas_photo: Optional[ImageTk.PhotoImage] = None
        self.view_zoom: float = 1.0
        self.view_pan_x: int = 0
        self.view_pan_y: int = 0
        self._panning = False
        self._pan_start: Tuple[int, int] = (0, 0)
        self._pan_origin: Tuple[int, int] = (0, 0)
        self._drawing = False
        self._draw_start: Tuple[int, int] = (0, 0)
        self._draw_preview_id = None
        self._manual_anchor_points: List[Tuple[int, int]] = []
        self._edit_mode: str = ""
        self._edit_q_idx: Optional[int] = None
        self._edit_start_img: Tuple[int, int] = (0, 0)
        self._edit_start_rect: Tuple[int, int, int, int] = (0, 0, 0, 0)

        self.cap: Optional[cv2.VideoCapture] = None
        self.running = False
        self.after_job = None
        self.last_eval_time = 0.0
        self.preview_photo: Optional[ImageTk.PhotoImage] = None
        self.last_total_score: Optional[float] = None
        self.last_align_mode: str = "未开始"
        self.last_align_confidence: float = 0.0
        self.last_marker_score: float = -1.0
        self.last_skip_reason: str = ""
        self._stable_signature: str = ""
        self._stable_count: int = 0
        self._stable_total: float = 0.0
        self._stable_lines: List[str] = []
        self._last_emitted_signature: str = ""

        self.voice = VoiceBroadcaster()

        self._init_vars()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _init_vars(self) -> None:
        self.var_qid = tk.StringVar(value="")
        self.var_qtype = tk.StringVar(value="single")
        self.var_option_count = tk.StringVar(value="4")
        self.var_correct = tk.StringVar(value="")

        self.var_single_correct = tk.StringVar(value="2")
        self.var_single_wrong = tk.StringVar(value="0")
        self.var_single_blank = tk.StringVar(value="0")
        self.var_multi_full = tk.StringVar(value="3")
        self.var_multi_partial = tk.StringVar(value="2")
        self.var_multi_wrong = tk.StringVar(value="0")
        self.var_multi_blank = tk.StringVar(value="0")

        self.var_camera_source = tk.StringVar(value="0")
        self.var_interval_sec = tk.StringVar(value="2")
        self.var_abs_fill = tk.StringVar(value="0.12")
        self.var_rel_fill = tk.StringVar(value="0.65")
        self.var_use_alignment = tk.BooleanVar(value=True)
        self.var_use_anchor_align = tk.BooleanVar(value=True)
        self.var_enable_quality_gate = tk.BooleanVar(value=True)
        self.var_require_full_anchor = tk.BooleanVar(value=True)
        self.var_block_low_confidence = tk.BooleanVar(value=True)
        self.var_block_bad_alignment = tk.BooleanVar(value=True)
        self.var_enable_stable_score = tk.BooleanVar(value=True)
        self.var_use_marker_score = tk.BooleanVar(value=True)
        self.var_stable_frames_required = tk.StringVar(value="3")
        self.var_blur_threshold = tk.StringVar(value="85")
        self.var_glare_threshold = tk.StringVar(value="0.18")
        self.var_single_gap_threshold = tk.StringVar(value="0.045")
        self.var_single_ratio_threshold = tk.StringVar(value="1.25")
        self.var_align_conf_threshold = tk.StringVar(value="0.08")
        self.var_marker_score_threshold = tk.StringVar(value="0.45")
        self.var_min_card_ratio = tk.StringVar(value="0.20")
        self.var_voice_enabled = tk.BooleanVar(value=True)

        self.var_status = tk.StringVar(value="请先上传答题卡模板图片，然后框选题目区域。")
        self.var_total_score = tk.StringVar(value="总分：--")
        self.var_last_time = tk.StringVar(value="最近批阅：--")
        self.var_draw_mode = tk.BooleanVar(value=False)
        self.var_marker_mode = tk.BooleanVar(value=False)
        self.var_manual_anchor_mode = tk.BooleanVar(value=False)
        self.var_zoom = tk.StringVar(value="100%")

        self.var_rect_x = tk.StringVar(value="")
        self.var_rect_y = tk.StringVar(value="")
        self.var_rect_w = tk.StringVar(value="")
        self.var_rect_h = tk.StringVar(value="")
        self.var_nudge_step = tk.StringVar(value="3")

    def _build_ui(self) -> None:
        root_pane = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        root_pane.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(root_pane, padding=8)
        root_pane.add(left, weight=3)
        right = ttk.Frame(root_pane, padding=8)
        root_pane.add(right, weight=2)

        self._build_left_panel(left)
        self._build_right_panel(right)

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        toolbar_top = ttk.Frame(parent)
        toolbar_top.pack(fill=tk.X, pady=(0, 4))
        toolbar_bottom = ttk.Frame(parent)
        toolbar_bottom.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(toolbar_top, text="上传模板图", command=self.load_template_image).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar_top, text="保存配置", command=self.save_config).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar_top, text="加载配置", command=self.load_config).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar_top, text="识别定位外框", command=self.detect_template_anchor).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar_top, text="整图作为外框", command=self.use_full_image_anchor).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(toolbar_top, text="放大", command=lambda: self.change_zoom(1.25)).pack(
            side=tk.LEFT, padx=(10, 4)
        )
        ttk.Button(toolbar_top, text="缩小", command=lambda: self.change_zoom(0.8)).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(toolbar_top, text="适应", command=self.reset_view).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(toolbar_top, textvariable=self.var_zoom, foreground="#334155").pack(
            side=tk.LEFT, padx=(4, 0)
        )

        ttk.Checkbutton(
            toolbar_bottom,
            text="框选模式（开启后左键拖拽新增题目框）",
            variable=self.var_draw_mode,
            command=self._on_toggle_draw_mode,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(
            toolbar_bottom,
            text="手动四点外框",
            variable=self.var_manual_anchor_mode,
            command=self._on_toggle_manual_anchor_mode,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(
            toolbar_bottom,
            text="标记点模式",
            variable=self.var_marker_mode,
            command=self._on_toggle_marker_mode,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(toolbar_bottom, text="清空标记点", command=self.clear_template_markers).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Label(
            toolbar_bottom,
            text="提示：手动四点外框开启后，依次点击外框四角；标记点可多点添加。",
            foreground="#334155",
        ).pack(side=tk.LEFT)

        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="#1f1f1f", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<MouseWheel>", self._on_canvas_mousewheel)
        self.canvas.bind("<Button-4>", lambda e: self._on_canvas_mousewheel(e, wheel_dir=1))
        self.canvas.bind("<Button-5>", lambda e: self._on_canvas_mousewheel(e, wheel_dir=-1))
        self.canvas.bind("<ButtonPress-3>", self._on_pan_start)
        self.canvas.bind("<B3-Motion>", self._on_pan_move)
        self.canvas.bind("<ButtonRelease-3>", self._on_pan_end)

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True)

        tab_setup = ttk.Frame(notebook, padding=8)
        tab_grade = ttk.Frame(notebook, padding=8)
        notebook.add(tab_setup, text="1) 模板与规则")
        notebook.add(tab_grade, text="2) 实时批阅")

        self._build_setup_tab(tab_setup)
        self._build_grade_tab(tab_grade)

        status_bar = ttk.Label(parent, textvariable=self.var_status, foreground="#1d4ed8")
        status_bar.pack(fill=tk.X, pady=(8, 0))

    def _on_toggle_draw_mode(self) -> None:
        if self.var_draw_mode.get():
            self.var_marker_mode.set(False)
            self.var_manual_anchor_mode.set(False)
            self._manual_anchor_points.clear()

    def _on_toggle_marker_mode(self) -> None:
        if self.var_marker_mode.get():
            self.var_draw_mode.set(False)
            self.var_manual_anchor_mode.set(False)
            self._manual_anchor_points.clear()
            self.var_status.set("标记点模式已开启：在模板上点击添加标记点。")

    def _on_toggle_manual_anchor_mode(self) -> None:
        if self.var_manual_anchor_mode.get():
            self.var_draw_mode.set(False)
            self.var_marker_mode.set(False)
            self._manual_anchor_points.clear()
            self.var_status.set("手动四点外框模式：请依次点击外框四个角。")
        else:
            self._manual_anchor_points.clear()
            self._render_canvas()

    def _build_setup_tab(self, parent: ttk.Frame) -> None:
        tip = (
            "操作步骤：\n"
            "1. 上传空白答题卡模板图\n"
            "2. 勾选“框选模式”，在左图拖拽框出每道题的选项区域\n"
            "3. 右侧选择题目，设置题号、单选/多选、正确答案\n"
            "4. 关闭框选模式后，可直接拖动题目框；拖右下角红点可缩放\n"
            "5. 可用“手动四点外框”点击四角标定定位外框；“标记点模式”可添加参考点\n"
            "6. 左侧支持放大/缩小/滚轮缩放；右键拖动可平移\n"
            "7. 设置评分规则后保存配置"
        )
        ttk.Label(parent, text=tip, justify=tk.LEFT, foreground="#374151").pack(
            fill=tk.X, pady=(0, 8)
        )

        list_frame = ttk.LabelFrame(parent, text="题目列表", padding=6)
        list_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("qid", "qtype", "options", "correct", "rect")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=11)
        self.tree.heading("qid", text="题号")
        self.tree.heading("qtype", text="题型")
        self.tree.heading("options", text="选项数")
        self.tree.heading("correct", text="标准答案")
        self.tree.heading("rect", text="区域(x,y,w,h)")
        self.tree.column("qid", width=55, anchor=tk.CENTER)
        self.tree.column("qtype", width=70, anchor=tk.CENTER)
        self.tree.column("options", width=65, anchor=tk.CENTER)
        self.tree.column("correct", width=90, anchor=tk.CENTER)
        self.tree.column("rect", width=220, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        op_row = ttk.Frame(list_frame)
        op_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(op_row, text="删除选中题目", command=self.delete_selected_question).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(op_row, text="按题号排序", command=self.sort_questions).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(op_row, text="清空全部题目", command=self.clear_questions).pack(
            side=tk.LEFT
        )

        edit = ttk.LabelFrame(parent, text="选中题目设置", padding=8)
        edit.pack(fill=tk.X, pady=(8, 8))

        row1 = ttk.Frame(edit)
        row1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row1, text="题号").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.var_qid, width=8).pack(side=tk.LEFT, padx=(6, 14))
        ttk.Label(row1, text="题型").pack(side=tk.LEFT)
        qtype_box = ttk.Combobox(
            row1,
            textvariable=self.var_qtype,
            values=["single", "multiple"],
            width=10,
            state="readonly",
        )
        qtype_box.pack(side=tk.LEFT, padx=(6, 14))
        ttk.Label(row1, text="选项数").pack(side=tk.LEFT)
        ttk.Spinbox(
            row1,
            from_=2,
            to=8,
            textvariable=self.var_option_count,
            width=6,
        ).pack(side=tk.LEFT, padx=(6, 0))

        row2 = ttk.Frame(edit)
        row2.pack(fill=tk.X)
        ttk.Label(row2, text="标准答案（示例：A 或 A,C）").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.var_correct, width=20).pack(
            side=tk.LEFT, padx=(8, 10)
        )
        ttk.Button(row2, text="保存该题设置", command=self.save_selected_question).pack(
            side=tk.LEFT
        )

        row3 = ttk.Frame(edit)
        row3.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(row3, text="坐标 X").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.var_rect_x, width=6).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row3, text="Y").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.var_rect_y, width=6).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row3, text="W").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.var_rect_w, width=6).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row3, text="H").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.var_rect_h, width=6).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Button(row3, text="应用坐标", command=self.apply_selected_rect).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(row3, text="步长").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.var_nudge_step, width=4).pack(side=tk.LEFT, padx=(4, 6))
        ttk.Button(row3, text="←", command=lambda: self.nudge_selected_rect(-1, 0)).pack(side=tk.LEFT)
        ttk.Button(row3, text="→", command=lambda: self.nudge_selected_rect(1, 0)).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(row3, text="↑", command=lambda: self.nudge_selected_rect(0, -1)).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(row3, text="↓", command=lambda: self.nudge_selected_rect(0, 1)).pack(side=tk.LEFT, padx=(4, 0))

        score = ttk.LabelFrame(parent, text="评分规则（可自由修改）", padding=8)
        score.pack(fill=tk.X)

        grid = ttk.Frame(score)
        grid.pack(fill=tk.X)
        labels = [
            ("单选-答对", self.var_single_correct),
            ("单选-答错", self.var_single_wrong),
            ("单选-空白", self.var_single_blank),
            ("多选-全对", self.var_multi_full),
            ("多选-漏选", self.var_multi_partial),
            ("多选-错选/多选", self.var_multi_wrong),
            ("多选-空白", self.var_multi_blank),
        ]
        for i, (title, var) in enumerate(labels):
            r = i // 2
            c = (i % 2) * 2
            ttk.Label(grid, text=title).grid(row=r, column=c, sticky="w", padx=(0, 6), pady=4)
            ttk.Entry(grid, textvariable=var, width=10).grid(
                row=r, column=c + 1, sticky="w", padx=(0, 16), pady=4
            )

    def _build_grade_tab(self, parent: ttk.Frame) -> None:
        param = ttk.LabelFrame(parent, text="识别参数", padding=8)
        param.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(param)
        row1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row1, text="摄像头源（0/1 或手机URL）").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.var_camera_source, width=28).pack(
            side=tk.LEFT, padx=(8, 12)
        )
        ttk.Label(row1, text="识别间隔(秒)").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.var_interval_sec, width=7).pack(
            side=tk.LEFT, padx=(8, 12)
        )
        ttk.Checkbutton(row1, text="启用透视对齐", variable=self.var_use_alignment).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(row1, text="优先外框定位", variable=self.var_use_anchor_align).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(row1, text="语音播报", variable=self.var_voice_enabled).pack(side=tk.LEFT)

        row2 = ttk.Frame(param)
        row2.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row2, text="填涂绝对阈值(0~1)").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.var_abs_fill, width=8).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Label(row2, text="相对阈值(0~1)").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.var_rel_fill, width=8).pack(side=tk.LEFT, padx=(8, 0))

        row3 = ttk.Frame(param)
        row3.pack(fill=tk.X, pady=(0, 4))
        ttk.Checkbutton(row3, text="启用质量门控", variable=self.var_enable_quality_gate).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(row3, text="外框不完整不计分", variable=self.var_require_full_anchor).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(row3, text="单选低置信度不计分", variable=self.var_block_low_confidence).pack(
            side=tk.LEFT
        )
        ttk.Checkbutton(row3, text="对齐低置信度不计分", variable=self.var_block_bad_alignment).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Checkbutton(row3, text="使用标记点评分", variable=self.var_use_marker_score).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Checkbutton(row3, text="连续稳定判分", variable=self.var_enable_stable_score).pack(
            side=tk.LEFT, padx=(12, 0)
        )

        row4 = ttk.Frame(param)
        row4.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row4, text="清晰度阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_blur_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="反光阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_glare_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="单选差值阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_single_gap_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="单选比值阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_single_ratio_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="对齐置信阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_align_conf_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="标记阈值").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_marker_score_threshold, width=7).pack(
            side=tk.LEFT, padx=(6, 12)
        )
        ttk.Label(row4, text="最小卡片占比").pack(side=tk.LEFT)
        ttk.Entry(row4, textvariable=self.var_min_card_ratio, width=7).pack(
            side=tk.LEFT, padx=(6, 0)
        )

        row5 = ttk.Frame(param)
        row5.pack(fill=tk.X)
        ttk.Label(row5, text="稳定帧数").pack(side=tk.LEFT)
        ttk.Entry(row5, textvariable=self.var_stable_frames_required, width=7).pack(
            side=tk.LEFT, padx=(6, 0)
        )

        btns = ttk.Frame(parent)
        btns.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(btns, text="启动实时批阅", command=self.start_camera).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btns, text="停止", command=self.stop_camera).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="识别本地答题卡图片", command=self.grade_image_file).pack(
            side=tk.LEFT, padx=(0, 6)
        )

        score_box = ttk.Frame(parent)
        score_box.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            score_box,
            textvariable=self.var_total_score,
            font=("Microsoft YaHei UI", 20, "bold"),
            foreground="#b91c1c",
        ).pack(side=tk.LEFT)
        ttk.Label(score_box, textvariable=self.var_last_time, foreground="#334155").pack(
            side=tk.LEFT, padx=(16, 0)
        )

        preview_frame = ttk.LabelFrame(parent, text="摄像头预览", padding=6)
        preview_frame.pack(fill=tk.X, pady=(0, 8))
        self.preview_label = ttk.Label(preview_frame, text="未启动")
        self.preview_label.pack(fill=tk.X)

        detail = ttk.LabelFrame(parent, text="识别详情", padding=6)
        detail.pack(fill=tk.BOTH, expand=True)
        self.result_text = tk.Text(detail, height=16, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self.result_text.insert(
            tk.END,
            "这里会显示每道题识别结果。\n"
            "提示：手机当摄像头时，可使用 DroidCam / Iriun / IP Webcam 等工具提供视频流。\n"
            "若画面模糊/反光/外框不完整，本帧会被跳过，不会计分。\n"
            "启用“连续稳定判分”后，会在连续稳定帧达到阈值时才出分。\n"
            "若“对齐低置信度不计分”开启，对齐可疑时会直接跳过。\n"
            "可通过“手动四点外框/标记点模式”增强定位稳定性。\n",
        )
        self.result_text.configure(state=tk.DISABLED)

    @staticmethod
    def _imread_unicode(path: str) -> Optional[np.ndarray]:
        try:
            data = np.fromfile(path, dtype=np.uint8)
            img = cv2.imdecode(data, cv2.IMREAD_COLOR)
            return img
        except Exception:
            return None

    @staticmethod
    def _order_quad_points(pts: np.ndarray) -> np.ndarray:
        pts = np.asarray(pts, dtype=np.float32).reshape(4, 2)
        s = pts.sum(axis=1)
        d = np.diff(pts, axis=1).reshape(-1)
        ordered = np.zeros((4, 2), dtype=np.float32)
        ordered[0] = pts[np.argmin(s)]  # tl
        ordered[2] = pts[np.argmax(s)]  # br
        ordered[1] = pts[np.argmin(d)]  # tr
        ordered[3] = pts[np.argmax(d)]  # bl
        return ordered

    def _get_full_image_quad(self) -> Optional[np.ndarray]:
        if self.template_bgr is None:
            return None
        h, w = self.template_bgr.shape[:2]
        return np.array(
            [[0, 0], [w - 1, 0], [w - 1, h - 1], [0, h - 1]],
            dtype=np.float32,
        )

    @staticmethod
    def _quad_size(quad: np.ndarray) -> Tuple[float, float]:
        q = np.asarray(quad, dtype=np.float32).reshape(4, 2)
        w1 = float(np.linalg.norm(q[1] - q[0]))
        w2 = float(np.linalg.norm(q[2] - q[3]))
        h1 = float(np.linalg.norm(q[3] - q[0]))
        h2 = float(np.linalg.norm(q[2] - q[1]))
        return (w1 + w2) / 2.0, (h1 + h2) / 2.0

    @staticmethod
    def _quad_touches_border(quad: np.ndarray, width: int, height: int, margin: int = 10) -> bool:
        q = np.asarray(quad, dtype=np.float32).reshape(4, 2)
        for x, y in q:
            if x <= margin or y <= margin or x >= (width - 1 - margin) or y >= (height - 1 - margin):
                return True
        return False

    def _find_page_quad(self, image_bgr: np.ndarray, allow_partial: bool = False) -> Optional[np.ndarray]:
        h, w = image_bgr.shape[:2]
        area_img = float(h * w)
        gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)

        edge = cv2.Canny(blur, 40, 140)
        kernel = np.ones((5, 5), np.uint8)
        edge = cv2.dilate(edge, kernel, iterations=1)
        edge = cv2.erode(edge, kernel, iterations=1)

        contours, _ = cv2.findContours(edge, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        expected_ratio = float(w) / max(1.0, float(h))
        if self.template_bgr is not None:
            th, tw = self.template_bgr.shape[:2]
            expected_ratio = float(tw) / max(1.0, float(th))
        if isinstance(self.template_anchor_quad, np.ndarray):
            twq, thq = self._quad_size(self.template_anchor_quad)
            if thq > 1e-6:
                expected_ratio = twq / thq
        expected_norm = max(expected_ratio, 1.0 / max(expected_ratio, 1e-6))

        best_quad = None
        best_area = 0.0
        for cnt in contours:
            peri = cv2.arcLength(cnt, True)
            if peri < 240:
                continue
            approx = cv2.approxPolyDP(cnt, 0.02 * peri, True)
            if len(approx) != 4:
                continue
            if not cv2.isContourConvex(approx):
                continue
            area = abs(cv2.contourArea(approx))
            if area < area_img * 0.20:
                continue
            quad = self._order_quad_points(approx.reshape(4, 2))
            if not allow_partial and self._quad_touches_border(quad, w, h):
                continue
            qw, qh = self._quad_size(quad)
            if qw < 20 or qh < 20:
                continue
            ratio = qw / max(qh, 1e-6)
            ratio_norm = max(ratio, 1.0 / max(ratio, 1e-6))
            if not (expected_norm * 0.55 <= ratio_norm <= expected_norm * 1.85):
                continue
            if area > best_area:
                best_area = area
                best_quad = quad

        if best_quad is not None:
            return best_quad

        if not contours:
            return None

        if not allow_partial:
            return None

        largest = max(contours, key=cv2.contourArea)
        if cv2.contourArea(largest) < area_img * 0.15:
            return None
        rect = cv2.minAreaRect(largest)
        box = cv2.boxPoints(rect)
        quad = self._order_quad_points(box)
        return quad

    def _set_template_anchor(self, quad: Optional[np.ndarray]) -> None:
        if quad is None:
            self.template_anchor_quad = None
            return
        self.template_anchor_quad = self._order_quad_points(quad)

    def _question_center(self, q: QuestionConfig) -> Tuple[float, float]:
        x, y, w, h = q.rect
        return float(x + w / 2.0), float(y + h / 2.0)

    def _is_point_inside_template_anchor(self, pt: Tuple[float, float]) -> bool:
        if not isinstance(self.template_anchor_quad, np.ndarray):
            return True
        polygon = self.template_anchor_quad.astype(np.float32).reshape(-1, 1, 2)
        return cv2.pointPolygonTest(polygon, pt, False) >= 0

    def _questions_outside_anchor(self, limit: int = 8) -> List[int]:
        outside: List[int] = []
        for q in self.questions:
            if not self._is_point_inside_template_anchor(self._question_center(q)):
                outside.append(q.qid)
                if len(outside) >= limit:
                    break
        return outside

    def _try_set_anchor_with_validation(self, quad: np.ndarray) -> Tuple[bool, List[int]]:
        old = None if self.template_anchor_quad is None else self.template_anchor_quad.copy()
        self._set_template_anchor(quad)
        outside = self._questions_outside_anchor()
        if outside:
            self.template_anchor_quad = old
            return False, outside
        return True, []

    def clear_template_markers(self) -> None:
        self.template_markers.clear()
        self._render_canvas()
        self.var_status.set("已清空全部标记点。")

    def detect_template_anchor(self) -> None:
        if self.template_bgr is None:
            messagebox.showinfo("提示", "请先上传模板图片。")
            return
        quad = self._find_page_quad(self.template_bgr, allow_partial=True)
        if quad is None:
            self.var_status.set("未识别到定位外框，已保留当前设置。")
            return
        ok, outside = self._try_set_anchor_with_validation(quad)
        if not ok:
            q_text = ",".join(str(x) for x in outside)
            self.var_status.set(f"定位外框更新被拒绝：题目 {q_text} 的中心点将落在外框外。")
            return
        self._render_canvas()
        self.var_status.set("定位外框识别成功。")

    def use_full_image_anchor(self) -> None:
        quad = self._get_full_image_quad()
        if quad is None:
            return
        ok, outside = self._try_set_anchor_with_validation(quad)
        if not ok:
            q_text = ",".join(str(x) for x in outside)
            self.var_status.set(f"整图外框设置失败：题目 {q_text} 超出模板边界。")
            return
        self._render_canvas()
        self.var_status.set("已设置为整图外框。")

    def load_template_image(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择答题卡模板图片",
            filetypes=[
                ("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.webp"),
                ("All Files", "*.*"),
            ],
        )
        if not file_path:
            return
        img = self._imread_unicode(file_path)
        if img is None:
            messagebox.showerror("读取失败", "模板图片读取失败，请换一张图片试试。")
            return

        self.template_path = file_path
        self.template_bgr = img
        self.template_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        self._build_template_feature()
        found_anchor = self._find_page_quad(img, allow_partial=True)
        self._set_template_anchor(found_anchor)
        if found_anchor is None:
            self._set_template_anchor(self._get_full_image_quad())
        self.reset_view(render=False)
        self._render_canvas()

        if found_anchor is None:
            self.var_status.set("模板已加载。未识别到外框，已改用整图外框。")
        else:
            self.var_status.set("模板已加载并识别到定位外框。")

    def _build_template_feature(self) -> None:
        if self.template_gray is None:
            self.template_kp, self.template_des = None, None
            self.template_edge = None
            self.template_binary = None
            return
        orb = cv2.ORB_create(nfeatures=1800)
        self.template_kp, self.template_des = orb.detectAndCompute(self.template_gray, None)
        self.template_edge = cv2.Canny(self.template_gray, 60, 180)
        if self.template_bgr is not None:
            self.template_binary = self._prepare_binary(self.template_bgr)
        else:
            self.template_binary = None

    def _update_zoom_text(self) -> None:
        self.var_zoom.set(f"{int(round(self.view_zoom * 100))}%")

    def reset_view(self, render: bool = True) -> None:
        self.view_zoom = 1.0
        self.view_pan_x = 0
        self.view_pan_y = 0
        self._update_zoom_text()
        if render:
            self._render_canvas()

    def change_zoom(self, factor: float, anchor_canvas: Optional[Tuple[int, int]] = None) -> None:
        if self.template_bgr is None:
            return
        old_zoom = self.view_zoom
        new_zoom = max(0.4, min(6.0, old_zoom * factor))
        if abs(new_zoom - old_zoom) < 1e-6:
            return

        if anchor_canvas is not None:
            ax, ay = anchor_canvas
            ix, iy = self._canvas_to_image(ax, ay)
            self.view_zoom = new_zoom
            self._update_zoom_text()
            h, w = self.template_bgr.shape[:2]
            cw = max(1, self.canvas.winfo_width())
            ch = max(1, self.canvas.winfo_height())
            fit_scale = min(cw / w, ch / h)
            scale = fit_scale * self.view_zoom
            show_w = max(1, int(w * scale))
            show_h = max(1, int(h * scale))
            base_ox = (cw - show_w) // 2
            base_oy = (ch - show_h) // 2
            self.view_pan_x = int(ax - ix * scale - base_ox)
            self.view_pan_y = int(ay - iy * scale - base_oy)
        else:
            self.view_zoom = new_zoom
            self._update_zoom_text()

        self._render_canvas()

    def _on_canvas_mousewheel(self, event, wheel_dir: Optional[int] = None) -> None:
        direction = wheel_dir
        if direction is None:
            delta = getattr(event, "delta", 0)
            if delta == 0:
                return
            direction = 1 if delta > 0 else -1
        factor = 1.12 if direction > 0 else 0.89
        self.change_zoom(factor, anchor_canvas=(event.x, event.y))

    def _on_pan_start(self, event) -> None:
        if self.template_bgr is None:
            return
        self._panning = True
        self._pan_start = (event.x, event.y)
        self._pan_origin = (self.view_pan_x, self.view_pan_y)

    def _on_pan_move(self, event) -> None:
        if not self._panning:
            return
        dx = event.x - self._pan_start[0]
        dy = event.y - self._pan_start[1]
        self.view_pan_x = self._pan_origin[0] + dx
        self.view_pan_y = self._pan_origin[1] + dy
        self._render_canvas()

    def _on_pan_end(self, _event) -> None:
        self._panning = False

    def _on_canvas_resize(self, _event) -> None:
        try:
            self._render_canvas()
        except Exception as e:
            self.var_status.set(f"画布刷新异常：{e}")

    def _render_canvas(self) -> None:
        self.canvas.delete("all")
        if self.template_bgr is None:
            self.canvas.create_text(
                20,
                20,
                anchor="nw",
                fill="#9ca3af",
                text="请先上传答题卡模板图片",
                font=("Microsoft YaHei UI", 12),
            )
            return

        h, w = self.template_bgr.shape[:2]
        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        fit_scale = min(cw / w, ch / h)
        scale = fit_scale * self.view_zoom
        show_w = max(1, int(w * scale))
        show_h = max(1, int(h * scale))
        base_ox = (cw - show_w) // 2
        base_oy = (ch - show_h) // 2

        if show_w <= cw:
            self.view_pan_x = 0
        else:
            min_pan_x = (cw - show_w) - base_ox
            max_pan_x = -base_ox
            self.view_pan_x = max(min_pan_x, min(max_pan_x, self.view_pan_x))
        if show_h <= ch:
            self.view_pan_y = 0
        else:
            min_pan_y = (ch - show_h) - base_oy
            max_pan_y = -base_oy
            self.view_pan_y = max(min_pan_y, min(max_pan_y, self.view_pan_y))

        ox = base_ox + self.view_pan_x
        oy = base_oy + self.view_pan_y

        self.canvas_scale = scale
        self.canvas_offset_x = ox
        self.canvas_offset_y = oy

        rgb = cv2.cvtColor(self.template_bgr, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(rgb).resize((show_w, show_h), PIL_RESAMPLE_LANCZOS)
        self.canvas_photo = ImageTk.PhotoImage(pil)
        self.canvas.create_image(ox, oy, image=self.canvas_photo, anchor="nw")

        if isinstance(self.template_anchor_quad, np.ndarray):
            pts = [self._image_to_canvas(int(p[0]), int(p[1])) for p in self.template_anchor_quad]
            for i in range(4):
                p1 = pts[i]
                p2 = pts[(i + 1) % 4]
                self.canvas.create_line(
                    p1[0], p1[1], p2[0], p2[1], fill="#38bdf8", width=2, dash=(5, 3)
                )

        for i, (mx, my) in enumerate(self.template_markers):
            cx, cy = self._image_to_canvas(mx, my)
            self.canvas.create_oval(cx - 5, cy - 5, cx + 5, cy + 5, fill="#22d3ee", outline="#083344")
            self.canvas.create_text(
                cx + 8,
                cy - 8,
                anchor="nw",
                fill="#06b6d4",
                text=f"M{i+1}",
                font=("Microsoft YaHei UI", 10, "bold"),
            )

        if self._manual_anchor_points:
            cpts = [self._image_to_canvas(x, y) for x, y in self._manual_anchor_points]
            for i, (cx, cy) in enumerate(cpts):
                self.canvas.create_oval(cx - 6, cy - 6, cx + 6, cy + 6, fill="#f97316", outline="#7c2d12")
                self.canvas.create_text(
                    cx + 8,
                    cy - 10,
                    anchor="nw",
                    fill="#fb923c",
                    text=f"P{i+1}",
                    font=("Microsoft YaHei UI", 10, "bold"),
                )
            if len(cpts) >= 2:
                for i in range(len(cpts) - 1):
                    self.canvas.create_line(
                        cpts[i][0], cpts[i][1], cpts[i + 1][0], cpts[i + 1][1], fill="#f97316", width=2
                    )

        selected_idx = self._get_selected_index()
        for idx, q in enumerate(self.questions):
            x, y, rw, rh = q.rect
            c1 = self._image_to_canvas(x, y)
            c2 = self._image_to_canvas(x + rw, y + rh)
            color = "#16a34a" if q.qtype == "single" else "#f59e0b"
            width = 3 if selected_idx == idx else 2
            if selected_idx == idx:
                color = "#ef4444"
            self.canvas.create_rectangle(c1[0], c1[1], c2[0], c2[1], outline=color, width=width)
            self.canvas.create_text(
                c1[0] + 4,
                c1[1] + 4,
                anchor="nw",
                fill=color,
                text=f"{q.qid}",
                font=("Microsoft YaHei UI", 11, "bold"),
            )
            if selected_idx == idx:
                handle_size = 6
                self.canvas.create_rectangle(
                    c2[0] - handle_size,
                    c2[1] - handle_size,
                    c2[0] + handle_size,
                    c2[1] + handle_size,
                    fill="#ef4444",
                    outline="#ffffff",
                    width=1,
                )

    def _canvas_to_image(self, cx: int, cy: int) -> Tuple[int, int]:
        if self.template_bgr is None:
            return 0, 0
        h, w = self.template_bgr.shape[:2]
        ix = int((cx - self.canvas_offset_x) / self.canvas_scale)
        iy = int((cy - self.canvas_offset_y) / self.canvas_scale)
        ix = max(0, min(w - 1, ix))
        iy = max(0, min(h - 1, iy))
        return ix, iy

    def _image_to_canvas(self, ix: int, iy: int) -> Tuple[int, int]:
        cx = int(ix * self.canvas_scale + self.canvas_offset_x)
        cy = int(iy * self.canvas_scale + self.canvas_offset_y)
        return cx, cy

    def _select_tree_by_index(self, idx: int) -> None:
        if self.tree is None:
            return
        for iid, i in self.tree_item_to_index.items():
            if i == idx:
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                break

    def _find_question_at_image_point(self, ix: int, iy: int) -> Optional[int]:
        hit: Optional[int] = None
        hit_area = 0
        for i, q in enumerate(self.questions):
            x, y, w, h = q.rect
            if x <= ix <= x + w and y <= iy <= y + h:
                area = w * h
                if hit is None or area < hit_area:
                    hit = i
                    hit_area = area
        return hit

    def _sync_rect_vars(self, q: QuestionConfig) -> None:
        x, y, w, h = q.rect
        self.var_rect_x.set(str(x))
        self.var_rect_y.set(str(y))
        self.var_rect_w.set(str(w))
        self.var_rect_h.set(str(h))

    def _clamp_rect_to_template(self, rect: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
        if self.template_bgr is None:
            return rect
        h_img, w_img = self.template_bgr.shape[:2]
        x, y, w, h = rect
        w = max(6, w)
        h = max(6, h)
        x = max(0, min(w_img - 2, x))
        y = max(0, min(h_img - 2, y))
        w = max(6, min(w, w_img - x))
        h = max(6, min(h, h_img - y))
        return x, y, w, h

    def _on_canvas_press(self, event) -> None:
        if self.template_bgr is None:
            return
        ix, iy = self._canvas_to_image(event.x, event.y)

        if self.var_manual_anchor_mode.get():
            self._edit_mode = ""
            self._edit_q_idx = None
            self._manual_anchor_points.append((ix, iy))
            if len(self._manual_anchor_points) >= 4:
                quad = np.array(self._manual_anchor_points[:4], dtype=np.float32)
                ok, outside = self._try_set_anchor_with_validation(quad)
                if ok:
                    self.var_status.set("手动四点外框已更新。")
                else:
                    q_text = ",".join(str(x) for x in outside)
                    self.var_status.set(f"手动外框被拒绝：题目 {q_text} 将落在外框外。")
                self._manual_anchor_points.clear()
                self.var_manual_anchor_mode.set(False)
            else:
                self.var_status.set(
                    f"手动四点外框：已记录第 {len(self._manual_anchor_points)} 点，还需 {4-len(self._manual_anchor_points)} 点。"
                )
            self._render_canvas()
            return

        if self.var_marker_mode.get():
            self._edit_mode = ""
            self._edit_q_idx = None
            duplicated = False
            for mx, my in self.template_markers:
                if abs(mx - ix) <= 5 and abs(my - iy) <= 5:
                    duplicated = True
                    break
            if not duplicated:
                self.template_markers.append((ix, iy))
                self.var_status.set(f"已添加标记点 M{len(self.template_markers)}。")
            else:
                self.var_status.set("该位置附近已有标记点。")
            self._render_canvas()
            return

        if self.var_draw_mode.get():
            self._drawing = True
            self._draw_start = (event.x, event.y)
            if self._draw_preview_id is not None:
                self.canvas.delete(self._draw_preview_id)
                self._draw_preview_id = None
            return

        hit_idx = self._find_question_at_image_point(ix, iy)
        if hit_idx is None:
            self._edit_mode = ""
            self._edit_q_idx = None
            return

        self._select_tree_by_index(hit_idx)
        self._on_tree_select(None)

        q = self.questions[hit_idx]
        x, y, w, h = q.rect
        on_resize_handle = abs(ix - (x + w)) <= 14 and abs(iy - (y + h)) <= 14

        self._edit_mode = "resize" if on_resize_handle else "move"
        self._edit_q_idx = hit_idx
        self._edit_start_img = (ix, iy)
        self._edit_start_rect = q.rect

    def _on_canvas_drag(self, event) -> None:
        if self._drawing:
            x0, y0 = self._draw_start
            if self._draw_preview_id is not None:
                self.canvas.delete(self._draw_preview_id)
            self._draw_preview_id = self.canvas.create_rectangle(
                x0, y0, event.x, event.y, outline="#38bdf8", width=2, dash=(4, 3)
            )
            return

        if self._edit_q_idx is None or self._edit_mode not in ("move", "resize"):
            return
        ix, iy = self._canvas_to_image(event.x, event.y)
        sx, sy = self._edit_start_img
        x0, y0, w0, h0 = self._edit_start_rect
        dx, dy = ix - sx, iy - sy

        if self._edit_mode == "move":
            new_rect = (x0 + dx, y0 + dy, w0, h0)
        else:
            new_rect = (x0, y0, w0 + dx, h0 + dy)

        self.questions[self._edit_q_idx].rect = self._clamp_rect_to_template(new_rect)
        self._sync_rect_vars(self.questions[self._edit_q_idx])
        self._render_canvas()

    def _on_canvas_release(self, event) -> None:
        if self._drawing:
            self._drawing = False
            if self._draw_preview_id is not None:
                self.canvas.delete(self._draw_preview_id)
                self._draw_preview_id = None

            x0, y0 = self._draw_start
            x1, y1 = event.x, event.y
            if abs(x1 - x0) < 8 or abs(y1 - y0) < 8:
                return

            ix0, iy0 = self._canvas_to_image(min(x0, x1), min(y0, y1))
            ix1, iy1 = self._canvas_to_image(max(x0, x1), max(y0, y1))
            rw = max(2, ix1 - ix0)
            rh = max(2, iy1 - iy0)

            next_qid = 1
            if self.questions:
                next_qid = max(q.qid for q in self.questions) + 1
            q = QuestionConfig(qid=next_qid, rect=(ix0, iy0, rw, rh))
            self.questions.append(q)
            self._refresh_question_tree(select_new=True, select_idx=len(self.questions) - 1)
            self._sync_rect_vars(q)
            self._render_canvas()
            self.var_status.set(f"已新增第 {next_qid} 题区域，请在右侧设置题型和标准答案。")
            return

        if self._edit_q_idx is not None and self._edit_mode in ("move", "resize"):
            idx = self._edit_q_idx
            self._refresh_question_tree(select_new=True, select_idx=idx)
            self._render_canvas()
            self.var_status.set("题目框已更新。")
        self._edit_mode = ""
        self._edit_q_idx = None

    def _refresh_question_tree(self, select_new: bool = False, select_idx: int = -1) -> None:
        if self.tree is None:
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_item_to_index.clear()

        indexed = list(enumerate(self.questions))
        indexed.sort(key=lambda x: x[1].qid)
        for idx, q in indexed:
            answer = ",".join(q.correct) if q.correct else "-"
            qtype_cn = "单选" if q.qtype == "single" else "多选"
            rect_text = f"{q.rect[0]},{q.rect[1]},{q.rect[2]},{q.rect[3]}"
            iid = self.tree.insert(
                "",
                tk.END,
                values=(q.qid, qtype_cn, q.option_count, answer, rect_text),
            )
            self.tree_item_to_index[iid] = idx

        if select_new and 0 <= select_idx < len(self.questions):
            for iid, idx in self.tree_item_to_index.items():
                if idx == select_idx:
                    self.tree.selection_set(iid)
                    self.tree.focus(iid)
                    break

    def _on_tree_select(self, _event) -> None:
        idx = self._get_selected_index()
        if idx is None:
            return
        q = self.questions[idx]
        self.var_qid.set(str(q.qid))
        self.var_qtype.set(q.qtype)
        self.var_option_count.set(str(q.option_count))
        self.var_correct.set(",".join(q.correct))
        self._sync_rect_vars(q)
        self._render_canvas()

    def _get_selected_index(self) -> Optional[int]:
        if self.tree is None:
            return None
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree_item_to_index.get(sel[0])

    def save_selected_question(self) -> None:
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("提示", "请先在题目列表中选择一题。")
            return

        try:
            qid = int(self.var_qid.get().strip())
            option_count = int(self.var_option_count.get().strip())
        except ValueError:
            messagebox.showerror("输入错误", "题号和选项数必须是整数。")
            return

        if qid <= 0:
            messagebox.showerror("输入错误", "题号必须大于 0。")
            return
        if option_count < 2 or option_count > 8:
            messagebox.showerror("输入错误", "选项数建议在 2~8 之间。")
            return

        for i, other in enumerate(self.questions):
            if i != idx and other.qid == qid:
                messagebox.showerror("题号冲突", f"题号 {qid} 已存在，请换一个。")
                return

        raw = self.var_correct.get().replace("，", ",").strip()
        correct: List[str] = []
        if raw:
            for part in raw.split(","):
                part = part.strip().upper()
                if part and part in LETTERS[:option_count] and part not in correct:
                    correct.append(part)

        q = self.questions[idx]
        q.qid = qid
        q.qtype = "multiple" if self.var_qtype.get() == "multiple" else "single"
        q.option_count = option_count
        q.correct = correct
        q.normalize()

        self._refresh_question_tree(select_new=True, select_idx=idx)
        self._render_canvas()
        self.var_status.set(f"第 {q.qid} 题设置已保存。")

    def apply_selected_rect(self) -> None:
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("提示", "请先选择题目。")
            return
        try:
            x = int(self.var_rect_x.get().strip())
            y = int(self.var_rect_y.get().strip())
            w = int(self.var_rect_w.get().strip())
            h = int(self.var_rect_h.get().strip())
        except ValueError:
            messagebox.showerror("输入错误", "坐标必须是整数。")
            return
        rect = self._clamp_rect_to_template((x, y, w, h))
        self.questions[idx].rect = rect
        self._sync_rect_vars(self.questions[idx])
        self._refresh_question_tree(select_new=True, select_idx=idx)
        self._render_canvas()
        self.var_status.set("已应用坐标。")

    def nudge_selected_rect(self, dx_dir: int, dy_dir: int) -> None:
        idx = self._get_selected_index()
        if idx is None:
            return
        try:
            step = int(self.var_nudge_step.get().strip())
        except Exception:
            step = 3
        step = max(1, min(step, 100))
        q = self.questions[idx]
        x, y, w, h = q.rect
        rect = self._clamp_rect_to_template((x + dx_dir * step, y + dy_dir * step, w, h))
        q.rect = rect
        self._sync_rect_vars(q)
        self._refresh_question_tree(select_new=True, select_idx=idx)
        self._render_canvas()

    def delete_selected_question(self) -> None:
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("提示", "请先选择要删除的题目。")
            return
        qid = self.questions[idx].qid
        del self.questions[idx]
        self._refresh_question_tree()
        self._render_canvas()
        self.var_rect_x.set("")
        self.var_rect_y.set("")
        self.var_rect_w.set("")
        self.var_rect_h.set("")
        self.var_status.set(f"已删除第 {qid} 题。")

    def sort_questions(self) -> None:
        self.questions.sort(key=lambda x: x.qid)
        self._refresh_question_tree()
        self._render_canvas()
        self.var_status.set("已按题号排序。")

    def clear_questions(self) -> None:
        if not self.questions:
            return
        ok = messagebox.askyesno("确认", "确定清空全部题目吗？")
        if not ok:
            return
        self.questions.clear()
        self._refresh_question_tree()
        self._render_canvas()
        self.var_rect_x.set("")
        self.var_rect_y.set("")
        self.var_rect_w.set("")
        self.var_rect_h.set("")
        self.var_status.set("已清空全部题目。")

    def _collect_score_rules(self) -> ScoreRules:
        def get_float(var: tk.StringVar, default: float) -> float:
            try:
                return float(var.get().strip())
            except Exception:
                return default

        return ScoreRules(
            single_correct=get_float(self.var_single_correct, 2.0),
            single_wrong=get_float(self.var_single_wrong, 0.0),
            single_blank=get_float(self.var_single_blank, 0.0),
            multi_full=get_float(self.var_multi_full, 3.0),
            multi_partial=get_float(self.var_multi_partial, 2.0),
            multi_wrong=get_float(self.var_multi_wrong, 0.0),
            multi_blank=get_float(self.var_multi_blank, 0.0),
        )

    def save_config(self) -> None:
        if self.template_bgr is None:
            messagebox.showinfo("提示", "请先上传模板图片后再保存配置。")
            return

        save_path = filedialog.asksaveasfilename(
            title="保存配置",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not save_path:
            return

        data = {
            "template_path": self.template_path,
            "template_anchor_quad": (
                self.template_anchor_quad.tolist()
                if isinstance(self.template_anchor_quad, np.ndarray)
                else None
            ),
            "template_markers": [[int(x), int(y)] for (x, y) in self.template_markers],
            "questions": [
                {
                    "qid": q.qid,
                    "rect": list(q.rect),
                    "qtype": q.qtype,
                    "option_count": q.option_count,
                    "correct": q.correct,
                }
                for q in self.questions
            ],
            "score_rules": self._collect_score_rules().__dict__,
            "detect_params": {
                "camera_source": self.var_camera_source.get().strip(),
                "interval_sec": self.var_interval_sec.get().strip(),
                "abs_fill": self.var_abs_fill.get().strip(),
                "rel_fill": self.var_rel_fill.get().strip(),
                "use_alignment": self.var_use_alignment.get(),
                "use_anchor_align": self.var_use_anchor_align.get(),
                "enable_quality_gate": self.var_enable_quality_gate.get(),
                "require_full_anchor": self.var_require_full_anchor.get(),
                "block_low_confidence": self.var_block_low_confidence.get(),
                "block_bad_alignment": self.var_block_bad_alignment.get(),
                "enable_stable_score": self.var_enable_stable_score.get(),
                "use_marker_score": self.var_use_marker_score.get(),
                "stable_frames_required": self.var_stable_frames_required.get().strip(),
                "blur_threshold": self.var_blur_threshold.get().strip(),
                "glare_threshold": self.var_glare_threshold.get().strip(),
                "single_gap_threshold": self.var_single_gap_threshold.get().strip(),
                "single_ratio_threshold": self.var_single_ratio_threshold.get().strip(),
                "align_conf_threshold": self.var_align_conf_threshold.get().strip(),
                "marker_score_threshold": self.var_marker_score_threshold.get().strip(),
                "min_card_ratio": self.var_min_card_ratio.get().strip(),
                "voice_enabled": self.var_voice_enabled.get(),
            },
        }

        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        self.var_status.set(f"配置已保存：{save_path}")

    def load_config(self) -> None:
        load_path = filedialog.askopenfilename(
            title="加载配置",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
        )
        if not load_path:
            return

        try:
            with open(load_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("读取失败", f"配置文件读取失败：{e}")
            return

        template_path = data.get("template_path", "")
        if template_path and os.path.exists(template_path):
            img = self._imread_unicode(template_path)
            if img is not None:
                self.template_path = template_path
                self.template_bgr = img
                self.template_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                self._build_template_feature()
                self._set_template_anchor(self._find_page_quad(img, allow_partial=True))
        elif self.template_bgr is None:
            messagebox.showwarning(
                "模板缺失",
                "配置里的模板图片路径不存在，请重新上传模板图。",
            )

        anchor_raw = data.get("template_anchor_quad")
        if anchor_raw is not None:
            try:
                anchor_arr = np.array(anchor_raw, dtype=np.float32).reshape(4, 2)
                self._set_template_anchor(anchor_arr)
            except Exception:
                pass
        if self.template_anchor_quad is None:
            self._set_template_anchor(self._get_full_image_quad())

        self.template_markers.clear()
        for item in data.get("template_markers", []):
            try:
                x, y = int(item[0]), int(item[1])
                self.template_markers.append((x, y))
            except Exception:
                continue

        self.questions.clear()
        for raw in data.get("questions", []):
            try:
                q = QuestionConfig(
                    qid=int(raw.get("qid", 0)),
                    rect=tuple(int(v) for v in raw.get("rect", [0, 0, 10, 10])),
                    qtype=str(raw.get("qtype", "single")),
                    option_count=int(raw.get("option_count", 4)),
                    correct=list(raw.get("correct", [])),
                )
                q.normalize()
                if q.qid > 0:
                    self.questions.append(q)
            except Exception:
                continue

        outside = self._questions_outside_anchor()
        if outside:
            self._set_template_anchor(self._get_full_image_quad())
            outside_after = self._questions_outside_anchor()
            if outside_after:
                q_text = ",".join(str(x) for x in outside_after)
                messagebox.showwarning(
                    "外框异常",
                    f"配置中的定位外框与题目框冲突（题号 {q_text}），请重新标定定位外框。",
                )

        rules = data.get("score_rules", {})
        self.var_single_correct.set(str(rules.get("single_correct", 2)))
        self.var_single_wrong.set(str(rules.get("single_wrong", 0)))
        self.var_single_blank.set(str(rules.get("single_blank", 0)))
        self.var_multi_full.set(str(rules.get("multi_full", 3)))
        self.var_multi_partial.set(str(rules.get("multi_partial", 2)))
        self.var_multi_wrong.set(str(rules.get("multi_wrong", 0)))
        self.var_multi_blank.set(str(rules.get("multi_blank", 0)))

        params = data.get("detect_params", {})
        self.var_camera_source.set(str(params.get("camera_source", "0")))
        self.var_interval_sec.set(str(params.get("interval_sec", "2")))
        self.var_abs_fill.set(str(params.get("abs_fill", "0.12")))
        self.var_rel_fill.set(str(params.get("rel_fill", "0.65")))
        self.var_use_alignment.set(bool(params.get("use_alignment", True)))
        self.var_use_anchor_align.set(bool(params.get("use_anchor_align", True)))
        self.var_enable_quality_gate.set(bool(params.get("enable_quality_gate", True)))
        self.var_require_full_anchor.set(bool(params.get("require_full_anchor", True)))
        self.var_block_low_confidence.set(bool(params.get("block_low_confidence", True)))
        self.var_block_bad_alignment.set(bool(params.get("block_bad_alignment", True)))
        self.var_enable_stable_score.set(bool(params.get("enable_stable_score", True)))
        self.var_use_marker_score.set(bool(params.get("use_marker_score", True)))
        self.var_stable_frames_required.set(str(params.get("stable_frames_required", "3")))
        self.var_blur_threshold.set(str(params.get("blur_threshold", "85")))
        self.var_glare_threshold.set(str(params.get("glare_threshold", "0.18")))
        self.var_single_gap_threshold.set(str(params.get("single_gap_threshold", "0.045")))
        self.var_single_ratio_threshold.set(str(params.get("single_ratio_threshold", "1.25")))
        self.var_align_conf_threshold.set(str(params.get("align_conf_threshold", "0.08")))
        self.var_marker_score_threshold.set(str(params.get("marker_score_threshold", "0.45")))
        self.var_min_card_ratio.set(str(params.get("min_card_ratio", "0.20")))
        self.var_voice_enabled.set(bool(params.get("voice_enabled", True)))

        self.reset_view(render=False)
        self._refresh_question_tree()
        self._render_canvas()
        self.var_status.set(f"配置已加载：{load_path}")

    def _parse_camera_source(self):
        src = self.var_camera_source.get().strip()
        if src.isdigit():
            return int(src)
        return src

    @staticmethod
    def _clamp(v: float, low: float, high: float) -> float:
        return max(low, min(high, v))

    def _parse_detect_params(self) -> Tuple[float, float, float]:
        try:
            interval = float(self.var_interval_sec.get().strip())
        except Exception:
            interval = 2.0
        try:
            abs_fill = float(self.var_abs_fill.get().strip())
        except Exception:
            abs_fill = 0.12
        try:
            rel_fill = float(self.var_rel_fill.get().strip())
        except Exception:
            rel_fill = 0.65

        interval = self._clamp(interval, 0.2, 30.0)
        abs_fill = self._clamp(abs_fill, 0.01, 0.95)
        rel_fill = self._clamp(rel_fill, 0.05, 0.99)
        return interval, abs_fill, rel_fill

    def _get_float_var(self, var: tk.StringVar, default: float, low: float, high: float) -> float:
        try:
            value = float(var.get().strip())
        except Exception:
            value = default
        return self._clamp(value, low, high)

    def _evaluate_frame_quality(
        self, frame_bgr: np.ndarray
    ) -> Tuple[bool, List[str], Optional[np.ndarray]]:
        if not self.var_enable_quality_gate.get():
            return True, [], None

        reasons: List[str] = []
        gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)

        blur_var = float(cv2.Laplacian(gray, cv2.CV_64F).var())
        blur_thr = self._get_float_var(self.var_blur_threshold, 85.0, 5.0, 2000.0)
        if blur_var < blur_thr:
            reasons.append(f"画面模糊({blur_var:.1f}<{blur_thr:g})")

        glare_ratio = float(np.mean(gray >= 245))
        glare_thr = self._get_float_var(self.var_glare_threshold, 0.18, 0.01, 0.9)
        if glare_ratio > glare_thr:
            reasons.append(f"反光过强({glare_ratio:.2f}>{glare_thr:.2f})")

        full_quad: Optional[np.ndarray] = None
        if self.var_use_anchor_align.get() and self.var_require_full_anchor.get():
            full_quad = self._find_page_quad(frame_bgr, allow_partial=False)
            if full_quad is None:
                reasons.append("未拍全定位外框")
            else:
                card_area = abs(
                    cv2.contourArea(full_quad.astype(np.float32).reshape(-1, 1, 2))
                )
                img_area = float(frame_bgr.shape[0] * frame_bgr.shape[1])
                card_ratio = card_area / img_area if img_area > 0 else 0.0
                min_ratio = self._get_float_var(self.var_min_card_ratio, 0.20, 0.05, 0.95)
                if card_ratio < min_ratio:
                    reasons.append(f"答题卡过远({card_ratio:.2f}<{min_ratio:.2f})")

        return len(reasons) == 0, reasons, full_quad

    def _align_by_anchor_quad(
        self, frame_bgr: np.ndarray, src_quad: Optional[np.ndarray] = None
    ) -> Optional[np.ndarray]:
        if self.template_bgr is None:
            return None
        if not isinstance(self.template_anchor_quad, np.ndarray):
            return None

        if src_quad is None:
            src_quad = self._find_page_quad(frame_bgr, allow_partial=False)
        if src_quad is None:
            return None

        th, tw = self.template_bgr.shape[:2]
        try:
            H = cv2.getPerspectiveTransform(
                src_quad.astype(np.float32),
                self.template_anchor_quad.astype(np.float32),
            )
            aligned = cv2.warpPerspective(frame_bgr, H, (tw, th))
            return aligned
        except Exception:
            return None

    def _align_by_feature_match(self, frame_bgr: np.ndarray) -> Optional[np.ndarray]:
        if self.template_bgr is None or self.template_des is None or self.template_kp is None:
            return None
        th, tw = self.template_bgr.shape[:2]
        gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
        orb = cv2.ORB_create(nfeatures=1800)
        kp2, des2 = orb.detectAndCompute(gray, None)
        if des2 is None:
            return None

        try:
            matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
            matches = matcher.match(self.template_des, des2)
        except Exception:
            return None

        if len(matches) < 16:
            return None

        matches = sorted(matches, key=lambda m: m.distance)
        keep = matches[: min(240, len(matches))]
        src_pts = np.float32([kp2[m.trainIdx].pt for m in keep]).reshape(-1, 1, 2)
        dst_pts = np.float32([self.template_kp[m.queryIdx].pt for m in keep]).reshape(
            -1, 1, 2
        )
        H, mask = cv2.findHomography(src_pts, dst_pts, cv2.RANSAC, 5.0)
        if H is None or mask is None or int(mask.sum()) < 12:
            return None

        return cv2.warpPerspective(frame_bgr, H, (tw, th))

    def _marker_similarity_score(self, aligned_gray: np.ndarray) -> float:
        if self.template_gray is None or not self.template_markers:
            return -1.0
        sims: List[float] = []
        half = 13
        h, w = aligned_gray.shape[:2]
        for mx, my in self.template_markers:
            x1 = mx - half
            y1 = my - half
            x2 = mx + half + 1
            y2 = my + half + 1
            if x1 < 0 or y1 < 0 or x2 > w or y2 > h:
                continue
            t_patch = self.template_gray[y1:y2, x1:x2].astype(np.float32)
            a_patch = aligned_gray[y1:y2, x1:x2].astype(np.float32)
            if t_patch.size < 30:
                continue
            t_patch -= t_patch.mean()
            a_patch -= a_patch.mean()
            denom = float(np.linalg.norm(t_patch) * np.linalg.norm(a_patch))
            if denom < 1e-6:
                continue
            corr = float((t_patch * a_patch).sum() / denom)
            sims.append((corr + 1.0) / 2.0)
        if not sims:
            return -1.0
        return float(np.mean(sims))

    def _alignment_consistency_score(self, aligned_bgr: np.ndarray) -> Tuple[float, float]:
        if self.template_gray is None:
            return 0.0, -1.0
        gray = cv2.cvtColor(aligned_bgr, cv2.COLOR_BGR2GRAY)

        edge_template = self.template_edge
        if edge_template is None:
            edge_template = cv2.Canny(self.template_gray, 60, 180)
        edge_aligned = cv2.Canny(gray, 60, 180)
        edge_union = np.logical_or(edge_template > 0, edge_aligned > 0).sum()
        if edge_union == 0:
            edge_iou = 0.0
        else:
            edge_inter = np.logical_and(edge_template > 0, edge_aligned > 0).sum()
            edge_iou = float(edge_inter) / float(edge_union)

        orb_ratio = 0.0
        if self.template_des is not None and self.template_kp is not None:
            orb = cv2.ORB_create(nfeatures=900)
            kp2, des2 = orb.detectAndCompute(gray, None)
            if des2 is not None and len(kp2) > 0:
                try:
                    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
                    matches = matcher.match(self.template_des, des2)
                    good = [m for m in matches if m.distance <= 58]
                    denom = max(40.0, float(len(self.template_kp)))
                    orb_ratio = float(len(good)) / denom
                except Exception:
                    orb_ratio = 0.0

        marker_score = -1.0
        if self.var_use_marker_score.get():
            marker_score = self._marker_similarity_score(gray)

        if marker_score >= 0.0:
            score = 0.28 * edge_iou + 0.47 * min(orb_ratio, 1.0) + 0.25 * marker_score
        else:
            score = 0.35 * edge_iou + 0.65 * min(orb_ratio, 1.0)
        return float(score), float(marker_score)

    def _align_frame_to_template(
        self, frame_bgr: np.ndarray, src_quad: Optional[np.ndarray] = None
    ) -> np.ndarray:
        if self.template_bgr is None:
            return frame_bgr
        th, tw = self.template_bgr.shape[:2]
        if not self.var_use_alignment.get():
            self.last_align_mode = "直接缩放"
            self.last_align_confidence = 0.0
            self.last_marker_score = -1.0
            return cv2.resize(frame_bgr, (tw, th))

        candidates: List[Tuple[str, np.ndarray, float, float]] = []

        if self.var_use_anchor_align.get():
            anchor_aligned = self._align_by_anchor_quad(frame_bgr, src_quad=src_quad)
            if anchor_aligned is not None:
                score, marker_score = self._alignment_consistency_score(anchor_aligned)
                candidates.append(("外框定位", anchor_aligned, score, marker_score))

        feature_aligned = self._align_by_feature_match(frame_bgr)
        if feature_aligned is not None:
            score, marker_score = self._alignment_consistency_score(feature_aligned)
            candidates.append(("特征对齐", feature_aligned, score, marker_score))

        if candidates:
            best_mode, best_aligned, best_score, best_marker_score = max(candidates, key=lambda it: it[2])
            self.last_align_mode = best_mode
            self.last_align_confidence = float(best_score)
            self.last_marker_score = float(best_marker_score)
            return best_aligned

        self.last_align_mode = "缩放兜底"
        self.last_align_confidence = 0.0
        self.last_marker_score = -1.0
        return cv2.resize(frame_bgr, (tw, th))

    @staticmethod
    def _prepare_binary(sheet_bgr: np.ndarray) -> np.ndarray:
        gray = cv2.cvtColor(sheet_bgr, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        binary = cv2.adaptiveThreshold(
            blur,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            35,
            10,
        )
        return binary

    @staticmethod
    def _split_options(roi: np.ndarray, option_count: int) -> List[np.ndarray]:
        h, w = roi.shape[:2]
        parts: List[np.ndarray] = []
        for i in range(option_count):
            x1 = int(i * w / option_count)
            x2 = int((i + 1) * w / option_count)
            seg = roi[:, x1:x2]
            sh, sw = seg.shape[:2]
            py = max(1, int(sh * 0.18))
            px = max(1, int(sw * 0.20))
            if sh - 2 * py > 2 and sw - 2 * px > 2:
                seg = seg[py : sh - py, px : sw - px]
            parts.append(seg)
        return parts

    def _extract_option_fill_ratios(self, binary: np.ndarray, q: QuestionConfig) -> List[float]:
        h, w = binary.shape[:2]
        x, y, rw, rh = q.rect
        x = max(0, min(w - 2, x))
        y = max(0, min(h - 2, y))
        rw = max(2, min(rw, w - x))
        rh = max(2, min(rh, h - y))

        roi = binary[y : y + rh, x : x + rw]
        parts = self._split_options(roi, q.option_count)
        ratios: List[float] = []
        for part in parts:
            ratio = float(np.count_nonzero(part)) / float(part.size) if part.size else 0.0
            ratios.append(ratio)
        return ratios

    def _detect_one_question(
        self, binary: np.ndarray, q: QuestionConfig, abs_fill: float, rel_fill: float
    ) -> Tuple[List[str], List[float]]:
        observed = self._extract_option_fill_ratios(binary, q)

        # 回归修复：单选保持原先稳定逻辑，避免被多选基线扣除干扰
        if q.qtype == "single":
            max_obs = max(observed) if observed else 0.0
            if max_obs < abs_fill:
                return [], observed
            best_idx = int(np.argmax(observed))
            return [LETTERS[best_idx]], observed

        baseline = [0.0 for _ in range(q.option_count)]
        if self.template_binary is not None:
            base_raw = self._extract_option_fill_ratios(self.template_binary, q)
            if len(base_raw) == q.option_count:
                baseline = base_raw

        # 关键改进：扣除模板基线，避免印刷噪声导致多选误判成 A,B,C,D
        effective = [max(0.0, r - b) for r, b in zip(observed, baseline)]
        max_eff = max(effective) if effective else 0.0
        abs_eff_thr = max(0.012, abs_fill * 0.18)
        if max_eff < abs_eff_thr:
            return [], effective

        dynamic_threshold = max(abs_eff_thr, max_eff * rel_fill)
        selected_idx = [i for i, r in enumerate(effective) if r >= dynamic_threshold]

        # 多选过选纠偏：若候选过多，收紧阈值，优先保留“强信号”选项
        if len(selected_idx) >= 3:
            strong_thr = max(abs_eff_thr * 1.35, max_eff * 0.82)
            strong = [i for i, r in enumerate(effective) if r >= strong_thr]
            if 0 < len(strong) < len(selected_idx):
                selected_idx = strong

        # 若仍全选，且对比度不足，则按“拐点”回退，防止 A/AB 误成 ABCD
        if len(selected_idx) == q.option_count and q.option_count >= 3:
            spread = max_eff - min(effective)
            spread_thr = max(0.010, abs_eff_thr * 0.60)
            if spread < spread_thr:
                order = sorted(range(q.option_count), key=lambda i: effective[i], reverse=True)
                drops = [effective[order[i]] - effective[order[i + 1]] for i in range(len(order) - 1)]
                max_drop = max(drops) if drops else 0.0
                if max_drop >= max(0.010, abs_eff_thr * 0.50):
                    k = drops.index(max_drop) + 1
                    selected_idx = order[:k]
                else:
                    selected_idx = [order[0]]

        selected = [LETTERS[i] for i in selected_idx if i < len(LETTERS)]
        return selected, effective

    @staticmethod
    def _score_question(q: QuestionConfig, selected: List[str], rules: ScoreRules) -> Tuple[float, str]:
        corr = set(q.correct)
        sel = set(selected)

        if q.qtype == "single":
            if not sel:
                return rules.single_blank, "空白"
            if len(sel) == 1 and sel == corr:
                return rules.single_correct, "正确"
            return rules.single_wrong, "错误"

        if not sel:
            return rules.multi_blank, "空白"
        if sel == corr:
            return rules.multi_full, "全对"
        if sel.issubset(corr) and len(sel) < len(corr):
            return rules.multi_partial, "漏选"
        return rules.multi_wrong, "错选/多选"

    def _grade_sheet(self, sheet_bgr: np.ndarray) -> Tuple[float, List[str], str]:
        if self.template_bgr is None:
            raise RuntimeError("模板未加载")
        if not self.questions:
            raise RuntimeError("题目区域未配置")

        _, abs_fill, rel_fill = self._parse_detect_params()
        rules = self._collect_score_rules()
        ok_quality, quality_reasons, full_quad = self._evaluate_frame_quality(sheet_bgr)
        if not ok_quality:
            raise RuntimeError("画面不达标：" + "；".join(quality_reasons))

        if self.var_use_anchor_align.get():
            outside = self._questions_outside_anchor()
            if outside:
                q_text = ",".join(str(x) for x in outside)
                raise RuntimeError(f"题目框中心不在定位外框内：{q_text}，请重新标定定位外框。")

        aligned = self._align_frame_to_template(sheet_bgr, src_quad=full_quad)
        if self.var_use_alignment.get() and self.var_block_bad_alignment.get():
            align_thr = self._get_float_var(self.var_align_conf_threshold, 0.08, 0.0, 1.0)
            if self.last_align_confidence < align_thr:
                raise RuntimeError(
                    f"定位疑似错误：对齐置信度 {self.last_align_confidence:.3f} < {align_thr:.3f}"
                )
            if self.var_use_marker_score.get() and self.template_markers and self.last_marker_score >= 0:
                marker_thr = self._get_float_var(self.var_marker_score_threshold, 0.45, 0.0, 1.0)
                if self.last_marker_score < marker_thr:
                    raise RuntimeError(
                        f"定位疑似错误：标记点匹配度 {self.last_marker_score:.3f} < {marker_thr:.3f}"
                    )
        binary = self._prepare_binary(aligned)

        total = 0.0
        detail_lines: List[str] = []
        low_conf_items: List[Tuple[int, float, float, float]] = []
        gap_thr = self._get_float_var(self.var_single_gap_threshold, 0.045, 0.0, 0.5)
        ratio_thr = self._get_float_var(self.var_single_ratio_threshold, 1.25, 1.0, 10.0)
        answer_signature_parts: List[str] = []
        for q in sorted(self.questions, key=lambda it: it.qid):
            selected, ratios = self._detect_one_question(binary, q, abs_fill, rel_fill)
            score, reason = self._score_question(q, selected, rules)
            total += score
            sel_text = ",".join(selected) if selected else "-"
            answer_signature_parts.append(f"{q.qid}:{sel_text}")
            corr_text = ",".join(q.correct) if q.correct else "未设置"
            ratio_preview = ",".join(f"{v:.2f}" for v in ratios)
            extra = ""
            if q.qtype == "single" and ratios:
                sorted_ratios = sorted(ratios, reverse=True)
                top1 = sorted_ratios[0]
                top2 = sorted_ratios[1] if len(sorted_ratios) > 1 else 0.0
                gap = top1 - top2
                ratio = top1 / max(top2, 1e-6)
                extra = f" 置信差={gap:.3f} 比值={ratio:.2f}"
                # 同时满足“差值低且比值低”才判为低置信，降低误拦截
                if selected and (gap < gap_thr and ratio < ratio_thr):
                    low_conf_items.append((q.qid, gap, top1, top2))
            detail_lines.append(
                f"第{q.qid}题 [{('单选' if q.qtype == 'single' else '多选')}] "
                f"作答={sel_text} 标准={corr_text} 得分={score:g} ({reason}) "
                f"填涂比=[{ratio_preview}]{extra}"
            )

        if low_conf_items:
            qids = ",".join(str(it[0]) for it in low_conf_items[:6])
            warning = f"单选低置信度题号：{qids}（差值阈值 {gap_thr:.3f}，比值阈值 {ratio_thr:.2f}）"
            if self.var_block_low_confidence.get():
                raise RuntimeError(warning + "，建议重拍或调整光线")
            detail_lines.insert(0, "警告：" + warning)

        signature = "|".join(answer_signature_parts)
        return total, detail_lines, signature

    def start_camera(self) -> None:
        if self.template_bgr is None:
            messagebox.showinfo("提示", "请先上传模板图片并配置题目。")
            return
        if not self.questions:
            messagebox.showinfo("提示", "请先框选并配置题目。")
            return

        self.stop_camera()
        src = self._parse_camera_source()
        cap = cv2.VideoCapture(src)
        if not cap.isOpened():
            messagebox.showerror(
                "摄像头打开失败",
                "无法打开摄像头源，请检查：\n"
                "1) 本机摄像头索引（0/1）\n"
                "2) 手机视频流 URL 是否可访问\n"
                "3) 手机和电脑是否在同一网络",
            )
            return

        self.cap = cap
        self.running = True
        self.last_eval_time = 0.0
        self.last_total_score = None
        self.last_skip_reason = ""
        self.last_align_confidence = 0.0
        self.last_marker_score = -1.0
        self._stable_signature = ""
        self._stable_count = 0
        self._stable_total = 0.0
        self._stable_lines = []
        self._last_emitted_signature = ""
        self.var_status.set("实时批阅已启动。")
        self._camera_loop()

    def stop_camera(self) -> None:
        self.running = False
        if self.after_job is not None:
            self.after_cancel(self.after_job)
            self.after_job = None
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.last_align_confidence = 0.0
        self.last_marker_score = -1.0
        self._stable_signature = ""
        self._stable_count = 0
        self._stable_total = 0.0
        self._stable_lines = []
        self._last_emitted_signature = ""
        self.preview_label.configure(text="未启动", image="")

    def _camera_loop(self) -> None:
        if not self.running or self.cap is None:
            return

        try:
            ret, frame = self.cap.read()
            if ret and frame is not None:
                self._update_preview(frame)
                interval, _, _ = self._parse_detect_params()
                now = time.time()
                if now - self.last_eval_time >= interval:
                    self.last_eval_time = now
                    self._run_grading(frame)
            else:
                self.var_status.set("读取摄像头画面失败，请检查连接。")
        except Exception as e:
            self.var_status.set(f"摄像头循环异常：{e}")

        self.after_job = self.after(30, self._camera_loop)

    def _parse_stable_frames_required(self) -> int:
        try:
            n = int(self.var_stable_frames_required.get().strip())
        except Exception:
            n = 3
        return max(1, min(10, n))

    def _update_preview(self, frame_bgr: np.ndarray) -> None:
        rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
        h, w = rgb.shape[:2]
        max_w, max_h = 560, 300
        scale = min(max_w / w, max_h / h, 1.0)
        show = Image.fromarray(rgb).resize(
            (int(w * scale), int(h * scale)), PIL_RESAMPLE_LANCZOS
        )
        self.preview_photo = ImageTk.PhotoImage(show)
        self.preview_label.configure(image=self.preview_photo, text="")

    def _emit_grading_result(self, total: float, lines: List[str]) -> None:
        try:
            self.var_total_score.set(f"总分：{total:g}")
            self.var_last_time.set(time.strftime("最近批阅：%H:%M:%S"))
            self.last_skip_reason = ""

            report = (
                f"[{time.strftime('%H:%M:%S')}] 总分：{total:g}\n"
                + "\n".join(lines)
                + "\n"
                + ("-" * 60)
                + "\n"
            )
            self._append_result(report)

            if self.var_voice_enabled.get():
                self.voice.speak(f"本次得分 {total:g} 分", enabled=True)
            self.last_total_score = total
            marker_text = (
                f"，标记匹配 {self.last_marker_score:.3f}"
                if (self.var_use_marker_score.get() and self.last_marker_score >= 0)
                else ""
            )
            self.var_status.set(
                f"批阅完成（{self.last_align_mode}，置信度 {self.last_align_confidence:.3f}{marker_text}）。"
            )
        except Exception as e:
            self.var_status.set(f"结果输出异常：{e}")

    def _run_grading(self, frame_bgr: np.ndarray) -> None:
        try:
            total, lines, signature = self._grade_sheet(frame_bgr)
        except Exception as e:
            reason = str(e)
            self.var_last_time.set(time.strftime("最近尝试：%H:%M:%S"))
            self.var_status.set(f"本帧未计分：{reason}")
            if reason != self.last_skip_reason:
                self._append_result(
                    f"[{time.strftime('%H:%M:%S')}] 跳过：{reason}\n"
                    + ("-" * 60)
                    + "\n"
                )
                self.last_skip_reason = reason
            self._stable_signature = ""
            self._stable_count = 0
            self._stable_lines = []
            return

        if self.var_enable_stable_score.get():
            required = self._parse_stable_frames_required()
            if signature == self._stable_signature:
                self._stable_count += 1
            else:
                self._stable_signature = signature
                self._stable_count = 1
            self._stable_total = total
            self._stable_lines = lines

            if self._stable_count < required:
                self.var_last_time.set(time.strftime("最近尝试：%H:%M:%S"))
                marker_text = (
                    f"，标记匹配 {self.last_marker_score:.3f}"
                    if (self.var_use_marker_score.get() and self.last_marker_score >= 0)
                    else ""
                )
                self.var_status.set(
                    f"识别中（稳定计数 {self._stable_count}/{required}，"
                    f"{self.last_align_mode}，置信度 {self.last_align_confidence:.3f}{marker_text}）"
                )
                return

            if self._last_emitted_signature == self._stable_signature:
                marker_text = (
                    f"，标记匹配 {self.last_marker_score:.3f}"
                    if (self.var_use_marker_score.get() and self.last_marker_score >= 0)
                    else ""
                )
                self.var_status.set(
                    f"识别稳定（{self.last_align_mode}，置信度 {self.last_align_confidence:.3f}{marker_text}）。"
                )
                return

            self._emit_grading_result(self._stable_total, self._stable_lines)
            self._last_emitted_signature = self._stable_signature
            return

        self._emit_grading_result(total, lines)
        self._last_emitted_signature = signature

    def grade_image_file(self) -> None:
        if self.template_bgr is None or not self.questions:
            messagebox.showinfo("提示", "请先加载模板并配置题目。")
            return
        path = filedialog.askopenfilename(
            title="选择要批阅的答题卡图片",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.webp"), ("All Files", "*.*")],
        )
        if not path:
            return
        img = self._imread_unicode(path)
        if img is None:
            messagebox.showerror("读取失败", "答题卡图片读取失败。")
            return

        try:
            total, lines, _ = self._grade_sheet(img)
        except Exception as e:
            messagebox.showerror("识别失败", str(e))
            return

        self.var_total_score.set(f"总分：{total:g}")
        self.var_last_time.set(time.strftime("最近批阅：%H:%M:%S"))
        report = (
            f"[{time.strftime('%H:%M:%S')}] 图片：{os.path.basename(path)} 总分：{total:g}\n"
            + "\n".join(lines)
            + "\n"
            + ("-" * 60)
            + "\n"
        )
        self._append_result(report)
        if self.var_voice_enabled.get():
            self.voice.speak(f"本次得分 {total:g} 分", enabled=True)
        self.var_status.set(f"已完成图片批阅：{os.path.basename(path)}")

    def _append_result(self, content: str) -> None:
        if not self.winfo_exists():
            return
        self.result_text.configure(state=tk.NORMAL)
        self.result_text.insert("1.0", content)
        self.result_text.configure(state=tk.DISABLED)

    def on_close(self) -> None:
        self.stop_camera()
        self.destroy()


def main() -> None:
    app = AnswerSheetGraderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
