#!/usr/bin/env python3
"""
快手评论抓取器 - GUI 界面
=========================
基于 tkinter 的一键操作界面，无需命令行。
双击运行 gui.py 即可。

功能：
  - 单个视频 / 批量输入（每行一个）
  - 解析分享链接（/f/xxx → photoId）
  - 可选抓取子回复（CDP 浏览器 DOM 模式）
  - Excel 文件名自动使用视频标题（≤10字）
  - 实时日志滚动
  - 进度条
  - 后台线程抓取，界面不卡顿
  - 完成后一键打开输出目录
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time
import os
import sys
import json
import socket
from urllib.parse import urlparse
from datetime import datetime

# ── Windows 控制台编码 ──
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from ks_scraper import KuaishouCommentScraper, extract_photo_id, launch_edge_cdp, sanitize_filename


class KuaishouScraperGUI:
    """快手评论抓取器 GUI"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("快手评论抓取器 v1.2")
        self.root.geometry("820x680")
        self.root.minsize(700, 580)

        # 状态
        self.scraper = None
        self.scraping = False
        self.video_list = []

        # 设置默认值
        self.cdp_var = tk.StringVar(value="http://127.0.0.1:28800")
        self.output_var = tk.StringVar(value="ks_output")
        self.cookie_var = tk.StringVar(value="ks_cookies.json")
        self.max_pages_var = tk.StringVar(value="50")
        self.delay_var = tk.StringVar(value="0.5")
        self.batch_delay_var = tk.StringVar(value="2.0")
        self.timeout_var = tk.StringVar(value="15")
        self.progress_var = tk.StringVar(value="就绪")
        self.status_var = tk.StringVar(value="等待开始...")
        self.sub_var = tk.BooleanVar(value=False)  # 子评论开关

        self._setup_ui()
        self._center_window()

        # 自动检测已保存的 cookie
        self._auto_detect_cookie()

    # ═══════════════════════════════════════
    # UI 搭建
    # ═══════════════════════════════════════

    def _setup_ui(self):
        root = self.root
        root.configure(padx=12, pady=10)

        # 样式
        style = ttk.Style()
        style.configure("Title.TLabel", font=("微软雅黑", 14, "bold"))
        style.configure("Section.TLabel", font=("微软雅黑", 10, "bold"))
        style.configure("Action.TButton", font=("微软雅黑", 10))
        style.configure("Small.TButton", font=("微软雅黑", 9))

        # ── 标题 ──
        title_frame = ttk.Frame(root)
        title_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(title_frame, text="🎬 快手评论抓取器", style="Title.TLabel").pack(side="left")
        ttk.Label(title_frame, text="v1.2", foreground="gray").pack(side="left", padx=(6, 0))

        # ── 输入区 ──
        input_frame = ttk.LabelFrame(root, text="📥 视频输入", padding=8)
        input_frame.pack(fill="both", expand=False, pady=(0, 8))

        ttk.Label(input_frame, text="批量输入（每行一个视频 URL 或 ID）：").pack(anchor="w")

        self.video_text = tk.Text(input_frame, height=5, font=("Consolas", 10), wrap="none")
        self.video_text.pack(fill="both", expand=True, pady=(4, 0))

        # 单行快捷输入
        quick_frame = ttk.Frame(input_frame)
        quick_frame.pack(fill="x", pady=(6, 0))
        ttk.Label(quick_frame, text="快捷输入：").pack(side="left")
        self.quick_entry = ttk.Entry(quick_frame, font=("Consolas", 10))
        self.quick_entry.pack(side="left", fill="x", expand=True, padx=(4, 4))
        self.quick_entry.bind("<Return>", self._on_quick_add)
        ttk.Button(quick_frame, text="添加 >>", style="Small.TButton",
                   command=self._add_video).pack(side="left")
        ttk.Button(quick_frame, text="清空", style="Small.TButton",
                   command=self._clear_videos).pack(side="left", padx=(4, 0))
        ttk.Button(quick_frame, text="🔗 解析分享链接", style="Small.TButton",
                   command=self._resolve_share_link).pack(side="left", padx=(12, 0))

        # 提示
        hint_frame = ttk.Frame(input_frame)
        hint_frame.pack(fill="x", pady=(4, 0))
        ttk.Label(hint_frame, text="💡 支持格式: 完整URL / short-video/xxx / 纯ID (如 3xqi7iru65ut3bk)  /  分享短链 /f/xxx",
                  foreground="gray", font=("微软雅黑", 8)).pack(anchor="w")

        # ── 设置区 ──
        settings_frame = ttk.LabelFrame(root, text="⚙ 抓取设置", padding=8)
        settings_frame.pack(fill="x", pady=(0, 8))

        # ── 连接状态 ──
        conn_row = ttk.Frame(settings_frame)
        conn_row.pack(fill="x", pady=(0, 6))
        self.conn_status = ttk.Label(conn_row, text="🔍 检测中...", foreground="gray")
        self.conn_status.pack(side="left")
        self.auto_btn = ttk.Button(conn_row, text="🔗 自动连接", style="Small.TButton",
                                   command=self._auto_connect)
        self.auto_btn.pack(side="left", padx=(8, 0))

        row1 = ttk.Frame(settings_frame)
        row1.pack(fill="x", pady=(0, 4))
        ttk.Label(row1, text="浏览器 CDP：").pack(side="left")
        ttk.Entry(row1, textvariable=self.cdp_var, width=30, font=("Consolas", 9)).pack(side="left", padx=(4, 16))

        ttk.Label(row1, text="输出目录：").pack(side="left")
        ttk.Entry(row1, textvariable=self.output_var, width=18, font=("Consolas", 9)).pack(side="left", padx=(4, 4))
        ttk.Button(row1, text="浏览...", style="Small.TButton",
                   command=self._browse_output).pack(side="left")

        row1b = ttk.Frame(settings_frame)
        row1b.pack(fill="x", pady=(0, 4))
        ttk.Label(row1b, text="🍪 Cookie 文件：").pack(side="left")
        ttk.Entry(row1b, textvariable=self.cookie_var, width=36, font=("Consolas", 9)).pack(side="left", padx=(4, 4))
        ttk.Button(row1b, text="💾 保存", style="Small.TButton",
                   command=self._save_cookies).pack(side="left")
        ttk.Button(row1b, text="浏览...", style="Small.TButton",
                   command=self._browse_cookie).pack(side="left", padx=(4, 0))
        self.cookie_status = ttk.Label(row1b, text="", foreground="gray")
        self.cookie_status.pack(side="left", padx=(8, 0))

        row2 = ttk.Frame(settings_frame)
        row2.pack(fill="x")
        ttk.Label(row2, text="最大翻页：").pack(side="left")
        ttk.Entry(row2, textvariable=self.max_pages_var, width=6, font=("Consolas", 9)).pack(side="left", padx=(4, 12))

        ttk.Label(row2, text="翻页间隔(秒)：").pack(side="left")
        ttk.Entry(row2, textvariable=self.delay_var, width=6, font=("Consolas", 9)).pack(side="left", padx=(4, 12))

        ttk.Label(row2, text="视频间隔(秒)：").pack(side="left")
        ttk.Entry(row2, textvariable=self.batch_delay_var, width=6, font=("Consolas", 9)).pack(side="left", padx=(4, 12))

        ttk.Label(row2, text="超时(秒)：").pack(side="left")
        ttk.Entry(row2, textvariable=self.timeout_var, width=6, font=("Consolas", 9)).pack(side="left", padx=(4, 0))

        # ── 子评论开关 ──
        sub_row = ttk.Frame(settings_frame)
        sub_row.pack(fill="x", pady=(6, 0))
        self.sub_cb = ttk.Checkbutton(
            sub_row, text="📎 包含子评论（需 CDP 浏览器，抓取较慢）",
            variable=self.sub_var,
        )
        self.sub_cb.pack(side="left")

        # ── 按钮区 ──
        btn_frame = ttk.Frame(root)
        btn_frame.pack(fill="x", pady=(0, 8))

        self.start_btn = ttk.Button(btn_frame, text="▶ 开始抓取", style="Action.TButton",
                                    command=self._start_scrape)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(btn_frame, text="■ 停止", style="Action.TButton",
                                   command=self._stop_scrape, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        ttk.Button(btn_frame, text="📂 打开输出目录", style="Action.TButton",
                   command=self._open_output).pack(side="left", padx=(8, 0))

        # 进度标签
        self.progress_label = ttk.Label(btn_frame, textvariable=self.progress_var, foreground="gray")
        self.progress_label.pack(side="right")

        # ── 进度条 ──
        self.progress_bar = ttk.Progressbar(root, mode="determinate", length=100)
        self.progress_bar.pack(fill="x", pady=(0, 4))

        # ── 日志区 ──
        log_frame = ttk.LabelFrame(root, text="📋 运行日志", padding=4)
        log_frame.pack(fill="both", expand=True)

        self.log_area = scrolledtext.ScrolledText(
            log_frame, height=12, font=("Consolas", 9),
            wrap="word", state="disabled",
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
        )
        self.log_area.pack(fill="both", expand=True)

        # 日志颜色标签
        self.log_area.tag_config("success", foreground="#4ec9b0")
        self.log_area.tag_config("error", foreground="#f44747")
        self.log_area.tag_config("info", foreground="#569cd6")
        self.log_area.tag_config("warn", foreground="#dcdcaa")
        self.log_area.tag_config("time", foreground="#808080")

        # ── 状态栏 ──
        status_bar = ttk.Frame(root)
        status_bar.pack(fill="x", pady=(4, 0))
        ttk.Label(status_bar, textvariable=self.status_var, foreground="gray").pack(side="left")

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"+{x}+{y}")

    # ═══════════════════════════════════════
    # 交互
    # ═══════════════════════════════════════

    def _log(self, text, tag=""):
        """线程安全写日志"""
        def _write():
            self.log_area.configure(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_area.insert("end", f"[{ts}] ", "time")
            self.log_area.insert("end", text + "\n", tag)
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.root.after(0, _write)

    def _on_quick_add(self, event):
        self._add_video()

    def _check_readiness(self):
        """检测 CDP / cookie 状态，更新顶部连接指示"""
        cdp = self.cdp_var.get().strip()
        cookie_path = self.cookie_var.get().strip() or "ks_cookies.json"

        cookie_ok = False
        if os.path.exists(cookie_path):
            try:
                with open(cookie_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data:
                    cookie_ok = True
                    if not self.sub_var.get():
                        self.conn_status.configure(
                            text="✅ 就绪（Cookie 文件）", foreground="#4ec9b0")
                        self.auto_btn.configure(state="disabled")
                        return True
            except Exception:
                pass

        cdp_ok = False
        try:
            parsed = urlparse(cdp)
            host = parsed.hostname or "127.0.0.1"
            s = socket.create_connection((host, parsed.port), timeout=2)
            s.close()
            cdp_ok = True
        except Exception:
            pass

        if cdp_ok:
            if self.sub_var.get():
                self.conn_status.configure(
                    text="✅ 就绪（CDP — 子评论模式）", foreground="#4ec9b0")
            else:
                self.conn_status.configure(
                    text="✅ 就绪（CDP 已连接）", foreground="#4ec9b0")
            self.auto_btn.configure(state="normal")
            return True

        if cookie_ok:
            if self.sub_var.get():
                self.conn_status.configure(
                    text="⚠ 子评论需 CDP 浏览器（Cookie 仅 API）", foreground="#dcdcaa")
            else:
                self.conn_status.configure(
                    text="✅ 就绪（Cookie 文件）", foreground="#4ec9b0")
            self.auto_btn.configure(state="normal")
            return True

        self.conn_status.configure(
            text="⚠ 未连接 — 点「自动连接」启动浏览器", foreground="#dcdcaa")
        self.auto_btn.configure(state="normal")
        return False

    def _auto_connect(self):
        """后台线程：自动启动 CDP 浏览器 + 提取 cookie"""
        self.auto_btn.configure(state="disabled")
        self.conn_status.configure(text="⏳ 启动浏览器...", foreground="#569cd6")
        self.status_var.set("自动连接中...")
        self.root.update()

        def _do():
            cdp = self.cdp_var.get().strip()
            cookie_path = self.cookie_var.get().strip() or "ks_cookies.json"
            try:
                scraper = KuaishouCommentScraper(cdp_url=cdp, cookie_file=cookie_path)
                scraper.auto_connect()
                self.root.after(0, lambda: self.conn_status.configure(
                    text="✅ 已连接！", foreground="#4ec9b0"))
                self.root.after(0, lambda: self.status_var.set("就绪"))
                self._log("✅ 自动连接成功，Cookie 已保存", "success")
                self._auto_detect_cookie()
            except Exception as e:
                self.root.after(0, lambda: self.conn_status.configure(
                    text="❌ 连接失败", foreground="#f44747"))
                self.root.after(0, lambda: self.auto_btn.configure(state="normal"))
                self.root.after(0, lambda: self.status_var.set("连接失败"))
                self._log(f"❌ 自动连接失败: {e}", "error")

        threading.Thread(target=_do, daemon=True).start()

    def _add_video(self):
        text = self.quick_entry.get().strip()
        if not text:
            return
        current = self.video_text.get("1.0", "end-1c").strip()
        if current:
            self.video_text.insert("end", "\n" + text)
        else:
            self.video_text.insert("1.0", text)
        self.quick_entry.delete(0, "end")

    def _clear_videos(self):
        self.video_text.delete("1.0", "end")

    def _resolve_share_link(self):
        """解析快手分享短链接（/f/xxx），提取视频 photoId 添加到列表"""
        import tkinter.simpledialog as sd
        url = sd.askstring(
            "解析分享链接",
            "粘贴快手分享链接：\n（如 https://www.kuaishou.com/f/xxxx）",
            parent=self.root,
        )
        if not url:
            return

        url = url.strip()
        if not url:
            return

        try:
            pid = KuaishouCommentScraper.resolve_share_link(url)
        except ValueError:
            messagebox.showwarning(
                "解析失败",
                f"无法从此链接提取视频 ID：\n{url[:80]}\n\n"
                "请确认这是快手视频链接（支持 /f/ 短链和完整视频 URL）。"
            )
            return
        except Exception as e:
            messagebox.showwarning("解析失败", f"网络错误：{e}")
            return

        # 添加到视频列表
        current = self.video_text.get("1.0", "end-1c").strip()
        existing = set(current.splitlines()) if current else set()
        if pid in existing:
            messagebox.showinfo("已存在", f"视频 {pid} 已在列表中")
            return

        if current:
            self.video_text.insert("end", "\n" + pid)
        else:
            self.video_text.insert("1.0", pid)

        self._log(f"🔗 解析分享链接 → {pid}", "info")

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_var.set(path)

    def _open_output(self):
        out = self.output_var.get().strip()
        if out:
            full = os.path.abspath(out)
            os.makedirs(full, exist_ok=True)
            os.startfile(full)

    def _auto_detect_cookie(self):
        """启动时检测连接状态"""
        path = self.cookie_var.get().strip() or "ks_cookies.json"
        if path and os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                count = len(data) if isinstance(data, dict) else 0
                self.cookie_status.configure(
                    text=f"✅ 已加载 ({count}个)", foreground="#4ec9b0"
                )
            except Exception:
                self.cookie_status.configure(text="⚠ 文件损坏", foreground="#f44747")
        self._check_readiness()

    def _save_cookies(self):
        """从浏览器提取 cookie 并保存"""
        cdp = self.cdp_var.get().strip()
        path = self.cookie_var.get().strip() or "ks_cookies.json"

        try:
            scraper = KuaishouCommentScraper(cdp_url=cdp)
            scraper.save_cookies(path)
            self.cookie_status.configure(
                text=f"✅ 已保存 ({path})", foreground="#4ec9b0"
            )
            self._log(f"💾 Cookie 已保存: {path}", "success")
        except Exception as e:
            self.cookie_status.configure(text=f"❌ 失败", foreground="#f44747")
            messagebox.showwarning(
                "保存失败",
                f"无法从浏览器提取 Cookie:\n{e}\n\n"
                f"请在浏览器中打开快手页面后重试。"
            )

    def _browse_cookie(self):
        path = filedialog.asksaveasfilename(
            title="选择 Cookie 保存位置",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")],
            initialfile="ks_cookies.json",
        )
        if path:
            self.cookie_var.set(path)
            self._auto_detect_cookie()

    def _set_ui_state(self, running: bool):
        """切换 UI 状态"""
        state = "disabled" if running else "normal"
        self.start_btn.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")
        self.video_text.configure(state="disabled" if running else "normal")
        self.quick_entry.configure(state="disabled" if running else "normal")
        self.sub_cb.configure(state="disabled" if running else "normal")

    def _update_progress(self, current, total, text=""):
        def _do():
            if total > 0:
                pct = int(current / total * 100)
                self.progress_bar["value"] = pct
                self.progress_var.set(f"进度: {current}/{total}" + (f" — {text}" if text else ""))
            else:
                self.progress_var.set(text)
        self.root.after(0, _do)

    # ═══════════════════════════════════════
    # 核心抓取逻辑（后台线程）
    # ═══════════════════════════════════════

    def _start_scrape(self):
        # 收集视频列表
        raw = self.video_text.get("1.0", "end-1c").strip()
        if not raw:
            messagebox.showwarning("提示", "请先输入至少一个视频 URL 或 ID")
            return

        lines = [l.strip() for l in raw.splitlines() if l.strip() and not l.strip().startswith("#")]
        self.video_list = []
        for line in lines:
            try:
                pid = extract_photo_id(line)
                self.video_list.append(pid)
            except ValueError:
                self._log(f"⚠ 无法解析，已跳过: {line}", "warn")

        # 去重
        seen = set()
        unique = []
        for v in self.video_list:
            if v not in seen:
                seen.add(v)
                unique.append(v)
        self.video_list = unique

        if not self.video_list:
            messagebox.showwarning("提示", "没有有效的视频 ID")
            return

        # 读取设置
        cdp = self.cdp_var.get().strip()
        output_dir = self.output_var.get().strip()
        cookie_path = self.cookie_var.get().strip() or "ks_cookies.json"
        include_sub = self.sub_var.get()

        try:
            max_pages = int(self.max_pages_var.get().strip())
        except ValueError:
            max_pages = 50
        try:
            delay = float(self.delay_var.get().strip())
        except ValueError:
            delay = 0.5
        try:
            batch_delay = float(self.batch_delay_var.get().strip())
        except ValueError:
            batch_delay = 2.0
        try:
            timeout = int(self.timeout_var.get().strip())
        except ValueError:
            timeout = 15

        # 切 UI 状态
        self.scraping = True
        self._set_ui_state(True)
        self.progress_bar["value"] = 0
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        self.status_var.set("正在抓取..." + (" (含子评论)" if include_sub else ""))

        # 启动后台线程
        thread = threading.Thread(
            target=self._scrape_thread,
            args=(cdp, output_dir, cookie_path, include_sub,
                  max_pages, delay, batch_delay, timeout),
            daemon=True,
        )
        thread.start()

    def _stop_scrape(self):
        self.scraping = False
        self._log("⏸ 正在停止...", "warn")
        self.status_var.set("已请求停止")

    def _scrape_thread(self, cdp, output_dir, cookie_path, include_sub,
                       max_pages, delay, batch_delay, timeout):
        """后台抓取线程"""
        os.makedirs(output_dir, exist_ok=True)

        # 初始化 + 自动连接
        try:
            self.scraper = KuaishouCommentScraper(
                cdp_url=cdp, cookie_file=cookie_path,
                timeout=timeout, page_delay=delay,
            )
            self.scraper.auto_connect()
        except Exception as e:
            self._log(f"❌ 连接失败: {e}", "error")
            self._finish_scrape(False)
            return

        self._log(f"🚀 开始抓取 {len(self.video_list)} 个视频", "info")
        if include_sub:
            self._log(f"   📎 子评论模式：需要 CDP 浏览器", "info")
        self._log(f"   输出: {os.path.abspath(output_dir)}")
        self._log("")

        success = 0
        fail = 0
        total_comments = 0
        total_likes = 0
        total_subs = 0
        start_time = time.time()

        for idx, pid in enumerate(self.video_list):
            if not self.scraping:
                self._log("⏸ 用户停止", "warn")
                break

            self._update_progress(idx + 1, len(self.video_list),
                                  f"当前: {pid[:16]}...")
            self._log(f"📹 [{idx+1}/{len(self.video_list)}] {pid}", "info")

            output_path = os.path.join(output_dir, f"ks_{pid}.xlsx")

            try:
                # 先获取标题用于日志
                title = self.scraper.get_video_title(pid)
                safe = sanitize_filename(title, max_len=10)
                if safe:
                    self._log(f"   标题: {title[:60]}", "info")
                    output_path = os.path.join(output_dir, f"{safe}.xlsx")

                comments = self.scraper.fetch_comments(
                    pid,
                    max_pages=max_pages,
                    verbose=False,
                )

                if comments:
                    likes = sum(c['likes'] for c in comments)
                    self._log(f"   ✅ 获取 {len(comments)} 条评论｜{likes} 赞", "success")

                # 子评论
                sub_comments = None
                if include_sub:
                    try:
                        self._log(f"   🔽 DOM 抓取子评论...", "info")
                        sub_comments = self.scraper.fetch_sub_comments_dom(
                            pid, verbose=False,
                        )
                        if sub_comments:
                            scount = sum(len(v) for v in sub_comments.values())
                            self._log(f"   ✅ 子评论: {scount} 条", "success")
                            total_subs += scount
                    except Exception as e:
                        self._log(f"   ⚠ 子评论失败: {e}", "warn")

                self.scraper.export_excel(
                    comments, output_path,
                    video_title=title,
                    video_id=pid,
                    sub_comments=sub_comments,
                )

                ccount = len(comments)
                clikes = sum(c["likes"] for c in comments)
                total_comments += ccount
                total_likes += clikes
                success += 1

            except Exception as e:
                self._log(f"   ❌ 失败: {e}", "error")
                fail += 1

            # 视频间延迟
            if idx < len(self.video_list) and self.scraping:
                time.sleep(batch_delay)

        self._update_progress(len(self.video_list), len(self.video_list),
                              "完成" if not fail else f"完成 ({success}成功/{fail}失败)")

        # 汇总
        elapsed = time.time() - start_time
        self._log("")
        self._log(f"{'='*50}", "info")
        self._log(f"📊 抓取完成！", "success")
        self._log(f"   成功: {success}  失败: {fail}", "info")
        self._log(f"   总评论: {total_comments}  总点赞: {total_likes}", "info")
        if total_subs:
            self._log(f"   子评论: {total_subs}", "info")
        self._log(f"   耗时: {elapsed:.1f} 秒", "info")
        self._log(f"   输出: {os.path.abspath(output_dir)}/", "info")
        self._log(f"{'='*50}", "info")

        self.status_var.set(
            f"完成 — {success}成功/{fail}失败 — {total_comments}条 — {elapsed:.0f}秒"
        )

        self._finish_scrape(True)

    def _finish_scrape(self, ok: bool):
        """恢复 UI"""
        self.scraping = False
        self.root.after(0, lambda: self._set_ui_state(False))


# ───────────────────────────────────────
# 入口
# ───────────────────────────────────────

def main():
    app = KuaishouScraperGUI()
    app.root.mainloop()


if __name__ == "__main__":
    main()
