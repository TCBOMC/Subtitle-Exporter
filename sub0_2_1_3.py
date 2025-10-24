import os
import re
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import shutil
import tempfile
import glob
import sys
import json
import winreg
from pathlib import Path
import ctypes
from fontTools.ttLib import TTFont
from pypinyin import lazy_pinyin

# =======================
# DPI 设置
# =======================
if sys.platform == "win32":
    try:
        # Per-monitor DPI awareness
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

shcore = ctypes.windll.shcore
user32 = ctypes.windll.user32
MONITOR_DEFAULTTONEAREST = 2
MDT_EFFECTIVE_DPI = 0

def get_monitor_dpi(hwnd):
    """返回窗口所在显示器的缩放系数"""
    hmon = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    dpiX = ctypes.c_uint()
    dpiY = ctypes.c_uint()
    shcore.GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI, ctypes.byref(dpiX), ctypes.byref(dpiY))
    return dpiX.value / 96  # 96 DPI 为 100% 缩放

class SubtitleExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量字幕提取工具")
        # 获取 DPI 缩放比例
        self.scale = get_monitor_dpi(self.root.winfo_id())
        # 初始化总表格
        self.font_name_registry = {}  # { "字体文件名": {nameID: {platformID: string, ...}, ... } }

        # 动态设置窗口大小
        base_width, base_height = 600, 400
        self.root.geometry(f"{int(base_width * self.scale)}x{int(base_height * self.scale)}")
        self.root.minsize(width=int(550 * self.scale), height=int(200 * self.scale))
        self.files = []  # 存储(全路径, 文件名)
        self.items = {}  # item_id: (checked, fullpath, filename)
        self.original_files = []  # 用来存储原始文件名，便于还原
        self.renamed_files = []  # 用来保存重命名后的文件与原始文件的映射

        # 文件列表字段顺序
        self.file_fields = ["fullpath", "filename", "checked", "width", "height", "fps", "probe_info"]

        # 定义源 codec 与目标字幕文件格式对应关系
        self.codec_to_subfmt = {
            'ass': 'ass',                # Advanced SubStation Alpha
            'ssa': 'ass',                # SubStation Alpha (同 ASS)
            'subrip': 'srt',             # SubRip
            'webvtt': 'vtt',             # WebVTT
            'dvd_subtitle': 'sub',       # VOBSUB / DVD 字幕
            'microdvd': 'sub',           # MicroDVD 字幕
            'hdmv_pgs_subtitle': 'sup',  # Blu-ray PGS
            'mov_text': 'srt'            # MP4 内嵌字幕，导出为 SRT
        }

        # ==============================
        # 兼容 PyCharm + 打包后两种情况
        # ==============================
        if getattr(sys, "frozen", False):
            # 打包后的 exe 运行环境
            self.base_dir = Path(sys.executable).parent
        else:
            # 源码运行时（PyCharm、命令行）
            self.base_dir = Path(__file__).parent

        self.spp2pgs_exe = self.base_dir / "spp2pgs" / "Spp2Pgs.exe"

        #self.program_dir = Path(sys.executable).parent  # exe 所在目录
        #self.spp2pgs_exe = self.program_dir / "spp2pgs" / "Spp2Pgs.exe"

        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        # 按钮、控件大小缩放
        self.padx = int(10 * self.scale)
        self.pady = int(10 * self.scale)
        self.ipady = int(round(1 * self.scale))
        self.scrollbar_width = int(16 * self.scale)
        self.line_height = int(20 * self.scale)
        self.parameter_width = int(50 * self.scale)
        style.configure("Treeview", rowheight=self.line_height)

        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, pady=self.pady)
        button_frame.columnconfigure(3, weight=1)  # 让按钮间隔可调整
        button_frame.columnconfigure(4, weight=1)  # 让按钮间隔可调整


        self.import_btn = ttk.Button(button_frame, text="导入文件", command=self.import_files)
        self.import_btn.grid(row=0, column=0, padx=(self.padx, self.padx), ipady=self.ipady)

        # 新增 删除按钮
        self.delete_btn = ttk.Button(button_frame, text="删除", command=self.delete_selected, width=5)
        self.delete_btn.grid(row=0, column=1, padx=(0, self.padx), ipady=self.ipady)

        self.clear_btn = ttk.Button(button_frame, text="清空", command=self.clear_all, width=5)
        self.clear_btn.grid(row=0, column=2, padx=(0, self.padx), ipady=self.ipady)

        self.subfmt_frame = ttk.Frame(button_frame)
        self.subfmt_frame.columnconfigure(1, weight=1)
        self.subfmt_frame.grid(row=0, column=3, sticky="e")
        ttk.Label(self.subfmt_frame, text="格式:").grid(row=0, column=0, padx=(0, 5), sticky="e")
        self.subfmt_var = tk.StringVar(value="ass")
        self.subfmt_cb = ttk.Combobox(self.subfmt_frame, textvariable=self.subfmt_var, width=5, state="readonly")
        self.subfmt_cb.grid(row=0, column=1, padx=(0, 0), ipady=self.ipady, sticky="e")

        self.load_subtitle_formats()

        self.ass_fix_var = tk.BooleanVar(value=False)
        self.ass_fix_cb = ttk.Checkbutton(button_frame, text="清空头部", variable=self.ass_fix_var)
        self.ass_fix_cb.grid(row=0, column=4, padx=(self.padx, 0), sticky="w")

        # 绑定事件：当字幕格式变化时更新 ass_fix_cb 显示状态
        self.subfmt_cb.bind("<<ComboboxSelected>>", lambda e: self.update_ass_fix_visibility())

        # 将原来的 restore_font_cb 改为下拉栏：子集合并 / 封装字体
        self.font_mode_var = tk.StringVar(value="封装字体")
        self.font_mode_frame = ttk.Frame(button_frame)
        self.font_mode_frame.grid(row=0, column=5, sticky="w")
        ttk.Label(self.font_mode_frame, text="").grid(row=0, column=0, padx=(0, 0), sticky="e")
        self.font_mode_cb = ttk.Combobox(self.font_mode_frame, textvariable=self.font_mode_var,
                                         values=("封装字体", "子集合并", "字体名还原", "无处理"), width=8, state="readonly")
        self.font_mode_cb.grid(row=0, column=1, padx=(0, self.padx), ipady=self.ipady, sticky="w")

        self.extract_btn = ttk.Button(button_frame, text="提取字幕", command=self.extract_subtitles_clicked)
        self.extract_btn.grid(row=0, column=6, padx=(0, self.scrollbar_width), ipady=self.ipady, sticky="e")

        # 初始显示状态，根据默认值判断
        self.update_ass_fix_visibility()

        # --- 把 Treeview 和 Scrollbar 放在一个 Frame 里 ---
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=(self.padx, 0), pady=(0, self.pady))

        # Treeview with checkbox column
        columns = ("Check", "filename", "width", "height", "fps")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("Check", text="☑", command=self.toggle_all_selection)
        self.tree.heading("filename", text="文件名")
        self.tree.heading("width", text="宽度")
        self.tree.heading("height", text="高度")
        self.tree.heading("fps", text="刷新率")

        self.tree.column("Check", width=self.line_height, anchor="center", stretch=False)
        self.tree.column("filename", width=200)
        self.tree.column("width", width=self.parameter_width, anchor="center", stretch=False)
        self.tree.column("height", width=self.parameter_width, anchor="center", stretch=False)
        self.tree.column("fps", width=self.parameter_width, anchor="center", stretch=False)

        # y 方向滚动条
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        # 布局：Treeview 左，纵向滚动条右
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # 拖拽支持
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', lambda event: self.on_files_dropped(event))

        # 绑定点击事件处理复选框切换
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.enable_treeview_edit(self.tree)

        # 绑定
        def block_resize(event):
            region = self.tree.identify_region(event.x, event.y)
            # 只阻止列标题边界（separator）拖拽
            if region == "separator":
                return "break"
            # 点击其他区域（row, cell, heading）正常处理

        self.tree.bind("<Button-1>", block_resize, add="+")

    def load_subtitle_formats(self):
        formats = [
            "ass",  # Advanced SubStation Alpha
            "srt",  # SubRip Subtitle
            "ssa",  # SubStation Alpha
            #"sub",  # MicroDVD, VobSub, DVD subtitles (图片字幕但常作为subtitle格式名)
            #"vtt",  # WebVTT
            "sup",  # SubPicture Subtitle
            "原格式",
            #"lrc",
            #"mov_text",  # QuickTime text subtitles (比如mp4里的字幕)
            #"pgs",  # Presentation Graphic Stream (Blu-ray)
            #"dvdsub",  # DVD subtitles (图像字幕)
            #"xsub",  # XVid subtitles
            #"hdmv_pgs_subtitle",  # 高清多媒体播放机PGS字幕
            #"webvtt"  # Web Video Text Tracks (类似vtt)
        ]

        self.subfmt_cb['values'] = formats
        self.subfmt_var.set("ass")

    def update_ass_fix_visibility(self):
        fmt = self.subfmt_var.get().lower()
        if fmt in ("ass", "ssa", "原格式"):
            self.ass_fix_cb.config(state="normal")
            self.font_mode_cb.config(state="normal")
        else:
            self.ass_fix_cb.config(state="disabled")
            self.font_mode_cb.config(state="disabled")

    # ✅ 静默运行子进程（Windows下隐藏控制台窗口）
    def run_silently(self, cmd, **kwargs):
        """静默运行命令行指令并在失败时抛出异常"""
        if os.name == 'nt':  # Windows系统
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            kwargs['startupinfo'] = startupinfo

        kwargs.setdefault('stdout', subprocess.PIPE)
        kwargs.setdefault('stderr', subprocess.PIPE)
        kwargs.setdefault('stdin', subprocess.DEVNULL)

        result = subprocess.run(cmd, **kwargs)
        if result.returncode != 0:
            stderr = result.stderr.decode(errors='ignore').strip()
            raise RuntimeError(f"ffmpeg 执行失败 (code={result.returncode}):\n{stderr}")
        return result

    # ✅ 自动定位 ffmpeg.exe 或系统 ffmpeg
    def get_ffmpeg_exe(self):
        ffmpeg_exe = os.path.join(self.base_dir, "ffmpeg", "ffmpeg.exe")
        return ffmpeg_exe if os.path.exists(ffmpeg_exe) else "ffmpeg"

    # ✅ 自动定位 ffprobe.exe 或系统 ffprobe
    def get_ffprobe_exe(self):
        ffprobe_exe = os.path.join(self.base_dir, "ffmpeg", "ffprobe.exe")
        return ffprobe_exe if os.path.exists(ffprobe_exe) else "ffprobe"

    # ✅ 替代 ffmpeg.run()
    def silent_ffmpeg_run(self, args, **kwargs):
        """
        使用 subprocess 直接调用 ffmpeg，静默执行。
        参数 args 为 ffmpeg 参数列表（不含可执行路径）。
        """
        ffmpeg_exe = self.get_ffmpeg_exe()
        cmd = [ffmpeg_exe] + args
        result = self.run_silently(cmd, **kwargs)

        if result.returncode != 0:
            raise RuntimeError(
                f"❌ ffmpeg 执行失败 (code={result.returncode})\n{result.stderr.decode(errors='ignore')}"
            )
        return result

    # ✅ 替代 ffmpeg.probe()
    def silent_ffmpeg_probe(self, filename, **kwargs):
        """
        使用 subprocess 调用 ffprobe 获取视频信息。
        """
        ffprobe_exe = self.get_ffprobe_exe()
        cmd = [
            ffprobe_exe,
            "-v", "quiet",
            "-print_format", "json",
            "-show_streams",
            "-show_format",
            filename
        ]

        result = self.run_silently(cmd, **kwargs)
        if result.returncode != 0:
            raise RuntimeError(
                f"❌ ffprobe 执行失败 (code={result.returncode})\n{result.stderr.decode(errors='ignore')}"
            )

        return json.loads(result.stdout.decode('utf-8', errors='ignore'))

    def import_files(self):
        paths = filedialog.askopenfilenames(title="选择视频文件")
        if not paths:
            return
        #self.add_files(paths)
        t = threading.Thread(target=self.add_files, args=(paths,))
        t.start()

    def add_files(self, paths):
        for p in paths:
            p = p.strip("'\"")
            if os.path.isfile(p) and p not in [f[0] for f in self.files]:
                filename = os.path.basename(p)
                width = ""
                height = ""
                fps = ""
                probe_info = None

                try:
                    info = self.silent_ffmpeg_probe(p)
                    probe_info = info  # 保存完整视频信息
                    for stream in info.get("streams", []):
                        if stream.get("codec_type") == "video":
                            width = stream.get("width", "")
                            height = stream.get("height", "")
                            r_frame_rate = stream.get("r_frame_rate", "")
                            if r_frame_rate and r_frame_rate != "0/0":
                                num, den = map(int, r_frame_rate.split("/"))
                                if den != 0:
                                    fps = round(num / den, 3)
                            break
                except Exception:
                    pass

                # 用统一字段顺序创建 tuple
                file_tuple = (
                    p,  # fullpath
                    filename,  # filename
                    True,  # checked
                    width,  # width
                    height,  # height
                    fps,  # fps
                    probe_info  # full probe 信息
                )
                self.files.append(file_tuple)
                self.original_files.append((p, filename))
                # 安全刷新UI
                self.root.after(0, self.refresh_tree)

            else:
                if not os.path.isfile(p):
                    print(f"无效的文件路径: {p}")  # 如果路径无效，输出提示
                else:
                    print(f"文件已存在: {p}")  # 如果文件已经在列表中，输出提示

    def clear_list(self):
        self.files.clear()
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for file_tuple in self.files:
            file_info = dict(zip(self.file_fields, file_tuple))
            checked = file_info.get("checked")
            fullpath = file_info.get("fullpath")
            filename = file_info.get("filename", "")
            width = file_info.get("width", "")
            height = file_info.get("height", "")
            fps = file_info.get("fps", "")
            probe_info = file_info.get("probe_info", "")
            chk = "☑" if checked else "☐"
            self.tree.insert("", tk.END, iid=fullpath, values=(chk, filename, width, height, fps, probe_info))
        self.update_header_checkbox()

    def delete_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选中要删除的文件行")
            return

        # 删除 self.files 中对应项
        paths_to_delete = set(selected_items)
        self.files = [f for f in self.files if f[0] not in paths_to_delete]

        self.refresh_tree()

    """def ask_output_directory(self):
        dialog = OutputDirDialog(self.root, title="字幕提取目标目录")
        return dialog.result"""

    def extract_subtitles_clicked(self):
        if not self.files:
            messagebox.showwarning("提示", "请先导入视频文件")
            return

        # ✅ 直接从 self.files 获取 fullpath、filename、height、fps
        selected_files = []
        for file_tuple in self.files:
            file_info = dict(zip(self.file_fields, file_tuple))
            if file_info.get("checked"):
                fullpath = file_info.get("fullpath")
                filename = file_info.get("filename")
                selected_files.append((fullpath, filename))

        if not selected_files:
            messagebox.showwarning("提示", "请先勾选要提取字幕的文件")
            return

        subfmt = self.subfmt_var.get().lower()
        if not subfmt:
            messagebox.showerror("错误", "请选择字幕格式")
            return

        # 使用 askyesnocancel 代替原来的 askquestion
        choice = messagebox.askyesnocancel(
            "字幕提取目标目录",
            "字幕默认会被提取到各个视频所在的目录！\n选择“是”继续操作；\n选择“否”指定提取目录。"
        )

        if choice is None:
            return  # 用户点击取消或关闭窗口

        if choice:  # Yes → 原目录
            outdir = None
        else:  # No → 指定目录
            outdir = filedialog.askdirectory(title="选择字幕保存目录")
            if not outdir:  # 用户在选择目录时点了取消
                return

            # 如果目录不存在，则回退到上级存在的目录
            temp_dir = outdir
            while not os.path.exists(temp_dir):
                temp_dir = os.path.dirname(temp_dir)
            outdir = temp_dir

            # 尝试创建最终目录
            try:
                os.makedirs(outdir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建目录：\n{outdir}\n\n错误信息：{e}")
                return

        self.save_and_disable_buttons()
        threading.Thread(target=self.extract_subtitles_all, args=(subfmt, selected_files, outdir), daemon=True).start()

    def extract_subtitles_all(self, subfmt, selected_files, outdir=None):
        need_merge_fonts = False
        temp_outdir = None
        used_temp_dir = False
        """统一调度字幕提取"""

        # ✅ 如果 outdir 为 None，仅创建临时目录备用（但不替换 outdir）
        if outdir is None:
            temp_outdir = tempfile.mkdtemp(prefix="subs_extract_")
            used_temp_dir = True
            self.log(f"未指定输出目录，已创建临时目录供合并字体使用: {temp_outdir}")

        # --- 清空所有 Treeview 行的染色 ---
        for item_id in self.tree.get_children():
            self.tree.item(item_id, tags=())

        # --- 设置染色标签 ---
        self.tree.tag_configure("success", background="#c8e6c9")  # 绿色
        self.tree.tag_configure("partial", background="#fff9c4")  # 黄色
        self.tree.tag_configure("fail", background="#ffcdd2")  # 红色
        self.tree.tag_configure("processing", background="#e0e0e0")  # 灰色

        for seq_num, (fullpath, filename) in enumerate(selected_files, 1):
            had_error = False
            item_id = None
            for iid in self.tree.get_children():
                if self.tree.item(iid, "values")[1] == filename:
                    item_id = iid
                    break

            # --- 开始处理前先染灰色 ---
            if item_id:
                self.tree.item(item_id, tags=("processing",))
                self.tree.update_idletasks()  # 确保立即刷新界面
            generated_subs_for_video = []
            #print(f"fullpath:{fullpath}")
            self.log(f"分析字幕轨道：{filename}")
            file_info = next((dict(zip(self.file_fields, f)) for f in self.files if f[0] == fullpath), None)
            print(outdir)
            if not file_info:
                self.log(f"无法找到视频信息: {filename}，跳过")
                if item_id:
                    self.tree.item(item_id, tags=("fail",))
                continue

            height = file_info.get("height", 1080)
            width = file_info.get("width", 1920)
            fps = file_info.get("fps", 23.976)
            probe = file_info.get("probe_info")  # 直接使用已存 probe
            print(f"文件参数：{file_info}")
            print(f"fullpath:{fullpath}, width:{width}, height:{height}, fps:{fps}")

            if not probe:
                self.log(f"缺少 probe 信息: {filename}，跳过")
                if item_id:
                    self.tree.item(item_id, tags=("fail",))
                continue

            subtitle_streams = [s for s in probe.get('streams', []) if s.get('codec_type') == 'subtitle']
            if not subtitle_streams:
                self.log(f"{filename} 没有字幕轨道，跳过")
                if item_id:
                    self.tree.item(item_id, tags=("fail",))
                continue

            if subfmt == "原格式":
                detected_formats = list({s.get("codec_name", "").lower() for s in subtitle_streams})
                self.log(f"检测到字幕格式: {detected_formats}")

            global_mapping = {}
            font_mode = self.font_mode_var.get() if hasattr(self, 'font_mode_var') else "子集合并"
            temp_font_dir = None

            has_ass_ssa = any(s.get("codec_name", "").lower() in ("ass", "ssa") for s in subtitle_streams)

            # 提前处理字体提取逻辑
            if (font_mode == "封装字体" and subfmt in ("ass", "ssa")) or subfmt == "sup" or (subfmt == "原格式" and font_mode == "封装字体" and has_ass_ssa):
                try:
                    temp_font_dir = self.extract_all_fonts_to_tempdir(fullpath)
                    self.log(f"临时字体目录已创建: {temp_font_dir}")
                except Exception as e:
                    self.log(f"提取临时字体失败: {e}")
                    temp_font_dir = None
                    had_error = True
            elif font_mode == "封装字体":
                self.log(f"跳过封装字体：输出格式为 {subfmt}，仅在 ass/ssa 时封装字体。")

            try:
                for stream in subtitle_streams:
                    try:
                        codec_name = stream.get("codec_name", "").lower()
                        mapped = self.codec_to_subfmt.get(codec_name, codec_name)  # 映射（若无则退回 codec_name）
                        #print(mapped)
                        # 如果用户指定 "原格式"，就用映射后的值；否则沿用外部传进来的 subfmt
                        cur_subfmt = mapped if subfmt == "原格式" else subfmt

                        outpath, mapping = self.extract_single_subtitle(
                            fullpath=fullpath,
                            filename=filename,
                            stream=stream,
                            subfmt=cur_subfmt,
                            outdir=outdir,
                            font_mode=font_mode,
                            temp_font_dir=temp_font_dir,
                            width=width,
                            height=height,
                            fps=fps
                        )
                        if outpath:
                            generated_subs_for_video.append(outpath)
                        else:
                            had_error = True
                        global_mapping.update(mapping)
                    except Exception as e:
                        # 可以记录这个流的错误，但继续处理其他流
                        self.log(f"字幕流 {stream} 处理失败: {e}")
                        had_error = True
                        continue
            finally:
                # 无论如何，所有流处理完后都会删除缓存
                if subfmt == "sup" and temp_font_dir:
                    shutil.rmtree(temp_font_dir)

            if subfmt == "sup":
                # --- 根据执行情况染色 ---
                if item_id:
                    if had_error and not generated_subs_for_video:
                        self.tree.item(item_id, tags=("fail",))
                    elif had_error and generated_subs_for_video:
                        self.tree.item(item_id, tags=("partial",))
                    else:
                        self.tree.item(item_id, tags=("success",))
                continue  # sup字幕不需要后续处理

            # 处理子集字体还原逻辑
            if (font_mode == "子集合并" and subfmt in ("ass", "ssa")) or (font_mode == "子集合并" and subfmt == "原格式" and has_ass_ssa):
                fonts_root = os.path.join(outdir or temp_outdir, "Fonts")
                os.makedirs(fonts_root, exist_ok=True)
                video_fonts_dir = os.path.join(fonts_root, f"{seq_num}_{os.path.splitext(filename)[0]}")
                if os.path.exists(video_fonts_dir):
                    shutil.rmtree(video_fonts_dir)
                os.makedirs(video_fonts_dir, exist_ok=True)
                try:
                    self.extract_fonts_from_video(fullpath, fonts_root, global_mapping, seq_num)
                    need_merge_fonts = True
                except Exception as e:
                    self.log(f"删除字体失败: {e}")
                    had_error = True

            # 封装字体逻辑
            if (font_mode == "封装字体" and temp_font_dir and subfmt in ("ass", "ssa")) or (subfmt == "原格式" and has_ass_ssa):
                try:
                    for subfile in generated_subs_for_video:
                        if os.path.splitext(subfile)[1].lower() in ('.ass', '.ssa'):
                            try:
                                self.embed_fonts_to_ass(subfile, temp_font_dir)
                                self.log(f"字体已封装到: {os.path.basename(subfile)}")
                            except Exception as e:
                                self.log(f"封装字体失败: {subfile} 错误: {e}")
                                had_error = True
                finally:
                    try:
                        shutil.rmtree(temp_font_dir)
                        self.log(f"临时字体目录已删除: {temp_font_dir}")
                    except Exception as e:
                        self.log(f"删除临时字体目录失败: {e}")
                        had_error = True

            # --- 根据执行情况染色 ---
            if item_id:
                if had_error and not generated_subs_for_video:
                    self.tree.item(item_id, tags=("fail",))
                elif had_error and generated_subs_for_video:
                    self.tree.item(item_id, tags=("partial",))
                else:
                    self.tree.item(item_id, tags=("success",))

        # 全局字体合并
        if (self.font_mode_var.get() == "子集合并" and subfmt != "sup") or (self.font_mode_var.get() == "子集合并" and subfmt == "原格式" and need_merge_fonts):
            fonts_root = os.path.join(outdir or temp_outdir, "Fonts")
            if os.path.exists(fonts_root):
                self.merge_fonts(fonts_root)
                self.log("📚 所有视频字体已合并到 Fonts 根目录")

        # ✅ 若使用了临时目录，则进行导出提示
        if used_temp_dir:
            fonts_root = os.path.join(temp_outdir, "Fonts")
            if os.path.exists(fonts_root):
                export_fonts = messagebox.askyesno("导出字体", "是否导出提取的字体？")
                if export_fonts:
                    export_dir = filedialog.askdirectory(title="选择导出字体的目标目录")
                    if export_dir:
                        try:
                            # 如果目录不存在，则回退到上级存在的目录
                            temp_dir = export_dir
                            while not os.path.exists(temp_dir):
                                temp_dir = os.path.dirname(temp_dir)
                            export_dir = temp_dir
                            # 🟢 在导出目录中创建 Fonts 文件夹
                            export_fonts_dir = os.path.join(export_dir, "Fonts")
                            os.makedirs(export_fonts_dir, exist_ok=True)

                            for item in os.listdir(fonts_root):
                                src = os.path.join(fonts_root, item)
                                dst = os.path.join(export_fonts_dir, item)
                                if os.path.isdir(src):
                                    shutil.copytree(src, dst, dirs_exist_ok=True)
                                else:
                                    shutil.copy2(src, dst)
                            self.log(f"字体已导出到: {export_fonts_dir}")
                        except Exception as e:
                            self.log(f"导出字体失败: {e}")
                # 删除临时目录
                try:
                    shutil.rmtree(temp_outdir)
                    self.log(f"临时输出目录已删除: {temp_outdir}")
                except Exception as e:
                    self.log(f"删除临时输出目录失败: {e}")

        self.log("✅ 所有文件字幕提取完成！")
        messagebox.showinfo("完成", "所有字幕提取完成！")
        self.restore_buttons_state()

    def extract_single_subtitle(self, fullpath, filename, stream, subfmt, outdir,
                                font_mode, temp_font_dir, width, height, fps):
        """提取单个字幕轨道，逻辑统一化，SUP 仅在无法直接copy时特殊处理"""
        idx = stream['index']
        tags = stream.get('tags', {})
        lang = tags.get('language', 'unknown')
        codec_name = stream.get('codec_name', '').lower()

        outfilename = f"{os.path.splitext(filename)[0]}.{lang}{idx}.{subfmt}"
        outpath = os.path.join(outdir, outfilename) if outdir else os.path.join(os.path.dirname(fullpath), outfilename)
        self.log(f"提取轨道 {idx} ({lang}) → {outfilename}")

        original_fmt = self.codec_to_subfmt.get(codec_name)
        can_copy = (original_fmt == subfmt)
        ffmpeg_exe = self.get_ffmpeg_exe()

        try:
            if can_copy:
                # ✅ 直接拷贝字幕轨
                cmd = [
                    ffmpeg_exe,
                    "-y",  # overwrite_output
                    "-i", fullpath,
                    "-map", f"0:{idx}",
                    "-c:s", "copy",
                    outpath
                ]
                self.run_silently(cmd)
                self.log(f"✅ 直接导出字幕: {outfilename}")
            elif subfmt == "sup":
                # 不能copy且目标为SUP时，使用特殊转换逻辑
                outpath = self.handle_sup_conversion(fullpath, filename, stream, outdir, temp_font_dir, width, height, fps)
                return outpath, {}
            else:
                # ✅ 转换字幕为目标格式
                cmd = [
                    ffmpeg_exe,
                    "-y",
                    "-i", fullpath,
                    "-map", f"0:{idx}",
                    "-c:s", subfmt,
                    outpath
                ]
                #print(f"执行命令：{' '.join(cmd)}")
                self.run_silently(cmd)
                #print(original_fmt)
                if original_fmt == "srt" and subfmt in ("ass", "ssa"):
                    print("修改分辨率")
                    #self.set_ass_resolution(outpath, width, height)
                self.log(f"✅ 成功导出字幕文件: {outfilename}")
        except Exception as e:
            # 第一次失败后尝试使用 original_fmt
            self.log(f"⚠️ 无法转换字幕为 {subfmt}: {e}")
            self.log(f"尝试使用备用格式 {original_fmt} 重新导出...")

            try:
                fallback_outfilename = f"{os.path.splitext(filename)[0]}.{lang}{idx}.{original_fmt}"
                fallback_outpath = os.path.join(outdir, fallback_outfilename) if outdir else \
                    os.path.join(os.path.dirname(fullpath), fallback_outfilename)

                cmd = [
                    ffmpeg_exe,
                    "-y",
                    "-i", fullpath,
                    "-map", f"0:{idx}",
                    "-c:s", "copy",
                    fallback_outpath
                ]
                self.run_silently(cmd)
                self.log(f"✅ 使用备用格式成功导出: {fallback_outfilename}")
                outpath = fallback_outpath  # 替换为备用导出路径
            except Exception as e2:
                self.log(f"❌ 备用格式提取也失败: {e2}")
                messagebox.showerror("提取失败", f"轨道 {idx} 文件 {outfilename} 与备用格式导出均失败。\n错误: {e2}")
                return None, {}

        # 后处理逻辑（ASS字体修复、字体还原）
        if self.ass_fix_var.get() and subfmt in ("ass", "ssa") and font_mode !="无":
            self.fix_ass_header(outpath)

        if subfmt in ("ass", "ssa") and font_mode in ("子集合并", "字体名还原"):
            try:
                mapping = self.restore_ass_fonts(outpath)
                self.log(f"ASS 字体已还原: {os.path.basename(outpath)}")
                return outpath, mapping
            except Exception as e:
                self.log(f"字体还原失败: {outfilename} 错误: {e}")

        return outpath, {}

    def handle_sup_conversion(self, fullpath, filename, stream, outdir, temp_font_dir, width, height, fps):
        """处理 SUP 特殊逻辑：生成临时 ASS，再转换为 SUP"""
        ass_temp_dir = tempfile.mkdtemp(prefix="ass_temp_")
        outpath = None  # 用于返回生成的 SUP 路径

        # 提取临时 ASS
        ass_path, _ = self.extract_single_subtitle(
            fullpath=fullpath,
            filename=filename,
            stream=stream,
            subfmt="ass",
            outdir=ass_temp_dir,
            font_mode="无",
            temp_font_dir=None,
            width=width,
            height=height,
            fps=fps
        )

        if not ass_path or not os.path.exists(ass_path):
            self.log(f"❌ 无法生成临时 ASS，SUP 转换中止。")
            shutil.rmtree(ass_temp_dir, ignore_errors=True)
            if temp_font_dir:
                shutil.rmtree(temp_font_dir, ignore_errors=True)
            return None

        # 转为 SUP
        lang = stream.get("tags", {}).get("language", "unknown")
        outfilename = f"{os.path.splitext(filename)[0]}.{lang}{stream['index']}.sup"
        outpath = os.path.join(outdir or os.path.dirname(fullpath), outfilename)

        try:
            success = self.generate_subtitles(
                ass_file=ass_path,
                fonts_dir=temp_font_dir or "",
                out_sup=outpath,
                height=height,
                fps=fps
            )
            if success:
                self.log(f"✅ SUP 生成成功: {outfilename}")
            else:
                self.log(f"⚠️ SUP 生成失败: {outfilename}")
                outpath = None  # 失败则返回 None
        except Exception as e:
            self.log(f"💥 调用 generate_subtitles 出错: {e}")
            outpath = None
        finally:
            # 清理临时目录
            try:
                shutil.rmtree(ass_temp_dir)
                self.log("SUP 临时文件已清理。")
            except Exception:
                pass

        return outpath

    def prepare_environment(self, fonts_dir: str):
        """环境准备：注册缺失字体，返回临时注册的字体信息"""
        fonts_dir = Path(fonts_dir)

        def get_font_name(font_path: Path) -> str:
            """读取字体内部名称"""
            try:
                tt = TTFont(font_path)
                name_records = tt["name"].names
                full_name = None
                family_name = None
                for record in name_records:
                    if record.nameID == 4 and not full_name:
                        full_name = record.string.decode(record.getEncoding(), errors="ignore").strip()
                    if record.nameID == 1 and not family_name:
                        family_name = record.string.decode(record.getEncoding(), errors="ignore").strip()
                tt.close()
                return full_name or family_name or font_path.stem
            except Exception:
                return font_path.stem

        def broadcast_font_change():
            """广播 WM_FONTCHANGE 通知系统字体表更新"""
            HWND_BROADCAST = 0xFFFF
            WM_FONTCHANGE = 0x001D
            SMTO_ABORTIFHUNG = 0x0002
            ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)

        def get_installed_fonts() -> set:
            """读取系统和当前用户注册的所有字体名称"""
            installed = set()
            reg_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"),
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"),
            ]
            for root, path in reg_paths:
                try:
                    with winreg.OpenKey(root, path) as key:
                        i = 0
                        while True:
                            try:
                                name, _, _ = winreg.EnumValue(key, i)
                                installed.add(name.lower())
                                i += 1
                            except OSError:
                                break
                except FileNotFoundError:
                    continue
            return installed

        def get_unique_filename(dst_dir: Path, src_name: str) -> Path:
            """生成唯一文件名"""
            dst = dst_dir / src_name
            counter = 1
            stem = dst.stem
            suffix = dst.suffix
            while dst.exists():
                dst = dst_dir / f"{stem}_{counter}{suffix}"
                counter += 1
            return dst

        username = os.getlogin()
        user_font_dir = Path(f"C:/Users/{username}/AppData/Local/Microsoft/Windows/Fonts")
        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"

        user_font_dir.mkdir(parents=True, exist_ok=True)
        files = list(fonts_dir.glob("*.[ot]tf"))
        if not files:
            print(f"字体目录为空：{fonts_dir}")
            return []

        installed_fonts = get_installed_fonts()
        registered = []

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE) as key:
            for src in files:
                font_name = get_font_name(src)
                reg_name = f"{font_name} (TrueType)"
                if reg_name.lower() in installed_fonts:
                    print(f"跳过已安装字体：{font_name}")
                    continue
                dst = get_unique_filename(user_font_dir, src.name)
                shutil.copy2(src, dst)
                winreg.SetValueEx(key, reg_name, 0, winreg.REG_SZ, str(dst))
                registered.append((reg_name, dst))
                print(f"注册字体：{font_name} → {dst.name}")

        if registered:
            broadcast_font_change()
            print(f"✅ 已临时注册 {len(registered)} 个字体。")
        else:
            print("ℹ️ 所有字体均已安装，无需注册。")

        return registered

    def generate_subtitles(self, ass_file: str, fonts_dir: str, out_sup: str, height: int = 1080, fps: float = 23.976):
        """字幕生成：注册字体、调用 Spp2Pgs"""
        ass_file = Path(ass_file)
        fonts_dir = Path(fonts_dir)
        out_sup = Path(out_sup)

        if not ass_file.exists():
            print("❌ ASS 文件不存在：", ass_file)
            return False
        if not fonts_dir.exists():
            print("❌ 字体目录不存在：", fonts_dir)
            return False

        # 优先使用本地路径，其次 PATH 环境变量
        exe_path = str(self.spp2pgs_exe if self.spp2pgs_exe.exists() else shutil.which("Spp2Pgs") or shutil.which("Spp2Pgs.exe"))
        #exe_path = str(self.spp2pgs_exe if self.spp2pgs_exe.exists() else shutil.which("Spp2Pgs") or shutil.which("Spp2Pgs.exe"))

        if not exe_path or not Path(exe_path).exists():
            print("❌ 未找到 Spp2Pgs 可执行文件。")
            #messagebox.showerror("错误", "未找到 Spp2Pgs 可执行文件。")
            return False

        print("使用 Spp2Pgs 可执行：", exe_path)

        registered_fonts = self.prepare_environment(fonts_dir)

        try:
            cmd = [exe_path, "-i", str(ass_file), "-s", str(height), "-r", str(fps), str(out_sup)]
            #print("执行命令：", " ".join(cmd))
            # 打包后调试去掉 CREATE_NO_WINDOW
            p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            print(p.stdout)
            if "Encoding successfully completed." in p.stdout:
                print("✅ Spp2Pgs 生成 SUP 成功：", out_sup)
                return True
            else:
                print("⚠️ Spp2Pgs 处理失败，请检查输出。")
                return False
        except Exception as e:
            print("💥 执行 Spp2Pgs 时发生异常：", e)
            return False
        finally:
            self.cleanup_environment(registered_fonts)

    def cleanup_environment(self, registered_fonts):
        """清理环境：卸载临时注册字体"""

        def broadcast_font_change():
            HWND_BROADCAST = 0xFFFF
            WM_FONTCHANGE = 0x001D
            SMTO_ABORTIFHUNG = 0x0002
            ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)

        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
        count = 0
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
                for reg_name, dst in registered_fonts:
                    try:
                        winreg.DeleteValue(key, reg_name)
                    except FileNotFoundError:
                        pass
                    if dst.exists():
                        try:
                            dst.unlink()
                        except Exception:
                            pass
                    count += 1
            if count:
                broadcast_font_change()
                print(f"🧹 已卸载 {count} 个临时字体。")
        except Exception as e:
            print(f"卸载字体时出错：{e}")

    def fix_ass_header(self, filepath):
        """
        修正 ASS/SSA 字幕的 [Script Info] 头部：
        - 保留注释顺序
        - 在注释下面插入标准字段
        - 保证 [Script Info] 区块末尾有一个空行
        - 不影响其他区块
        """
        try:
            with open(filepath, "rb") as f:
                content = f.read()

            # 移除 UTF-8 BOM
            if content.startswith(b'\xef\xbb\xbf'):
                content = content[3:]

            text = content.decode("utf-8", errors="ignore")

            # 标准头部字段（顺序固定）
            standard_info = [
                ("Title", "Untitled"),
                ("ScriptType", "v4.00+"),
                ("Collisions", "Normal"),
                ("PlayDepth", "0")
            ]

            lines = text.splitlines()
            new_lines = []
            inside_script_info = False

            for i, line in enumerate(lines):
                stripped = line.strip()
                if stripped.lower() == "[script info]":
                    inside_script_info = True
                    new_lines.append("[Script Info]")
                    continue

                # 到达下一个区块时结束 [Script Info]
                if inside_script_info and stripped.startswith("[") and stripped.endswith("]"):
                    # 在注释下面插入标准字段
                    for key, val in standard_info:
                        new_lines.append(f"{key}: {val}")
                    # 确保 [Script Info] 区块末尾有一个空行
                    if new_lines[-1].strip() != "":
                        new_lines.append("")
                    inside_script_info = False
                    new_lines.append(line)
                    continue

                if inside_script_info:
                    # 保留注释和空行
                    if stripped.startswith(";") or stripped == "":
                        new_lines.append(line)
                    # 非注释行忽略（用标准字段替换）
                    continue
                else:
                    new_lines.append(line)

            # 如果文件以 [Script Info] 结束，需要在最后插入标准字段
            if inside_script_info:
                for key, val in standard_info:
                    new_lines.append(f"{key}: {val}")
                if new_lines[-1].strip() != "":
                    new_lines.append("")

            # 写回文件
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("\n".join(new_lines) + "\n")  # 文件末尾再加一个换行

            self.log(f"ASS头部已修正: {os.path.basename(filepath)}")
        except Exception as e:
            self.log(f"ASS头部修正失败: {e}")

    def set_ass_resolution(self, filepath, width, height):
        """
        设置 ASS/SSA 字幕的分辨率参数：
        - 定位到 [Script Info] 区块
        - 如果存在 PlayResX / PlayResY，则修改为指定值
        - 如果不存在，则插入新字段
        - 保留注释和其他字段顺序
        - 不影响其他区块
        """
        print(f"width:{width}, height:{height}")
        try:
            with open(filepath, "rb") as f:
                content = f.read()

            # 移除 UTF-8 BOM
            if content.startswith(b'\xef\xbb\xbf'):
                content = content[3:]

            text = content.decode("utf-8", errors="ignore")
            lines = text.splitlines()
            new_lines = []
            inside_script_info = False
            found_x = found_y = False

            for i, line in enumerate(lines):
                stripped = line.strip()

                # 进入 [Script Info]
                if stripped.lower() == "[script info]":
                    inside_script_info = True
                    new_lines.append("[Script Info]")
                    continue

                # 区块结束
                if inside_script_info and stripped.startswith("[") and stripped.endswith("]"):
                    # 如果没有找到 PlayResX/Y，则补上
                    if not found_x:
                        new_lines.append(f"PlayResX: {width}")
                    if not found_y:
                        new_lines.append(f"PlayResY: {height}")
                    # 区块结束前确保有空行
                    if new_lines[-1].strip() != "":
                        new_lines.append("")
                    inside_script_info = False
                    new_lines.append(line)
                    continue

                if inside_script_info:
                    # 修改现有分辨率字段
                    if stripped.lower().startswith("playresx:"):
                        new_lines.append(f"PlayResX: {width}")
                        found_x = True
                    elif stripped.lower().startswith("playresy:"):
                        new_lines.append(f"PlayResY: {height}")
                        found_y = True
                    else:
                        new_lines.append(line)
                    continue
                else:
                    new_lines.append(line)

            # 如果文件以 [Script Info] 结束
            if inside_script_info:
                if not found_x:
                    new_lines.append(f"PlayResX: {width}")
                if not found_y:
                    new_lines.append(f"PlayResY: {height}")
                if new_lines[-1].strip() != "":
                    new_lines.append("")

            # 写回文件
            with open(filepath, "w", encoding="utf-8") as f:
                f.write("\n".join(new_lines) + "\n")

            self.log(f"已更新分辨率: {os.path.basename(filepath)} ({width}x{height})")

        except Exception as e:
            self.log(f"修改分辨率失败: {e}")

    # ----------------- 新增：extract_all_fonts_to_tempdir（含字体重命名） -----------------
    def extract_all_fonts_to_tempdir(self, video_path):
        temp_dir = tempfile.mkdtemp(prefix="sub_fonts_")

        ffmpeg_exe = self.get_ffmpeg_exe()
        old_cwd = os.getcwd()
        try:
            os.chdir(temp_dir)

            # ✅ 构建 ffmpeg 命令
            cmd = [
                ffmpeg_exe,
                "-dump_attachment:t", "",  # 提取所有附件
                "-i", video_path
            ]

            #self.log(f"📦 提取字体附件: {os.path.basename(video_path)} → {temp_dir}")
            result = self.run_silently(cmd)
            #self.log("✅ 字体提取完成")

        except RuntimeError as e:
            # 如果报错信息里包含“At least one output file must be specified”，忽略
            if "At least one output file must be specified" in str(e):
                self.log("⚠️ 忽略 ffmpeg 报错：附件已经提取")
        finally:
            os.chdir(old_cwd)

        # 保留 .ttf/.otf 并重命名
        for f in os.listdir(temp_dir):
            fpath = os.path.join(temp_dir, f)
            ext = os.path.splitext(f)[1].lower()
            if ext not in ('.ttf', '.otf'):
                try:
                    os.remove(fpath)
                except Exception:
                    pass
                continue

            fname = os.path.splitext(f)[0].split('.')[0]  # 去掉第一个 . 之后的内容
            new_name = fname + ext
            new_path = os.path.join(temp_dir, new_name)
            if fpath != new_path:
                os.rename(fpath, new_path)

        return temp_dir

    # ----------------- 新增：embed_fonts_to_ass -----------------
    def embed_fonts_to_ass(self, ass_path, font_dir):

        if not os.path.exists(ass_path):
            raise FileNotFoundError(ass_path)

        font_files = [p for p in glob.glob(os.path.join(font_dir, "*"))
                      if os.path.splitext(p)[1].lower() in ('.ttf', '.otf')]
        if not font_files:
            return

        def encode_font_bytes(data: bytes) -> str:
            """将字体二进制转换为 Aegisub 样式的 UUencode 文本（每行80字符）"""
            encoded = []
            for i in range(0, len(data), 3):
                chunk = data[i:i + 3]
                while len(chunk) < 3:
                    chunk += b'\0'
                b1, b2, b3 = chunk
                v1 = b1 >> 2
                v2 = ((b1 & 0x3) << 4) | (b2 >> 4)
                v3 = ((b2 & 0xF) << 2) | (b3 >> 6)
                v4 = b3 & 0x3F
                encoded.extend(chr(v + 33) for v in (v1, v2, v3, v4))
            text = "".join(encoded)
            return "\n".join(text[i:i + 80] for i in range(0, len(text), 80))

        # 构建 [Fonts] 段落
        entries = []
        for fpath in font_files:
            fname = os.path.basename(fpath)
            with open(fpath, 'rb') as fb:
                b = fb.read()
            enc_text = encode_font_bytes(b)
            entries.append(f"fontname: {fname}\n{enc_text}")

        with open(ass_path, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        font_block = "[Fonts]\n" + "\n".join(entries) + "\n"

        if "[Fonts]" in text:
            text = text.replace("[Fonts]", font_block, 1)
        else:
            m = re.search(r"(?m)^(\[V4\+ Styles\]|\[Events\])", text)
            insert_block = font_block + "\n"
            if m:
                idx = m.start(0)
                text = text[:idx] + insert_block + text[idx:]
            else:
                text = insert_block + text

        with open(ass_path, 'w', encoding='utf-8') as f:
            f.write(text)

        return

    def restore_ass_fonts(self, filepath):
        """
        批量将ASS文件中子集化字体名还原为原字体名：
        - 替换 [V4+ Styles] 中的字体
        - 替换 Dialogue 行中 {} 内的字体
        """

        # 读取文件
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.readlines()

        # 提取字体映射
        mapping = {}
        for line in content:
            m = re.search(r";\s*Font subset:\s*([A-Z0-9]+)\s*-\s*(.+)", line)
            if m:
                subset, realname = m.groups()
                mapping[subset.strip()] = realname.strip()

        # 匹配 Style 行
        style_pattern = re.compile(r"^(Style:\s*[^,]+,)([^,]+)(,.*)$")
        # 匹配 Dialogue 中 {...}
        braces_pattern = re.compile(r"\{([^}]*)\}")

        new_lines = []
        for line in content:
            # 替换 Style 字体
            m = style_pattern.match(line)
            if m:
                prefix, fontname, suffix = m.groups()
                new_fontname = mapping.get(fontname, fontname)
                new_lines.append(f"{prefix}{new_fontname}{suffix}\n")
                continue

            # 替换 Dialogue 内所有 {} 的字体
            if line.startswith("Dialogue:"):
                def replace_inside(match):
                    text = match.group(1)
                    for sub, real in mapping.items():
                        if sub in text:
                            text = text.replace(sub, real)
                    return "{" + text + "}"

                new_line = braces_pattern.sub(replace_inside, line)
                new_lines.append(new_line)
                continue

            # 其他行保持原样
            new_lines.append(line)

        # 保存文件
        with open(filepath, "w", encoding="utf-8") as f:
            f.writelines(new_lines)
            return mapping

    def normalize_to_ascii(self, name: str) -> str:
        """生成符合 PostScript 名称的 ASCII 字符串"""
        name = ''.join(lazy_pinyin(name))  # 中文转拼音
        name = name.replace(' ', '_')
        name = re.sub(r'[^A-Za-z0-9_\-]', '', name)
        return name

    def replace_font_name_complete(self, font_path, old_name, new_name, output_path=None):
        if not os.path.exists(font_path):
            return ""

        try:
            font = TTFont(font_path)
            name_table = font['name']

            ascii_name = self.normalize_to_ascii(new_name)
            chinese_name = new_name

            supports_ttf_name = 'CFF ' not in font or 'glyf' in font

            # 临时存储所有修改后的名称
            name_records_dict = {}

            for record in name_table.names:
                try:
                    if record.platformID == 1:
                        decoded = record.string.decode('mac_roman', errors='ignore')
                    elif record.platformID == 3:
                        decoded = record.string.decode('utf-16be', errors='ignore')
                    else:
                        decoded = record.string.decode('utf-16be', errors='ignore')
                except Exception:
                    decoded = None

                if decoded and old_name in decoded:
                    if record.nameID == 6:
                        record.string = ascii_name.encode('utf-16be') if record.platformID != 1 else ascii_name.encode(
                            'mac_roman')
                    else:
                        if supports_ttf_name:
                            is_english = False
                            if record.platformID == 3:
                                is_english = record.langID in [0x0409, 0x0c09]
                            elif record.platformID == 1:
                                is_english = True
                            if is_english:
                                record.string = self.normalize_to_ascii(new_name).encode(
                                    'utf-16be') if record.platformID != 1 else self.normalize_to_ascii(new_name).encode(
                                    'mac_roman')
                            else:
                                record.string = chinese_name.encode(
                                    'utf-16be') if record.platformID != 1 else chinese_name.encode('mac_roman')
                        else:
                            record.string = chinese_name.encode(
                                'utf-16be') if record.platformID != 1 else chinese_name.encode('mac_roman')

                # 保存到临时字典
                name_records_dict.setdefault(record.nameID, {})[record.platformID] = record.string

            # 确保中文记录存在
            from fontTools.ttLib.tables._n_a_m_e import NameRecord
            for nameID in [1, 4, 16, 17]:
                exists = any(
                    r.nameID == nameID and
                    ((r.platformID == 3 and chinese_name in r.string.decode('utf-16be', errors='ignore')) or
                     (r.platformID == 1 and chinese_name in r.string.decode('mac_roman', errors='ignore')))
                    for r in name_table.names
                )
                if not exists:
                    new_record = NameRecord()
                    new_record.nameID = nameID
                    new_record.platformID = 3
                    new_record.platEncID = 1
                    new_record.langID = 0x0804
                    new_record.string = chinese_name.encode('utf-16be')
                    name_table.names.append(new_record)
                    # 更新临时字典
                    name_records_dict.setdefault(nameID, {})[3] = new_record.string

            if output_path is None:
                base, ext = os.path.splitext(font_path)
                output_path = f"{base}_replaced{ext}"

            font.save(output_path)
            font.close()

            # 保存到总表格
            self.font_name_registry[os.path.basename(font_path)] = name_records_dict

            return output_path

        except Exception as e:
            print("替换失败:", e)
            return ""

    def extract_fonts_from_video(self, video_path, workdir, mapping, seq_num):
        """
        使用 ffmpeg.exe 提取视频附件到工作目录，并重命名字体文件
        :param video_path: 视频路径
        :param workdir: 工作目录 Fonts/
        :param mapping: 子集名->原名映射
        :param seq_num: 当前视频序号
        :return: 视频字体目录路径
        """

        video_name = os.path.splitext(os.path.basename(video_path))[0]
        video_dir = os.path.join(workdir, f"{seq_num}_{video_name}")
        os.makedirs(video_dir, exist_ok=True)

        # === 1️⃣ 提取附件 ===
        ffmpeg_exe = self.get_ffmpeg_exe()
        # 切换到输出目录（因为 dump_attachment 会保存到当前工作目录）
        old_cwd = os.getcwd()
        os.chdir(video_dir)

        cmd = [
            ffmpeg_exe,
            "-dump_attachment:t", "",  # 提取所有附件
            "-i", video_path
        ]

        self.log(f"📦 提取附件：{video_name} -> {video_dir}")
        try:
            result = self.run_silently(cmd)
        except Exception as e:
            # 如果报错信息里包含“At least one output file must be specified”，忽略
            if "At least one output file must be specified" in str(e):
                self.log("⚠️ 忽略 ffmpeg 报错：附件已经提取")
            else:
                self.log(f"❌ 附件提取失败：{video_name} 错误: {e}")
        finally:
            os.chdir(old_cwd)

        # === 2️⃣ 重命名字体 ===
        renamed_count = 0
        for file in os.listdir(video_dir):
            file_path = os.path.join(video_dir, file)
            fname, ext = os.path.splitext(file)
            if ext.lower() not in [".ttf", ".otf"]:
                continue

            # 取“.”前半段（有的字体是 8A905FBC.XCJVKWC5.ttf）
            front_name = fname.split(".")[0].upper()

            for sub, real in mapping.items():
                if front_name == sub.upper():
                    new_file_path = os.path.join(video_dir, f"{real}{ext}")

                    # 先重命名文件
                    os.rename(file_path, new_file_path)

                    # 内部字体名替换：确保路径用原始字符串或用 / 分隔
                    normalized_path = new_file_path.replace("\\", "/")
                    #print(f"执行字体重命名参数：{normalized_path}，{front_name}，{real}，{normalized_path}")
                    self.replace_font_name_complete(normalized_path, front_name, real, output_path=normalized_path)

                    renamed_count += 1
                    break

        self.log(f"🔤 字体重命名完成：共 {renamed_count} 个字体文件")

        return video_dir

    def fix_name_table_with_records(self, font_path, name_records):
        """
        使用已记录的 name 表信息修复字体
        """
        font = TTFont(font_path)
        name_table = font['name']

        # 清空原有 name 表
        name_table.names.clear()

        from fontTools.ttLib.tables._n_a_m_e import NameRecord
        for nameID, platforms in name_records.items():
            for platformID, string in platforms.items():
                record = NameRecord()
                record.nameID = nameID
                record.platformID = platformID
                record.platEncID = 1
                record.langID = 0x0804
                record.string = string
                name_table.names.append(record)

        font.save(font_path)
        font.close()

    def _find_fontforge_executable(self):
        """
        查找fontforge可执行文件，按优先级：
        1. 项目目录下的 FontForge/bin/fontforge.exe
        2. 系统PATH中的fontforge
        3. 常见安装路径
        """
        # 优先级1：项目目录下的FontForge
        project_ff_path = self.base_dir / "FontForge" / "bin" / "fontforge.exe"
        if os.path.exists(project_ff_path):
            self.log("✅ 使用项目内的FontForge")
            return project_ff_path

        # 优先级2：当前工作目录下的FontForge
        cwd_ff_path = os.path.join(os.getcwd(), "FontForge", "bin", "fontforge.exe")
        if os.path.exists(cwd_ff_path):
            self.log("✅ 使用工作目录内的FontForge")
            return cwd_ff_path

        # 优先级3：系统PATH
        try:
            result = subprocess.run(['where', 'fontforge'], capture_output=True, text=True)
            if result.returncode == 0:
                ff_path = result.stdout.strip().split('\n')[0]
                self.log(f"✅ 使用系统PATH中的FontForge: {ff_path}")
                return ff_path
        except:
            pass

        # 优先级4：常见安装路径
        common_paths = [
            r"C:\Program Files\FontForgeBuilds\bin\fontforge.exe",
            r"C:\Program Files (x86)\FontForgeBuilds\bin\fontforge.exe",
            r"C:\Program Files (x86)\FontForgeBuilds\bin\fontforge.exe",
        ]

        for path in common_paths:
            if os.path.exists(path):
                self.log(f"✅ 使用常见路径的FontForge: {path}")
                return path

        self.log("❌ 未找到FontForge可执行文件")
        return None

    def _run_fontforge_script(self, script_path):
        ff_path = self._find_fontforge_executable()
        if not ff_path:
            self.log("❌ 无法找到FontForge，跳过字体合并")
            return False

        try:
            self.log(f"🔄 执行FontForge脚本: {os.path.basename(script_path)}")
            result = subprocess.run([
                ff_path, '-lang=py', '-script', script_path
            ], capture_output=True, text=True, timeout=600, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW)

            stdout = result.stdout.strip()
            stderr = result.stderr.strip()

            if stdout:
                self.log(f"📄 FontForge输出:\n{stdout}")
            if stderr:
                self.log(f"⚠️ FontForge警告/错误:\n{stderr}")

            # 使用 stdout 判断是否成功
            if "✅ 字体保存成功" in stdout or "=== 合并完成" in stdout:
                return True
            else:
                return False

        except subprocess.TimeoutExpired:
            self.log("❌ FontForge执行超时")
            return False
        except Exception as e:
            self.log(f"❌ FontForge执行异常: {e}")
            return False
        finally:
            try:
                os.unlink(script_path)
            except:
                pass

    def merge_fonts(self, workdir):
        """
        将 Fonts 子文件夹下的所有视频文件夹内的 TTF/OTF 字体合并到 Fonts 根目录。
        """
        if not self._find_fontforge_executable():
            self.log("❌ FontForge不可用，无法进行字体合并")
            return

        fonts_root = workdir
        # 清空根目录残留字体
        for file in os.listdir(fonts_root):
            file_path = os.path.join(fonts_root, file)
            if os.path.isfile(file_path):
                _, ext = os.path.splitext(file)
                if ext.lower() in [".ttf", ".otf"]:
                    try:
                        os.remove(file_path)
                        self.log(f"🗑 已删除残留字体文件: {file}")
                    except Exception as e:
                        self.log(f"❌ 删除文件失败: {file} 错误: {e}")

        font_groups = {}
        for subdir in os.listdir(workdir):
            subdir_path = os.path.join(workdir, subdir)
            if not os.path.isdir(subdir_path):
                continue
            for file in os.listdir(subdir_path):
                file_path = os.path.join(subdir_path, file)
                if not os.path.isfile(file_path):
                    continue
                _, ext = os.path.splitext(file)
                if ext.lower() not in [".ttf", ".otf"]:
                    continue
                base_name = os.path.splitext(file)[0]
                font_groups.setdefault(base_name, []).append(file_path)

        success_count = 0
        total_count = len(font_groups)

        for base_name, font_files in font_groups.items():
            if not font_files:
                continue

            try:
                if len(font_files) == 1:
                    # 单个字体直接复制
                    src = font_files[0]
                    dst = os.path.join(fonts_root, os.path.basename(src))
                    if os.path.abspath(src) != os.path.abspath(dst):
                        shutil.copy2(src, dst)

                    success_count += 1
                    self.log(f"✅ 复制字体: {base_name}")

                else:
                    # 多个字体合并
                    _, ext = os.path.splitext(font_files[0])
                    dst = os.path.join(fonts_root, f"{base_name}{ext}")
                    i = 1
                    while os.path.exists(dst):
                        dst = os.path.join(fonts_root, f"{base_name}_{i}{ext}")
                        i += 1

                    merge_script = self._create_fontforge_merge_script(font_files, dst)
                    if self._run_fontforge_script(merge_script):
                        font_basename = os.path.basename(font_files[0])
                        name_records = self.font_name_registry.get(font_basename, {})
                        if name_records:
                            self.fix_name_table_with_records(dst, name_records)
                        else:
                            familyname = base_name
                            fullname = base_name
                            postscriptname = base_name.replace(" ", "_")
                            self.fix_name_table(dst, familyname, fullname, postscriptname)

                        success_count += 1
                        self.log(f"✅ 合并字体并修复 name 表: {base_name} ({len(font_files)}个文件)")
                    else:
                        self.log(f"❌ 合并失败: {base_name}")

            except Exception as e:
                self.log(f"❌ 字体处理失败: {base_name} 错误: {e}")

        # --- 清空总表格 ---
        self.font_name_registry.clear()
        self.log("🧹 已清空 font_name_registry")
        self.log(f"🎨 字体处理完成: {success_count}/{total_count} 个字体组处理成功")

    def _create_fontforge_merge_script(self, font_files, output_path):
        """
        创建FontForge合并脚本（修正版）
        使用 importOutlines() 安全复制字形，避免 'glyph' 无 copy 方法错误
        """
        # 构建输入文件列表
        input_files_str = "[\n        " + ",\n        ".join([f'r"{f}"' for f in font_files]) + "\n    ]"

        # 脚本内容
        script_content = f'''# -*- coding: utf-8 -*-
import fontforge
import os
import sys
import tempfile
import traceback

def main():
    merged_font = fontforge.font()
    merged_font.encoding = 'UnicodeFull'
    total_glyphs = 0

    input_files = {input_files_str}
    output_file = r"{output_path}"

    print("=== 字体合并开始 ===")
    print(f"输入文件: {{input_files}}")
    print(f"输出文件: {{output_file}}")

    for i, font_path in enumerate(input_files):
        print(f"\\n--- 处理第 {{i+1}}/{{len(input_files)}} 个字体 ---")
        print(f"字体路径: {{font_path}}")

        if not os.path.exists(font_path):
            print(f"❌ 文件不存在: {{font_path}}")
            continue

        try:
            font = fontforge.open(font_path)
            all_glyphs = list(font.glyphs())
            print(f"✅ 打开字体成功: {{font.fontname}}")
            print(f"字体 {{font.fontname}} 共包含 {{len(all_glyphs)}} 个字形")

            if i == 0:
                try:
                    merged_font.fontname = font.fontname
                    merged_font.familyname = font.familyname
                    merged_font.fullname = font.fullname
                    print(f"设置字体元信息: {{font.fontname}} / {{font.familyname}}")
                except Exception as e:
                    print(f"⚠️ 设置字体元信息失败: {{e}}")

            glyph_count = 0
            for glyph in all_glyphs:
                name = glyph.glyphname
                if not glyph.isWorthOutputting() or name in merged_font:
                    continue
                try:
                    new_glyph = merged_font.createChar(glyph.encoding, name)
                    tmp_svg = tempfile.NamedTemporaryFile(delete=False, suffix=".svg")
                    tmp_svg.close()
                    glyph.export(tmp_svg.name)
                    new_glyph.importOutlines(tmp_svg.name)
                    new_glyph.width = glyph.width
                    os.unlink(tmp_svg.name)
                    glyph_count += 1
                    total_glyphs += 1
                except Exception as e:
                    print(f"⚠️ 复制字形失败: {{name}} -> {{e}}")

            print(f"✅ 完成字体 {{font.fontname}}，成功复制 {{glyph_count}} 个字形")
            font.close()

        except Exception as e:
            print(f"⚠️ 打开或处理字体失败: {{font_path}} -> {{e}}")
            traceback.print_exc()

    print(f"\\n=== 合并完成，共合并 {{total_glyphs}} 个字形 ===")

    if total_glyphs == 0:
        print("❌ 没有成功合并任何字形")
        return False
        
    # 🔧 修正字体元信息
    if not merged_font.fontname:
        merged_font.fontname = os.path.splitext(os.path.basename(output_file))[0].replace(" ", "_")

    if not merged_font.familyname:
        try:
            first_font = fontforge.open(input_files[0])
            merged_font.familyname = first_font.familyname or merged_font.fontname
            merged_font.fullname = first_font.fullname or merged_font.fontname
            first_font.close()
        except Exception:
            merged_font.familyname = merged_font.fontname
            merged_font.fullname = merged_font.fontname

    merged_font.sfnt_names = (
        ("English (US)", "Family", merged_font.familyname),
        ("English (US)", "Fullname", merged_font.fullname),
        ("English (US)", "PostScriptName", merged_font.fontname),
    )

    try:
        print("💾 保存合并后的字体...")
        ext = os.path.splitext(output_file.lower())[1]
        if ext == '.otf':
            merged_font.generate(output_file, flags=('opentype',))
        else:  # .ttf 或其他
            merged_font.generate(output_file)  # 不加 flags
        merged_font.close()
        print(f"✅ 字体保存成功: {{output_file}}")
        return True
    except Exception as e:
        print(f"❌ 字体保存失败: {{e}}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
'''

        # 创建临时脚本文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False, encoding='utf-8') as f:
            f.write(script_content)
            return f.name

    def log(self, msg):
        print(msg)  # 你可以改为写入GUI的文本框或者日志窗口

    def set_buttons_state(self, state):
        """统一设置所有按钮的状态"""
        self.import_btn.config(state=state)
        self.delete_btn.config(state=state)
        self.clear_btn.config(state=state)
        self.subfmt_cb.config(state=state)
        self.extract_btn.config(state=state)
        self.ass_fix_cb.config(state=state)
        self.font_mode_cb.config(state=state)
        #self.tree.config(state=state)

    def save_and_disable_buttons(self):
        """保存控件原始状态，并将它们全部禁用"""
        self._original_states = {
            'import_btn': self.import_btn['state'],
            'delete_btn': self.delete_btn['state'],
            'clear_btn': self.clear_btn['state'],
            'subfmt_cb': self.subfmt_cb['state'],
            'extract_btn': self.extract_btn['state'],
            'ass_fix_cb': self.ass_fix_cb['state'],
            'font_mode_cb': self.font_mode_cb['state'],
            #'tree': self.tree['state']
        }
        self.set_buttons_state('disabled')

    def restore_buttons_state(self):
        """将控件状态还原为保存的原始状态"""
        if not self._original_states:
            return  # 防止未保存直接还原

        self.import_btn.config(state=self._original_states.get('import_btn', 'normal'))
        self.delete_btn.config(state=self._original_states.get('delete_btn', 'normal'))
        self.clear_btn.config(state=self._original_states.get('clear_btn', 'normal'))
        self.subfmt_cb.config(state=self._original_states.get('subfmt_cb', 'normal'))
        self.extract_btn.config(state=self._original_states.get('extract_btn', 'normal'))
        self.ass_fix_cb.config(state=self._original_states.get('ass_fix_cb', 'normal'))
        self.font_mode_cb.config(state=self._original_states.get('font_mode_cb', 'normal'))
        #self.tree.config(state=self._original_states.get('tree', 'normal'))

    def toggle_all_selection(self):
        self.header_checked = not self.header_checked
        for idx, file_tuple in enumerate(self.files):
            file_info = dict(zip(self.file_fields, file_tuple))
            # 更新 checked 状态
            file_info["checked"] = self.header_checked
            # 保留其他字段
            self.files[idx] = tuple(file_info[field] for field in self.file_fields)
            self.tree.set(file_info["fullpath"], "Check", "☑" if self.header_checked else "☐")
        self.update_header_checkbox()

    def update_header_checkbox(self):
        if self.files and all(dict(zip(self.file_fields, f)).get("checked") for f in self.files):
            self.header_checked = True
        else:
            self.header_checked = False
        self.tree.heading(
            "Check",
            text="☑" if self.header_checked else "☐",
            command=self.toggle_all_selection
        )

    def get_selected_files(self):
        return [fullpath for checked, fullpath, _ in self.items.values() if checked]

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.items.clear()
        self.update_header_checkbox()
        self.refresh_tree()

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        if column != "#1":  # 只处理复选框列
            return

        item = self.tree.identify_row(event.y)
        if not item:
            return

        # 切换选中状态
        for i, file_tuple in enumerate(self.files):
            file_info = dict(zip(self.file_fields, file_tuple))
            fullpath = file_info.get("fullpath")
            if fullpath == item:
                # 切换 checked 状态
                file_info["checked"] = not file_info.get("checked", False)
                # 将 dict 转回 tuple，保持字段顺序
                self.files[i] = tuple(file_info[field] for field in self.file_fields)
                self.tree.set(item, "Check", "☑" if file_info["checked"] else "☐")
                break

        self.update_header_checkbox()
        return "break"  # 阻止默认行选中

    def enable_treeview_edit(self, tree):
        """允许指定 treeview 的单元格可编辑"""
        tree.bind("<Button-1>", lambda e, t=tree: self.on_tree_single_click_edit(e, t), add=True)

    def on_tree_single_click_edit(self, event, tree):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = tree.identify_column(event.x)
        row = tree.identify_row(event.y)
        if not row or column == "#1":  # 跳过复选框列
            return

        # 只有该行已被选中时才允许编辑
        if row not in tree.selection():
            return

        # ---- 关闭已有的编辑 Entry ----
        if hasattr(self, "_editing_entry") and self._editing_entry.winfo_exists():
            if hasattr(self, "_editing_entry_save_fn"):
                self._editing_entry_save_fn()

        # 获取列显示索引
        display_col_index = int(column.replace("#", "")) - 1

        # 显示列 -> file_fields 索引映射
        display_to_filefield_idx = {
            0: 2,  # Check -> checked
            1: 1,  # Filename -> filename
            2: 3,  # Width -> width
            3: 4,  # Height -> height
            4: 5,  # FPS -> fps
        }

        if display_col_index not in display_to_filefield_idx:
            return

        file_field_index = display_to_filefield_idx[display_col_index]
        field_name = self.file_fields[file_field_index]

        # ---- 禁止编辑文件名列 ----
        if field_name == "filename":
            return

        # ---- 创建新的 Entry ----
        x, y, width, height = tree.bbox(row, column)
        value = tree.set(row, column)

        entry = ttk.Entry(tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)

        # 🔑 延迟焦点设置，保证新 Entry 获得焦点
        tree.after(1, lambda e=entry: e.focus_force())

        # 保存当前 Entry 到实例变量，方便下一次点击关闭
        self._editing_entry = entry

        def save_edit(event=None):
            new_value = entry.get()
            tree.set(row, column, new_value)
            entry.destroy()

            # 更新 self.files
            for i, file_tuple in enumerate(self.files):
                file_info = dict(zip(self.file_fields, file_tuple))
                if file_info.get("fullpath") == row:
                    file_info[field_name] = new_value
                    self.files[i] = tuple(file_info.get(f, "") for f in self.file_fields)
                    break

            # 编辑 checked 列时刷新表头状态
            if field_name == "checked":
                self.update_header_checkbox()

            # 清理引用
            if hasattr(self, "_editing_entry"):
                del self._editing_entry
            if hasattr(self, "_editing_entry_save_fn"):
                del self._editing_entry_save_fn

        # 保存函数引用，方便下次点击前手动调用
        self._editing_entry_save_fn = save_edit

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def on_files_dropped(self, event):
        #self.set_buttons_state("disabled")  # 禁用按钮

        print(event.data)
        files_data = event.data

        files_data = files_data.replace(r'\{', '{').replace(r'\}', '}').replace(r'\ ', ' ')

        # Step 1: 处理外部的{}，将内部的空格替换为|
        # 这个正则表达式匹配外部的{}并替换其中的空格
        files_data = re.sub(r'\{([^{}]+)\}', lambda m: '{' + m.group(1).replace(' ', '|') + '}', files_data)
        print(files_data)

        # Step 2: 使用空格分隔文件路径
        files = files_data.split()

        final_files = []
        buffer = ""  # 用来暂时保存无法找到的文件路径

        for i, file in enumerate(files):
            file = file.strip()  # 去掉两端空格
            print(f"当前分段路径：{file}")
            is_renamed = False

            # Step 3: 对于被{}符号包裹的路径，恢复{}符号
            if file.startswith("{") and file.endswith("}"):
                file = file[1:-1]  # 删除开头和结尾的{}符号
                is_renamed = True

            file = file.replace('|', ' ')  # 替换所有的|符号为空格

            # 检查文件是否存在
            if os.path.isfile(file):
                print("分段文件路径存在")
                if buffer:  # 如果有缓冲路径，且当前路径存在，将缓冲路径和当前路径添加到final_files
                    if buffer.startswith("{") and buffer.endswith("}"):
                        buffer = buffer[1:-1]

                    final_files.append(f"'{buffer}'")
                    final_files.append(f"'{file}'")
                    buffer = ""  # 清空缓冲路径
                else:
                    final_files.append(f"'{file}'")
            else:
                if buffer:  # 如果当前路径不存在且buffer有路径，拼接它们
                    print("分段文件路径不存在正在拼接")
                    if is_renamed:
                        file = "{" + file + "}"  # 在这里给file的前后加上{}符号
                    buffer += " " + file  # 将当前路径加入到缓冲区
                    print(f"buffer内容：{buffer}")
                else:
                    print("分段文件路径不存在写入第一段")
                    if is_renamed:
                        file = "{" + file + "}"  # 在这里给file的前后加上{}符号
                    buffer = file  # 如果缓冲区为空，则将当前路径存入缓冲区
                    print(f"buffer内容：{buffer}")

        # 如果缓冲区最后仍有路径，表示最后一段路径是无效的，我们将它添加到final_files
        if buffer:
            final_files.append(f"'{buffer}'")

        # 打印最终文件路径列表（可选）
        print(", ".join(final_files))

        # Step 4: 在Treeview中添加文件
        t = threading.Thread(target=self.add_files, args=(final_files,), daemon=True)
        t.start()

        # 更新文件列表视图（仅显示当前文件名）
        #self.update_tree()

        # 启动新线程来计算 CRC32 值并更新新文件名
        #threading.Thread(target=self.process_files_in_background, daemon=True).start()
        #self.refresh_tree()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = SubtitleExtractorApp(root)
    root.mainloop()

# 打包指令：pyinstaller -F -w --collect-data=tkinterdnd2 .\.venv\Scripts\sub0_0_6.py