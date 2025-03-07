import ctypes
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import threading
import time
import random
import pyautogui
import win32gui
import win32con
from pynput import keyboard, mouse
import pygame
from pygame import mixer
"""
新增预览播放功能，暂缺游戏音频文件
"""
the_title = "卡拉彼丘琴房助手 v1.4.1 (25.3.7)"


class GlobalHotkey:
    """热键监控"""
    def __init__(self, play, pause, stop):
        self.hotkeys = None
        self.play = play
        self.pause = pause
        self.stop = stop
        self.keys = {'play': '<F10>',
                     'pause': '<F11>',
                     'stop': '<F12>'}
        self.running = True
        self.listener_thread = None
        self.start()

    def start(self):
        def listen():
            with keyboard.GlobalHotKeys({
                self.keys['play']: self.play,
                self.keys['pause']: self.pause,
                self.keys['stop']: self.stop
            }) as h:
                self.hotkeys = h
                while self.running:
                    time.sleep(0.001)
                    h.join(0.1)

        self.listener_thread = threading.Thread(target=listen, daemon=True)
        self.listener_thread.start()

    def stop_listener(self):
        self.running = False
        if self.hotkeys:
            self.hotkeys.stop()


class SheetEditor:
    """乐谱管理器界面"""
    def __init__(self, player):

        self.app = player
        self.edit_window = None
        self.listbox = None
        self.current_beat = 0.0
        self.key_buttons = []
        self.selected_index = -1  # 新增选中音符索引

        self.preview_playing = False  # 新增预览播放状态
        self.preview_thread = None    # 新增预览线程
        self.sound_blocks = {}  # 存储音阶与声音文件的映射
        self.load_files = False

        # 初始化音频系统
        pygame.init()
        mixer.init(frequency=44100, size=-16, channels=2, buffer=1024)

    def load_sound_files(self):
        """加载声音文件"""
        sound_dir = "sounds"  # 声音文件存放目录
        if not os.path.exists(sound_dir):
            print("未找到文件目录sounds")
            return
        # 支持多格式音频文件
        for file in os.listdir(sound_dir):
            if file.lower().endswith(('.wav', '.mp3', '.ogg')):
                try:
                    # 从文件名提取音阶编号（例如："1.wav" -> 1）
                    block_num = int(os.path.splitext(file)[0].strip('.'))
                    path = os.path.join(sound_dir, file)
                    self.sound_blocks[block_num] = mixer.Sound(path)
                except (ValueError, pygame.error) as e:
                    print(f"声音加载失败：{file} - {str(e)}")
        self.load_files = True

    def create_editor(self):
        """创建窗口"""
        if self.edit_window and self.edit_window.winfo_exists():
            self.edit_window.lift()
            return

        # 编辑器页面
        self.edit_window = tk.Toplevel(self.app.window)
        # 新增窗口关闭协议
        self.edit_window.protocol("WM_DELETE_WINDOW", self.on_editor_close)

        self.edit_window.title("乐谱编辑器")
        self.edit_window.geometry("800x750")

        # 置顶控制栏
        top_control = ttk.Frame(self.edit_window)
        top_control.pack(fill='x', padx=5, pady=2)

        # 节拍控制
        beat_frame = ttk.Frame(top_control)
        beat_frame.pack(side='left', padx=5)
        ttk.Label(beat_frame, text="当前节拍").pack(side='left')
        self.beat_entry = ttk.Entry(beat_frame, width=8)
        self.beat_entry.insert(0, "0.25")
        self.beat_entry.pack(side='left', padx=2)
        ttk.Button(beat_frame, text="半拍", width=5,
                   command=lambda: self.adjust_beat(0.125)).pack(side='left')
        ttk.Button(beat_frame, text="一拍", width=5,
                   command=lambda: self.adjust_beat(0.25)).pack(side='left')
        ttk.Button(beat_frame, text="空半拍", width=8,
                   command=lambda: self.adjust_beat(-0.125)).pack(side='left')
        ttk.Button(beat_frame, text="空一拍", width=8,
                   command=lambda: self.adjust_beat(-0.25)).pack(side='left')

        ttk.Checkbutton(top_control, text="窗口置顶",
                        command=lambda: self.edit_window.attributes('-topmost',
                        not self.edit_window.attributes('-topmost'))).pack(side='right')

        # 主内容区
        main_frame = ttk.Frame(self.edit_window)
        main_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 左侧按钮矩阵
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side='left', fill='y', padx=5)

        matrix_frame = ttk.LabelFrame(left_frame, text="音阶矩阵 (1-16)")
        matrix_frame.pack(pady=5)

        # 创建一个自定义样式来调整按钮的高度
        style = ttk.Style()
        style.configure("Tall.TButton", padding=(10, 20))  # 调整 padding 来控制高度

        for row in range(4):
            frame_row = ttk.Frame(matrix_frame)
            frame_row.pack()
            for col in range(4):
                btn_num = (3 - row) * 4 + col + 1  # 修正矩阵布局
                btn = ttk.Button(frame_row, text=str(btn_num), width=5,
                                 style="Tall.TButton",  # 应用自定义样式
                                 command=lambda b=btn_num: self.add_by_button(b))
                btn.pack(side='left', padx=5, pady=5)
                self.key_buttons.append(btn)

        # 右侧音符列表
        list_frame = ttk.LabelFrame(main_frame, text="乐谱列表")
        list_frame.pack(side='right', fill='both', expand=True, padx=5)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')

        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set,
                                  selectmode='single', width=30)
        self.listbox.pack(fill='both', expand=True)
        scrollbar.config(command=self.listbox.yview)

        # 绑定选择事件
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        # 底部控制按钮
        control_frame = ttk.Frame(self.edit_window)
        control_frame.pack(fill='x', padx=5, pady=5)
        # 插入音节
        self.note_entry = ttk.Entry(control_frame, width=4)
        self.note_entry.pack(side='left', padx=2)
        self.note_entry.insert(0, "")
        # 控制按钮
        ttk.Button(control_frame, text="插入(向前)", command=self.insert_note, width=10).pack(side='left', padx=2)
        ttk.Button(control_frame, text="删除(当前)", command=self.delete_note, width=10).pack(side='left', padx=2)
        ttk.Button(control_frame, text="保存乐谱", command=self.save_sheet, width=8).pack(side='right', padx=2)
        ttk.Button(control_frame, text="停止预览", command=self.stop_preview, width=8).pack(side='right', padx=2)  # 新增停止按钮
        ttk.Button(control_frame, text="播放试听", command=self.play_preview, width=8).pack(side='right', padx=2)

        self.refresh_list()

    def add_by_button(self, block):
        """通过按键插入音节"""
        try:
            beat = float(self.beat_entry.get())
            # beat负数，空音节置空
            if beat < 0:
                block = 0
            self.app.sheet_data.append({'beat': beat, 'block': block})
            self.app.refresh_sheet_display()
        except ValueError:
            messagebox.showerror("错误", "未知错误，获取音节失败")

    def adjust_beat(self, delta):
        """调整节拍值"""
        try:
            self.beat_entry.delete(0, tk.END)
            self.beat_entry.insert(0, delta)

        except ValueError:
            self.beat_entry.delete(0, tk.END)
            self.beat_entry.insert(0, "0")

    def refresh_list(self):
        """刷新编辑器乐谱列表"""
        if self.edit_window and self.edit_window.winfo_exists():
            self.listbox.delete(0, tk.END)
            for note in self.app.sheet_data:
                self.listbox.insert(tk.END, f"节拍: {note['beat']:.3f} | 方块: {note['block']}")

    def on_select(self, event):
        """处理列表选择事件"""
        selection = self.listbox.curselection()
        if selection:
            self.selected_index = selection[0]

    def insert_note(self):
        """在选中位置前插入新音符"""
        try:
            beat = float(self.beat_entry.get())
            block = int(self.note_entry.get())
            if self.selected_index >= 0:
                if 0 <= block <= 16:
                    # 自动纠错 0按键只能为负，非0按键只能为正
                    if block == 0: beat = -abs(beat)
                    else: beat = abs(beat)
                    new_note = {'beat': beat, 'block': block}
                    insert_index = self.selected_index
                    self.app.sheet_data.insert(insert_index, new_note)
                    self.selected_index = -1
                    self.refresh_list()
                    self.app.refresh_sheet_display()
                else:
                    messagebox.showerror("错误", "不存在的音节")
            else:
                messagebox.showerror("错误", "未选择右侧音节")
        except ValueError:
            messagebox.showerror("错误", "获取音节失败")

    def delete_note(self):
        """删除选中音符"""
        if self.selected_index >= 0 and hasattr(self.app, 'sheet_data'):
            if len(self.app.sheet_data) > self.selected_index:
                del self.app.sheet_data[self.selected_index]
                self.refresh_list()
                self.app.refresh_sheet_display()
                self.selected_index = -1

    def play_preview(self):
        """新增预览播放方法"""
        if not self.load_files:
            self.load_sound_files()
            self.app.update_status("载入预览按键音频成功")# 初始化时加载声音文件
        if self.preview_playing:
            return
        try:
            # 获取预览参数
            bpm = int(self.app.bpm_entry.get())
            notes = self.app.sheet_data
            if not notes:
                messagebox.showinfo("提示", "乐谱为空")
                return
        except ValueError:
            messagebox.showerror("错误", "无效的BPM值")
            return

        self.preview_playing = True
        self.preview_thread = threading.Thread(
            target=self.run_preview,
            args=(bpm, notes),
            daemon=True
        )
        self.preview_thread.start()

    def run_preview(self, bpm, notes):
        """新增预览播放核心逻辑"""
        base_delay = 60 / bpm  # 每拍基础时间
        channel = mixer.Channel(0)  # 使用独立音频通道
        try:
            for idx, note in enumerate(notes):
                if not self.preview_playing:
                    break

                # 更新界面显示
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
                self.edit_window.update()

                # 节拍非空播放对应声音
                beat, block = note['beat'], note['block']
                if beat > 0 and block != 0:
                    if block in self.sound_blocks:
                        try:
                            if channel.get_busy():  # 停止前一个音
                                channel.stop()
                            channel.play(self.sound_blocks[block])
                        except pygame.error as e:
                            print(f"播放失败：{str(e)}")
                # 计算节拍等待时间（考虑正负节拍）
                wait_time = abs(beat) * base_delay
                time.sleep(wait_time)

                # 更新主界面高亮
                self.app.highlight_note(idx)
                self.app.sheet_canvas.xview_moveto(idx / len(notes))
        finally:
            channel.stop()
            self.preview_playing = False
            self.listbox.selection_clear(0, tk.END)
            self.app.highlight_note(-1)  # 清除高亮

    def stop_preview(self):
        """新增停止预览方法"""
        self.preview_playing = False
        if self.preview_thread and self.preview_thread.is_alive():
            self.preview_thread.join(0.5)

    def new_sheet(self):
        """新建乐谱"""
        self.app.new_sheet()
        self.refresh_list()
        self.beat_entry.delete(0, tk.END)
        self.beat_entry.insert(0, "0")

    def save_sheet(self):
        """保存乐谱"""
        self.app.save_sheet()

    def on_editor_close(self):
        """处理编辑器窗口关闭事件"""
        self.stop_preview()  # 关闭时停止预览
        if self.edit_window:
            self.edit_window.destroy()
        self.edit_window = None  # 清除窗口引用
        self.listbox = None  # 清除listbox引用


class MusicAutoPlayer:
    """控制器主界面"""
    def __init__(self):
        self.window = self.create_window()
        self.sheet_editor = SheetEditor(self)  # 新增编辑器实例
        self.sheet_data = []  # 新增乐谱数据初始化
        self.current_file = None  # 新增当前文件路径存储
        self.init_ui()  # UI初始化必须在编辑器之后
        self.state = {
            'playing': False,
            'paused': False,
            'hwnd': None,  # 匹配窗口的句柄
            'rect': None,  # 匹配窗口尺寸大小
            'coordinate': [None, None],  # 存储窗口坐标
            'blocks': [(0, 0)] * 16,  # 16个区块的状态
            'current_note': -1,  # 表示当前正在处理或播放的音符
            'hotkeys': None,  # 存储热键配置信息
            'bpm': 60,  # 演奏BPM
            'mouse': 10  # 鼠标抖动（模拟人类鼠标抖动，还没有在程序中实现）
        }
        self.note_labels = {'beat': [], 'block': []}  # 新增初始化
        self.setup_listeners()
        self.check_window_active()

    def create_window(self):
        """1、创建主窗口"""
        window = tk.Tk()
        window.title(the_title)
        window.geometry("700x750")
        window.columnconfigure(0, weight=1)
        return window

    def init_ui(self):
        """2、初始化界面组件"""
        # 总控制面板
        control_frame = ttk.LabelFrame(self.window, text="控制面板")
        control_frame.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')
        # 第一行 窗口捕捉和置顶按钮
        ttk.Button(control_frame, text="捕捉窗口", command=self.capture_window).grid(row=0, column=0, padx=5)
        ttk.Checkbutton(control_frame, text="窗口置顶", command=self.toggle_topmost).grid(row=0, column=1, padx=5)
        self.status_label = ttk.Label(control_frame, text="就绪", foreground='gray')
        self.status_label.grid(row=0, column=2, padx=10, sticky='nsew')

        # 第二行 校准面板
        cal_frame = ttk.LabelFrame(self.window, text="坐标校准")
        cal_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')
        ttk.Button(cal_frame, text="定位左上", command=lambda: self.get_coordinate(0)).grid(row=0, column=0)
        ttk.Button(cal_frame, text="定位右下", command=lambda: self.get_coordinate(1)).grid(row=0, column=1)
        self.cal_labels = [
            ttk.Label(cal_frame, text="未设置", foreground='darkgray'),
            ttk.Label(cal_frame, text="未设置", foreground='darkgray')
        ]
        self.cal_labels[0].grid(row=0, column=3, padx=5)
        self.cal_labels[1].grid(row=0, column=4, padx=5)

        # 第三行 坐标矩阵
        grid_frame = ttk.LabelFrame(self.window, text="校准坐标")
        grid_frame.grid(row=2, column=0, padx=10, pady=5)

        self.grid_labels = []
        for i in range(16):
            lbl = ttk.Label(grid_frame, text="(0,0)", width=10, relief='ridge')
            lbl.grid(row=i // 4, column=i % 4, padx=2, pady=2)
            self.grid_labels.append(lbl)

        # 第四行 播放控制
        play_frame = ttk.LabelFrame(self.window, text="演奏控制")
        play_frame.grid(row=3, column=0, padx=10, pady=5, sticky='nsew')

        control_play_frame = ttk.Frame(play_frame)
        control_play_frame.pack(pady=0)

        self.start_button = ttk.Button(control_play_frame, text="开始 (F10)", command=self.start_playing)
        self.start_button.grid(row=0, column=0, sticky='nsew')

        self.pause_button = ttk.Button(control_play_frame, text="⏸ 暂停 (F11)", command=self.toggle_pause)
        self.pause_button.grid(row=0, column=1, sticky='nsew')

        self.stop_button = ttk.Button(control_play_frame, text="■ 停止 (F12)", command=self.stop_playing)
        self.stop_button.grid(row=0, column=2, sticky='nsew')

        # 第五行 演奏设置
        play_setting_frame = ttk.LabelFrame(self.window, text="演奏设置")
        play_setting_frame.grid(row=4, column=0, padx=10, pady=5, sticky='nsew')
        # BPM设置
        ttk.Label(play_setting_frame, text="BPM速度").grid(row=0, column=0, sticky='nsew')
        self.bpm_entry = ttk.Entry(play_setting_frame, width=6)
        self.bpm_entry.insert(0, "60")
        self.bpm_entry.grid(row=0, column=1, padx=5, sticky='nsew')
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=6).grid(row=0, column=2, sticky='nsew')
        # 抖动设置
        ttk.Label(play_setting_frame, text="鼠标抖动").grid(row=0, column=3, sticky='e')
        self.mouse_move = ttk.Entry(play_setting_frame, width=6)
        self.mouse_move.insert(0, "10")
        self.mouse_move.grid(row=0, column=4, padx=5)
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=6).grid(row=0, column=5, sticky='e')

        # 第六行 乐谱控制
        sheet_frame = ttk.LabelFrame(self.window, text="乐谱管理")
        sheet_frame.grid(row=5, column=0, padx=10, pady=5, sticky='nsew')

        control_btn_frame = ttk.Frame(sheet_frame)
        control_btn_frame.pack(pady=5)

        ttk.Button(control_btn_frame, text="🎹 打开编辑器", command=self.sheet_editor.create_editor).pack(side='left', padx=5)
        ttk.Button(control_btn_frame, text="加载乐谱", command=self.load_sheet).pack(side='left', padx=5)
        ttk.Button(control_btn_frame, text="清空乐谱", command=self.new_sheet).pack(side='left', padx=5)

        # 左侧固定行名
        left_header = ttk.Frame(sheet_frame)
        left_header.pack(side='left', fill='y')
        ttk.Label(left_header, text="节拍", width=4, relief="raised").grid(row=0, column=0, padx=5, pady=2)
        ttk.Label(left_header, text="方块", width=4, relief="raised").grid(row=1, column=0, padx=5, pady=2)

        # 右侧可滚动区域
        right_canvas_frame = ttk.Frame(sheet_frame)
        right_canvas_frame.pack(side='right', expand=True, fill='both')

        self.sheet_canvas = tk.Canvas(right_canvas_frame, height=80)
        scroll_x = ttk.Scrollbar(right_canvas_frame, orient="horizontal", command=self.sheet_canvas.xview)
        scroll_x.pack(side="bottom", fill="x")
        self.sheet_canvas.pack(side="top", fill="both", expand=True)

        self.sheet_table = ttk.Frame(self.sheet_canvas)
        self.sheet_canvas.create_window((0, 0), window=self.sheet_table, anchor="nw")

        self.sheet_table.bind("<Configure>", lambda e: self.sheet_canvas.configure(
            scrollregion=self.sheet_canvas.bbox("all"),
            xscrollcommand=scroll_x.set))

    def update_status(self, message, color='gray'):
        """3、更新状态栏"""
        self.status_label.config(text=message, foreground=color)

    def setup_listeners(self):
        """4、设置事件监听"""
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        self.state['hotkeys'] = GlobalHotkey(self.start_playing, self.toggle_pause, self.stop_playing)

    def check_window_active(self):
        """5-1、定时检测窗口激活状态"""
        if self.state['playing'] and not self.is_window_active():
            self.toggle_pause()
            self.update_status("窗口未激活，自动暂停", 'orange')
        self.window.after(1000, self.check_window_active)

    def is_window_active(self):
        """5-2、检测游戏窗口是否激活"""
        try:
            active_hwnd = win32gui.GetForegroundWindow()
            return active_hwnd == self.state['hwnd']
        except:
            return False

    """-----------------以下为实际功能-----------------"""

    def toggle_topmost(self):
        """1、切换置顶状态"""
        current = self.window.attributes('-topmost')
        self.window.attributes('-topmost', not current)

    def capture_window(self):
        """2、捕捉游戏窗口"""

        def on_click(x, y, button, pressed):
            if pressed and button == mouse.Button.left:
                hwnd = win32gui.WindowFromPoint((x, y))
                root_hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
                if win32gui.IsWindowVisible(root_hwnd):
                    title = win32gui.GetWindowText(root_hwnd)
                    rect = win32gui.GetWindowRect(root_hwnd)
                    self.state.update(hwnd=root_hwnd, rect=rect)
                    self.update_status(f"已捕捉：{title}", 'green')
                    return False

        self.update_status("请点击游戏窗口任意位置...", 'blue')
        mouse.Listener(on_click=on_click).start()

    def get_coordinate(self, corner):
        """3-1、开始坐标校准"""
        if not self.state['hwnd']:
            self.update_status("请先捕捉窗口", 'red')
            return

        def on_click(x, y, button, pressed):
            if pressed and button == mouse.Button.left:
                self.state['coordinate'][corner] = (x, y)
                self.cal_labels[corner].config(text=f"({x}, {y})", foreground='green')
                self.calculate_blocks()
                return False

        self.update_status(f"请点击{'左上' if corner == 0 else '右下'}角...", 'blue')
        mouse.Listener(on_click=on_click).start()

    def calculate_blocks(self):
        """3-2、计算方块坐标"""
        if None in self.state['coordinate']: return
        left, top = self.state['coordinate'][0]
        right, bottom = self.state['coordinate'][1]
        w, h = (right - left) / 4, (bottom - top) / 4
        for i in range(16):
            x = left + (i % 4) * w + w / 2
            y = top + (i // 4) * h + h / 2
            self.state['blocks'][i] = (x, y)
            self.grid_labels[i].config(text=f"({int(x)}, {int(y)})")
        self.update_status("坐标已更新", 'green')

    def load_sheet(self):
        """4、加载乐谱"""
        path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 严格数据验证
                if not isinstance(data, dict):
                    self.update_status("格式存在问题", 'green')
                if 'bpm' not in data or not isinstance(data['bpm'], (int, float)):
                    self.update_status("缺少bpm", 'green')
                if 'notes' not in data or not isinstance(data['notes'], list):
                    self.update_status("缺少notes", 'green')
                for note in data['notes']:
                    if 'beat' not in note or 'block' not in note:
                        self.update_status("音节数据缺少beat或block")
                    if not (0 <= note['block'] <= 16):
                        self.update_status("方块编号必须在0-16之间")

                self.state['bpm'] = data['bpm']  # 新增状态更新
                self.sheet_data = data['notes']  # 新增数据同步
                self.bpm_entry.delete(0, tk.END)
                self.bpm_entry.insert(0, str(data['bpm']))
                self.refresh_sheet_display()  # 调用主界面刷新方法
                self.sheet_editor.refresh_list()  # 主动刷新编辑器

            # 读取提示
            self.update_status(f"成功加载乐谱: {os.path.basename(path)}", 'green')
        except Exception as e:
            self.update_status(f"加载失败: {str(e)}", 'red')

    def start_playing(self):
        """4-1、开始演奏"""
        if not hasattr(self, 'sheet_data'):
            self.update_status("请先加载乐谱", 'red')
            return
        if self.state['playing']:  # 防止重复启动
            return
        for i in [3, 2, 1]:
            self.update_status(f"{i},将鼠标移入游戏窗口并点击", 'red')
            time.sleep(1)
        self.state['playing'] = True
        self.state['paused'] = False
        self.update_status("开始演奏，请勿移动鼠标", 'red')
        threading.Thread(target=self.play_notes, daemon=True).start()

    def toggle_pause(self):
        """4-2、切换暂停状态"""
        if not self.state['playing']: return
        self.state['paused'] = not self.state['paused']
        status = "已暂停" if self.state['paused'] else "继续演奏"
        self.update_status(status, 'orange')

        # 根据状态更新按钮文本
        if self.state['paused']:
            self.pause_button.config(text="▶ 继续 (F11)")  # 暂停时，按钮文本改为“继续”
        else:
            self.pause_button.config(text="⏸ 暂停 (F11)")  # 继续时，按钮文本改为“暂停”

    def stop_playing(self):
        """4-3、停止演奏"""
        self.state['playing'] = False
        self.state['paused'] = False
        self.highlight_note(-1)
        self.sheet_canvas.xview_moveto(0)  # 新增滚动重置
        self.update_status("演奏已停止", 'gray')

    def update_bpm(self):
        """5-1、更新BPM"""
        try:
            bpm = int(self.bpm_entry.get())
            self.state['bpm'] = bpm
            self.update_status(f"BPM已修改为{bpm}", 'green')
        except ValueError:
            self.update_status("BPM必须为整数", 'red')

    def update_mouse(self):
        """5-2、更新鼠标抖动"""
        try:
            mouse_move = int(self.mouse_move.get())
            self.state['mouse'] = mouse_move
            self.update_status(f"鼠标抖动已修改{mouse_move}", 'green')

        except ValueError:
            self.update_status("抖动必须为整数", 'red')

    def play_notes(self):
        """6-1、演奏核心逻辑"""
        start_time = time.time()
        # 键位映射：实际键位编号 → 原索引
        index_map = [
            12, 13, 14, 15,  # 坐标存储索引0-3 → 原索引12-15（对应13-16）
            8, 9, 10, 11,  # 坐标存储4-7 → 原索引8-11（对应9-12）
            4, 5, 6, 7,  # 坐标存储8-11 → 原索引4-7（对应5-8）
            0, 1, 2, 3  # 坐标存储12-15 → 原索引0-3（对应1-4）
        ]

        try:
            bpm = int(self.state['bpm'])
            print(bpm)
            delay = 60 / bpm

            for idx, note in enumerate(self.sheet_data):
                if not self.state['playing']: break
                # 处理暂停和窗口激活检测
                while True:
                    if not self.state['playing']: return
                    if not self.state['paused'] and self.is_window_active():
                        break

                # 节拍同步  记录空节拍时间
                target_time = abs(note['beat']) * delay
                while time.time() - start_time < target_time:
                    if not self.state['playing'] or self.state['paused']: break

                # 界面更新
                self.window.after(0, self.highlight_note, idx)
                # 非空节拍才演奏
                if note['beat'] > 0:
                    self.window.after(0, lambda: self.sheet_canvas.xview_moveto(idx / len(self.sheet_data)))
                    # 取得索引
                    block = note['block']-1
                    mouse_shift = self.state.get('mouse')
                    # 修改索引映射逻辑
                    original_index = index_map[block]
                    x, y = self.state['blocks'][original_index]
                    # 坐标移动模拟点击
                    x += random.gauss(0, mouse_shift)
                    y += random.gauss(0, mouse_shift)
                    pyautogui.moveTo(x, y, duration=random.uniform(target_time*0.85, target_time*0.95))
                    pyautogui.click()
                else:
                    # 空节拍延时
                    time.sleep(random.uniform(target_time*0.95, target_time*1.05))
        except Exception as e:
            self.update_status(f"演奏出错：{str(e)}", 'red')
        finally:
            self.stop_playing()

    def highlight_note(self, idx):
        """6-2、高亮当前音符"""
        # 清除旧高亮
        if self.state['current_note'] >= 0:
            for lbl in self.note_labels.values():
                try:
                    lbl[self.state['current_note']].config(background='')
                except:
                    pass
        # 设置新高亮
        if idx >= 0:
            for lbl in self.note_labels.values():
                try:
                    lbl[idx].config(background='#FFF3CD')
                except:
                    pass
            self.state['current_note'] = idx

    def refresh_sheet_display(self):
        """刷新主界面乐谱显示"""
        for col in self.note_labels.values():
            for widget in col:
                widget.destroy()
        self.note_labels['beat'].clear()
        self.note_labels['block'].clear()

        # 重新生成显示
        if hasattr(self, 'sheet_data'):
            for col_idx, note in enumerate(self.sheet_data):
                beat_lbl = ttk.Label(self.sheet_table, text=f"{note['beat']:.3f}", width=6)
                beat_lbl.grid(row=0, column=col_idx, padx=2, pady=2)
                self.note_labels['beat'].append(beat_lbl)
                block_lbl = ttk.Label(self.sheet_table, text=note['block'], width=6)
                block_lbl.grid(row=1, column=col_idx, padx=2, pady=2)
                self.note_labels['block'].append(block_lbl)

        # 主界面乐谱移动到开头
        self.sheet_canvas.xview_moveto(0)

        # 判断编辑器窗口是否存在后再刷新
        if self.sheet_editor.edit_window:
            try:
                if self.sheet_editor.edit_window.winfo_exists():
                    self.sheet_editor.refresh_list()
            except Exception as e:
                print(e)

    def new_sheet(self):
        """清空乐谱"""
        self.sheet_data = []
        if hasattr(self, 'current_file'):
            del self.current_file
        self.refresh_sheet_display()
        self.update_status("已清空", 'green')

    def save_sheet(self):
        """保存乐谱"""
        if not hasattr(self, 'sheet_data'):
            self.update_status("没有可保存的乐谱数据", 'red')
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")]
        )
        if path:
            data = {
                "bpm": self.state['bpm'],
                "notes": self.sheet_data
            }
            try:
                with open(path, 'w') as f:
                    json.dump(data, f, indent=2)
                self.current_file = path
                self.update_status(f"乐谱已保存至: {path}", 'green')
            except Exception as e:
                self.update_status(f"保存失败: {str(e)}", 'red')

    def on_close(self):
        if self.state['hotkeys']:
            self.state['hotkeys'].stop_listener()
        self.state['playing'] = False
        self.window.destroy()


if __name__ == "__main__":
    app = MusicAutoPlayer()
    app.window.mainloop()
