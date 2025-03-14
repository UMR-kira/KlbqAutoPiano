import ctypes
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import threading
import time
import random

import cv2
import numpy as np
import pydirectinput
import pyautogui
import win32gui
import win32con
from pynput import keyboard, mouse
import pygame
from pygame import mixer
"""
尝试采用自动识别琴键
"""
the_title = "卡拉彼丘琴房助手 v1.5.3cv (25.3.14)"


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
        self.selected_index = -1

        self.preview_playing = False
        self.preview_thread = None
        self.sound_blocks = {}
        self.load_files_status = False
        # 新增编辑状态变量
        self.editing = False
        self.edit_entry = None
        self.edit_index = -1

        # 初始化音频系统
        pygame.init()
        mixer.init(frequency=44100, size=-16, channels=2, buffer=1024)

    def load_sound_files(self):
        """加载声音文件"""
        sound_dir = "sounds"
        if not os.path.exists(sound_dir):
            messagebox.showinfo("提示", "未发现音频目录")
            return
        # 支持多格式音频文件
        sound_files = os.listdir(sound_dir)
        if len(sound_files) == 0:
            messagebox.showinfo("错误", "文件夹为空")
            self.app.update_status("载入音频失败")
        for file in sound_files:
            if file.lower().endswith(('.wav', '.mp3', '.ogg')):
                try:
                    block_num = int(os.path.splitext(file)[0].strip('.'))
                    path = os.path.join(sound_dir, file)
                    self.sound_blocks[block_num] = mixer.Sound(path)
                except:
                    messagebox.showinfo("错误", "读取音频失败")
                    self.app.update_status("载入音频失败")
                    return
        if self.sound_blocks.__len__() == 16:
            self.load_files_status = True
            self.app.update_status("载入音频成功")
        else:
            messagebox.showinfo("错误", "音频文件数量错误")
            self.app.update_status("载入音频失败")
            return

    def create_editor(self):
        """创建窗口"""
        if self.edit_window and self.edit_window.winfo_exists():
            self.edit_window.lift()
            return

        # 编辑器页面
        self.edit_window = tk.Toplevel(self.app.window)
        # 窗口关闭协议
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
        self.beat_entry = ttk.Entry(beat_frame, width=4)
        self.beat_entry.insert(0, "1.0")
        self.beat_entry.pack(side='left', padx=2)
        ttk.Button(beat_frame, text="半拍", width=5,
                   command=lambda: self.adjust_beat(0.5)).pack(side='left')
        ttk.Button(beat_frame, text="一拍", width=5,
                   command=lambda: self.adjust_beat(1.0)).pack(side='left')
        ttk.Button(beat_frame, text="换行空拍", width=8,
                   command=lambda: self.add_blank(-2.0)).pack(side='right')
        ttk.Button(beat_frame, text="空一拍", width=6,
                   command=lambda: self.add_blank(-1.0)).pack(side='right')
        ttk.Button(beat_frame, text="空半拍", width=6,
                   command=lambda: self.add_blank(-0.5)).pack(side='right')

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

        # 在创建 listbox 后添加双击事件绑定
        self.listbox.bind("<Double-Button-1>", self.start_edit)  # 新增双击绑定

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

    def start_edit(self, event):
        """启动编辑模式"""
        if self.editing:  # 防止重复编辑
            return
        # 获取点击位置索引
        index = self.listbox.nearest(event.y)
        if index < 0: return
        # 获取原数据
        original = self.app.sheet_data[index]
        text = f"{original['beat']:.2f} | {original['block']}"

        # 计算输入框位置
        x, y, _, h = self.listbox.bbox(index)
        y += self.listbox.winfo_y()

        # 创建输入框
        self.edit_entry = ttk.Entry(self.listbox, width=20)
        self.edit_entry.place(x=x, y=y, width=150, height=h)
        self.edit_entry.insert(0, text)
        self.edit_entry.focus()

        # 绑定事件
        self.edit_entry.bind("<Return>", lambda e: self.finish_edit(index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.finish_edit(index))
        self.edit_entry.bind("<Escape>", lambda e: self.cancel_edit())

        self.editing = True
        self.edit_index = index

    def finish_edit(self, index):
        """完成编辑"""
        if not self.editing: return
        try:
            # 解析输入内容
            text = self.edit_entry.get()
            beat, block = text.split("|")
            beat = float(beat.strip())
            block = int(block.strip())
            # 数据校验
            if not (0 <= block <= 16):
                raise ValueError("方块编号必须为0-16")
            # 更新数据
            self.app.sheet_data[index] = {
                'beat': -abs(beat) if block == 0 else abs(beat),  # 自动处理正负节拍
                'block': block
            }
            # 刷新显示
            self.refresh_list()
            self.app.refresh_sheet_display()
        except Exception as e:
            messagebox.showerror("输入错误", f"无效格式: {str(e)}")

        self.cancel_edit()

    def cancel_edit(self):
        """取消编辑"""
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
        self.editing = False
        self.edit_index = -1

    def add_by_button(self, block):
        """通过按键插入音节"""
        try:
            beat = float(self.beat_entry.get())
            # beat负数，空音节置空
            if beat < 0:
                block = 0
            self.app.sheet_data.append({'beat': beat, 'block': block})
            self.refresh_list()
            self.listbox.see(tk.END)  # 确保新条目可见
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

    def add_blank(self, delta):
        """增加空白节拍值"""
        try:
            self.beat_entry.delete(0, tk.END)
            self.beat_entry.insert(0, delta)
            self.add_by_button(0)
            self.listbox.yview_moveto(self.selected_index / self.app.sheet_data.__len__())
        except ValueError:
            self.beat_entry.delete(0, tk.END)
            self.beat_entry.insert(0, "")

    def refresh_list(self):
        """刷新编辑器乐谱列表（保持滚动位置）"""
        if self.edit_window and self.edit_window.winfo_exists():
            # 保存当前滚动位置和选中状态
            scroll_pos = self.listbox.yview()
            selected = self.listbox.curselection()
            # 刷新列表内容
            self.listbox.delete(0, tk.END)
            for note in self.app.sheet_data:
                self.listbox.insert(tk.END, f"节拍: {note['beat']:.2f} | 方块: {note['block']}")
            # 恢复滚动位置
            self.listbox.yview_moveto(scroll_pos[0])
            # 恢复选中状态（如果存在）
            if selected:
                try:
                    self.listbox.selection_set(selected[0])
                except tk.TclError:
                    pass
            # 如果之前正在编辑，重新定位
            if self.editing and self.edit_index >= 0:
                self.listbox.selection_set(self.edit_index)
                self.listbox.see(self.edit_index)

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
                    self.refresh_list()
                    # 插入位置保持选中
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(insert_index)
                    self.listbox.see(insert_index)
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
                # 删除后自动选中下一个条目
                new_index = min(self.selected_index, len(self.app.sheet_data) - 1)
                if new_index >= 0:
                    self.listbox.selection_set(new_index)
                    self.listbox.see(new_index)

    def play_preview(self):
        """预览播放方法"""
        if self.preview_playing:    # 防止重复点击播放
            return
        if not self.load_files_status:
            self.load_sound_files()
        # 读取成功才能预览播放
        if self.load_files_status:
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
            index = self.selected_index
            self.preview_playing = True
            self.preview_thread = threading.Thread(
                target=self.run_preview,
                args=(bpm, notes, index),
                daemon=True
            )
            self.preview_thread.start()

    def run_preview(self, bpm, notes, index):
        """新增预览播放核心逻辑"""
        base_delay = 60 / bpm  # 每拍基础时间
        channel = mixer.Channel(0)  # 使用独立音频通道
        try:
            flag = 0
            if index == -1:
                index = 0
            for idx, note in enumerate(notes):
                # 跳过索引前的音节
                if idx != index and flag == 0:
                    continue
                flag = 1
                if not self.preview_playing:
                    break

                # 更新界面显示
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
                self.edit_window.update()

                # 节拍非空，播放对应声音
                beat, block = note['beat'], note['block']
                if beat > 0 and block != 0:
                    if block in self.sound_blocks:
                        try:
                            if channel.get_busy():  # 停止前一个音
                                channel.stop()
                            channel.play(self.sound_blocks[block])
                        except pygame.error as e:
                            messagebox.showerror("错误", f"播放失败：{str(e)}")
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
        index = self.selected_index
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.see(index)

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
            'coordinate': (0, 0, 0, 0),  # 存储匹配区域坐标
            'blocks': [(0, 0)] * 16,  # 16个区块的状态
            'current_note': -1,  # 表示当前正在处理或播放的音符
            'hotkeys': None,  # 存储热键配置信息
            'bpm': 60,  # 演奏BPM
            'mouse': 10  # 鼠标抖动（模拟人类鼠标抖动，还没有在程序中实现）
        }
        self.note_labels = {'beat': [], 'block': []}  # 新增初始化
        self.setup_listeners()
        self.check_window_active()

    @staticmethod
    def create_window():
        """创建主窗口"""
        window = tk.Tk()
        window.title(the_title)
        window.geometry("700x750")
        window.columnconfigure(0, weight=1)
        return window

    def init_ui(self):
        """初始化界面组件"""
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
        ttk.Button(cal_frame, text="开始定位按键", command=self.locate_coordinate).grid(row=0, column=0)
        self.cal_label = ttk.Label(cal_frame, text="未定位", foreground='darkgray')
        self.cal_label.grid(row=0, column=2, padx=10, sticky='nsew')

        # 第三行 坐标矩阵
        grid_frame = ttk.LabelFrame(self.window, text="校准坐标")
        grid_frame.grid(row=2, column=0, padx=10, pady=5)

        self.grid_labels = []
        for i in range(16):
            lbl = ttk.Label(grid_frame, text="(0,0)", width=10, relief='ridge')
            lbl.grid(row=i // 4, column=i % 4, padx=2, pady=2, sticky='nsew')
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
        self.bpm_entry = ttk.Entry(play_setting_frame, width=4)
        self.bpm_entry.insert(0, "60")
        self.bpm_entry.grid(row=0, column=1, padx=5, sticky='nsew')
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=6).grid(row=0, column=2, sticky='nsew')
        # 抖动设置
        ttk.Label(play_setting_frame, text="鼠标抖动").grid(row=0, column=3, sticky='e')
        self.mouse_move = ttk.Entry(play_setting_frame, width=4)
        self.mouse_move.insert(0, "3")
        self.mouse_move.grid(row=0, column=4, padx=5)
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=6).grid(row=0, column=5, sticky='e')

        # 在UI中添加以下控件
        ttk.Label(play_setting_frame, text="灵敏度").grid(row=0, column=6)
        self.sensitivity_entry = ttk.Entry(play_setting_frame, width=4)
        self.sensitivity_entry.insert(0, "1.65")
        self.sensitivity_entry.grid(row=0, column=7)

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
        """更新状态栏"""
        self.status_label.config(text=message, foreground=color)

    def setup_listeners(self):
        """设置事件监听"""
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        self.state['hotkeys'] = GlobalHotkey(self.start_playing, self.toggle_pause, self.stop_playing)

    def check_window_active(self):
        """5-1、定时检测窗口激活状态"""
        # if self.state['playing'] and not self.is_window_active():
        #     self.toggle_pause()
        #     self.update_status("窗口未激活，自动暂停", 'orange')
        # self.window.after(1000, self.check_window_active)
        pass

    def is_window_active(self):
        """5-2、检测游戏窗口是否激活"""
        try:
            active_hwnd = win32gui.GetForegroundWindow()
            return active_hwnd == self.state['hwnd']
        except:
            return False

    """-----------------以下为实际功能-----------------"""

    def toggle_topmost(self):
        """切换置顶状态"""
        current = self.window.attributes('-topmost')
        self.window.attributes('-topmost', not current)

    def capture_window(self):
        """捕捉游戏窗口"""
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

    def locate_coordinate(self):
        """开始坐标校准"""
        if not self.state['hwnd']:
            self.update_status("请先捕捉窗口", 'red')
            return
        flag = 0
        try:
            location_button = pyautogui.locateOnScreen('key.png', confidence=0.8)
            left, top, width, height = location_button
            button_region = (int(left), int(top), int(width), int(height))
            self.state['coordinate'] = button_region
            self.cal_label.config(text=f"匹配区域({left}, {top}),({width}, {height})", foreground='green')
            flag = 1
        except Exception as e:
            self.update_status(f"区域定位失败:{str(e)}", 'red')

        if flag:
            try:
                # 找到区域截图
                im = pyautogui.screenshot(region=button_region)
                # 将 PIL 图像转换为 NumPy 数组
                image = np.array(im)
                gray_image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
                if gray_image is None:
                    messagebox.showinfo("错误", "无法加载图像")
                    return
                # 固定阈值分割
                _, binary_image = cv2.threshold(gray_image, thresh=80, maxval=255, type=cv2.THRESH_BINARY)

                # 创建结构元素
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))  # 矩形结构元素
                # 开运算
                opened_image = cv2.morphologyEx(binary_image, cv2.MORPH_OPEN, kernel, iterations=1)

                # 使用findContours获取轮廓
                contours, _ = cv2.findContours(opened_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                # 完成过滤和质心计算
                centers = []
                for cnt in contours:
                    M = cv2.moments(cnt)
                    area = M['m00']     # 轮廓的面积
                    if area >= 500:     # 面积过滤
                        cx = int(M['m10'] / area)   # 轮廓对 x 轴的矩
                        cy = int(M['m01'] / area)   # 轮廓对 y 轴的矩
                        centers.append((cx, cy))
                if len(centers) != 16:
                    messagebox.showinfo("错误", f"检测到 {len(centers)} 个有效轮廓，应为16")
                    return

                # 按Y坐标排序
                sorted_y = sorted(centers, key=lambda c: c[1])
                # 4个一组依次排序顺序赋值
                count = 3
                center_x = []
                for idx, (x, y) in enumerate(sorted_y):
                    if count:
                        center_x.append((x, y))
                        count -= 1
                        continue
                    center_x.append((x, y))
                    sorted_x = sorted(center_x, key=lambda c: c[0])
                    for i in range(4):
                        x, y = sorted_x[i]
                        absolute_x = int(left) + int(x)
                        absolute_y = int(top) + int(y)
                        index = idx-(3-i)
                        self.state['blocks'][index] = (absolute_x, absolute_y)
                        self.grid_labels[index].config(text=f"({absolute_x},{absolute_y})")
                    count = 3
                    center_x = []
                self.update_status(f"坐标已更新", 'green')
            except Exception as e:
                print(e)
                self.update_status(f"按键定位失败:{str(e)}", 'red')

    def load_sheet(self):
        """加载乐谱"""
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
                # 单独调用回到开头
                self.sheet_canvas.xview_moveto(0)
                self.sheet_editor.refresh_list()  # 主动刷新编辑器

            # 读取提示
            self.update_status(f"成功加载乐谱: {os.path.basename(path)}", 'green')
        except Exception as e:
            self.update_status(f"加载失败: {str(e)}", 'red')

    def start_playing(self):
        """开始演奏"""
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
        """切换暂停状态"""
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
        """停止演奏"""
        self.state['playing'] = False
        self.state['paused'] = False
        self.highlight_note(-1)
        self.sheet_canvas.xview_moveto(0)  # 新增滚动重置
        self.update_status("演奏已停止", 'gray')

    def update_bpm(self):
        """更新BPM"""
        try:
            bpm = int(self.bpm_entry.get())
            self.state['bpm'] = bpm
            self.update_status(f"BPM已修改为{bpm}", 'green')
        except ValueError:
            self.update_status("BPM必须为整数", 'red')

    def update_mouse(self):
        """更新鼠标抖动"""
        try:
            mouse_move = int(self.mouse_move.get())
            self.state['mouse'] = mouse_move
            self.update_status(f"鼠标抖动已修改{mouse_move}", 'green')
        except ValueError:
            self.update_status("抖动必须为整数", 'red')

    def play_notes(self):
        """演奏核心逻辑"""
        # 修改键位映射：实际键位编号 → 原索引
        index_map = [
            12, 13, 14, 15,  # 坐标存储索引0-3 → 原索引12-15（对应13-16）
            8, 9, 10, 11,  # 坐标存储4-7 → 原索引8-11（对应9-12）
            4, 5, 6, 7,  # 坐标存储8-11 → 原索引4-7（对应5-8）
            0, 1, 2, 3  # 坐标存储12-15 → 原索引0-3（对应1-4）
        ]
        start_time = time.time()
        try:
            bpm = int(self.state['bpm'])
            delay = 60 / bpm
            # 初始移动中心
            midx, midy = pyautogui.position()
            for idx, note in enumerate(self.sheet_data):
                if not self.state['playing']: break
                # 处理暂停和窗口激活检测
                while True:
                    if not self.state['playing']:
                        return
                    if not self.state['paused'] and self.is_window_active():
                        break
                self.update_status(f"开始演奏", 'red')
                # 节拍同步  记录空节拍时间
                target_time = abs(note['beat']) * delay
                while time.time() - start_time < target_time:
                    if not self.state['playing'] or self.state['paused']: break
                # 界面更新
                self.window.after(0, self.highlight_note, idx)
                self.window.after(0, lambda: self.sheet_canvas.xview_moveto(idx / len(self.sheet_data)))
                # 非空节拍才演奏
                stime = time.time()
                if note['beat'] > 0:
                    try:
                        mouse_shift = float(self.mouse_move.get())  # 从界面获取鼠标抖动
                        sensitivity = float(self.sensitivity_entry.get())  # 从界面获取灵敏度
                        block = note['block'] - 1   # 取得索引
                        original_index = index_map[block]   # 修改索引映射逻辑
                        target_x, target_y = self.state['blocks'][original_index]
                        # 计算目标节点与初始移动中心的距离相对鼠标偏移量
                        dx = int((target_x - midx - random.gauss(0, mouse_shift)) * sensitivity)
                        dy = int((target_y - midy - random.gauss(0, mouse_shift)) * sensitivity)
                        # print(dx, dy)
                        # 移动鼠标点击
                        move_time = random.uniform(target_time * 0.75, target_time * 0.85)
                        print(move_time)
                        # 移动相对位置
                        pydirectinput.moveRel(dx, dy, relative=True, duration=move_time, tween=pyautogui.easeInOutQuad)
                        pydirectinput.click()
                        # 结束时间
                        etime = time.time()
                        wait_time = etime - stime
                        print(wait_time)
                        # 补全一拍等待时间，时间同步
                        midx, midy = target_x, target_y
                        if wait_time < target_time:
                            time.sleep(target_time-wait_time)
                    except Exception as e:
                        print(f"移动出错：{str(e)}")
                        self.update_status(f"移动出错：{str(e)}", 'red')
                else:
                    # 空节拍延时
                    time.sleep(target_time)
        except Exception as e:
            self.update_status(f"演奏出错：{str(e)}", 'red')
        finally:
            self.stop_playing()

    def highlight_note(self, idx):
        """高亮当前音符"""
        if self.state['current_note'] >= 0:     # 清除旧高亮
            for lbl in self.note_labels.values():
                try:
                    lbl[self.state['current_note']].config(background='')
                except:
                    pass
        if idx >= 0:    # 设置新高亮
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
                beat_lbl = ttk.Label(self.sheet_table, text=f"{note['beat']:.2f}", width=6)
                beat_lbl.grid(row=0, column=col_idx, padx=2, pady=2)
                self.note_labels['beat'].append(beat_lbl)
                block_lbl = ttk.Label(self.sheet_table, text=note['block'], width=6)
                block_lbl.grid(row=1, column=col_idx, padx=2, pady=2)
                self.note_labels['block'].append(block_lbl)

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
    # 必须用管理员模式才能捕捉游戏窗口
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 0)
        sys.exit()
    app = MusicAutoPlayer()
    app.window.mainloop()
