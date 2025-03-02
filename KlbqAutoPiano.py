import ctypes
import sys
import tkinter as tk
from tkinter import ttk, filedialog
import json
import threading
import time
import random
import pyautogui
import win32gui
import win32con
from pynput import keyboard, mouse

the_title = "卡拉彼丘琴房助手 v1.2 (25.3.2)"

class GlobalHotkey:
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
                    time.sleep(0.01)
                    h.join(0.1)

        self.listener_thread = threading.Thread(target=listen, daemon=True)
        self.listener_thread.start()

    def stop_listener(self):
        self.running = False
        if self.hotkeys:
            self.hotkeys.stop()

class MusicAutoPlayer:
    def __init__(self):
        self.window = self.create_window()
        self.init_ui()
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
        window.geometry("700x700")
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

        self.load_button = ttk.Button(play_frame, text="加载乐谱", command=self.load_sheet)
        self.load_button.grid(row=0, column=0, sticky='nsew')

        self.start_button = ttk.Button(play_frame, text="开始 (F10)", command=self.start_playing)
        self.start_button.grid(row=0, column=1, sticky='nsew')

        self.pause_button = ttk.Button(play_frame, text="⏸ 暂停 (F11)", command=self.toggle_pause)
        self.pause_button.grid(row=0, column=2, sticky='nsew')

        self.stop_button = ttk.Button(play_frame, text="■ 停止 (F12)", command=self.stop_playing)
        self.stop_button.grid(row=0, column=3, sticky='nsew')


        # 第五行 演奏设置
        play_setting_frame = ttk.LabelFrame(self.window, text="演奏设置")
        play_setting_frame.grid(row=4, column=0, padx=10, pady=5, sticky='nsew')
        # BPM设置
        ttk.Label(play_setting_frame, text="BPM").grid(row=0, column=0, sticky='nsew')
        self.bpm_entry = ttk.Entry(play_setting_frame, width=8)
        self.bpm_entry.insert(0, "60")
        self.bpm_entry.grid(row=0, column=1, padx=5, sticky='nsew')
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=8).grid(row=0, column=2, sticky='nsew')
        # 抖动设置
        ttk.Label(play_setting_frame, text="鼠标抖动").grid(row=0, column=3, sticky='nsew')
        self.mouse_move = ttk.Entry(play_setting_frame, width=8)
        self.mouse_move.insert(0, "10")
        self.mouse_move.grid(row=0, column=4, padx=5)
        ttk.Button(play_setting_frame, text="修改", command=self.update_bpm, width=8).grid(row=0, column=5, sticky='nsew')

        # 第六行 乐谱显示
        sheet_frame = ttk.LabelFrame(self.window, text="乐谱")
        sheet_frame.grid(row=5, column=0, padx=10, pady=5, sticky='nsew')  # 调整到第7行

        # 左侧固定行名
        left_header = ttk.Frame(sheet_frame)
        left_header.pack(side='left', fill='y')
        ttk.Label(left_header, text="节拍", width=4, relief="raised").grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(left_header, text="方块", width=4, relief="raised").grid(row=1, column=0, padx=5, pady=5)

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
        self.window.after(500, self.check_window_active)

    def is_window_active(self):
        """5-2、检测游戏窗口是否激活"""
        try:
            active_hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(active_hwnd)
            print(title)
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
            with open(path) as f:
                data = json.load(f)
                # 添加数据有效性验证
                if not all(k in data for k in ('notes', 'bpm')):
                    raise ValueError("Invalid sheet format")

                self.sheet_data = sorted(data['notes'], key=lambda x: x['beat'])
                self.bpm_entry.delete(0, tk.END)
                self.bpm_entry.insert(0, str(data['bpm']))

            # 清除旧标签（修复索引越界问题）
            for col in self.note_labels.values():
                for widget in col:
                    widget.destroy()
            self.note_labels['beat'].clear()
            self.note_labels['block'].clear()

            # 创建新标签（调整列索引）
            for col_idx, note in enumerate(self.sheet_data):
                beat_lbl = ttk.Label(self.sheet_table, text=f"{note['beat']:.2f}", width=6)
                beat_lbl.grid(row=0, column=col_idx, padx=2, pady=2)
                self.note_labels['beat'].append(beat_lbl)

                block_lbl = ttk.Label(self.sheet_table, text=note['block'], width=6)
                block_lbl.grid(row=1, column=col_idx, padx=2, pady=2)
                self.note_labels['block'].append(block_lbl)

            self.sheet_canvas.xview_moveto(0)
            self.update_status(f"已加载乐谱：{path.split('/')[-1]}，BPM{data['bpm']}", 'green')
        except Exception as e:
            self.update_status(f"加载失败：{str(e)}", 'red')

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
        try:
            bpm = int(self.state['bpm'])
            delay = 60 / bpm

            for idx, note in enumerate(self.sheet_data):
                if not self.state['playing']: break
                # 处理暂停和窗口激活检测
                while True:
                    if not self.state['playing']: return
                    if not self.state['paused'] and self.is_window_active():
                        break
                    time.sleep(0.1)
                # 节拍同步
                target_time = note['beat'] * delay
                while time.time() - start_time < target_time:
                    if not self.state['playing'] or self.state['paused']: break
                    time.sleep(0.001)

                # 界面更新
                self.window.after(0, self.highlight_note, idx)
                self.window.after(0, lambda: self.sheet_canvas.xview_moveto(idx / len(self.sheet_data)))
                block = note['block'] - 1

                # 坐标移动模拟点击
                mouse_shift = self.state.get('mouse')

                x, y = self.state['blocks'][block]
                x += random.gauss(0, mouse_shift)
                y += random.gauss(0, mouse_shift)

                pyautogui.moveTo(x, y, duration=random.uniform(0.1, 0.3))
                pyautogui.click()

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

    def on_close(self):
        if self.state['hotkeys']:
            self.state['hotkeys'].stop_listener()
        self.state['playing'] = False
        self.window.destroy()


if __name__ == "__main__":
    # if ctypes.windll.shell32.IsUserAnAdmin() == 0:
    #     ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    #     sys.exit()
    app = MusicAutoPlayer()
    app.window.mainloop()
