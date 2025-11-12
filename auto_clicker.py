#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自動點擊工具 - Auto Clicker
功能:在指定座標自動點擊滑鼠,支援 GUI 介面控制
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import json
import os
import platform
from datetime import datetime
import pyautogui
import keyboard

# macOS 多桌面支援
try:
    from AppKit import NSWorkspace, NSApplication
    from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionAll, kCGNullWindowID
    MACOS_DESKTOP_SUPPORT = True
except ImportError:
    MACOS_DESKTOP_SUPPORT = False


class ConfigManager:
    """設定檔管理類別"""

    def __init__(self, config_file='config.json'):
        self.config_file = config_file

    def save_config(self, x, y, interval, max_clicks=0):
        """儲存設定到 JSON 檔案"""
        try:
            config = {
                'x': x,
                'y': y,
                'interval': interval,
                'max_clicks': max_clicks,
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"儲存設定失敗: {e}")
            return False

    def load_config(self):
        """從 JSON 檔案載入設定"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return (config.get('x'), config.get('y'),
                       config.get('interval'), config.get('max_clicks', 0))
            return None, None, None, None
        except Exception as e:
            print(f"載入設定失敗: {e}")
            return None, None, None, None


class StatisticsTracker:
    """統計追蹤類別"""

    def __init__(self):
        self.click_count = 0
        self.start_time = None

    def start(self):
        """開始計時"""
        self.click_count = 0
        self.start_time = time.time()

    def increment(self):
        """增加計數"""
        self.click_count += 1

    def get_elapsed_time(self):
        """取得已運行時間 (格式化為 HH:MM:SS)"""
        if self.start_time is None:
            return "00:00:00"
        elapsed = int(time.time() - self.start_time)
        hours = elapsed // 3600
        minutes = (elapsed % 3600) // 60
        seconds = elapsed % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def reset(self):
        """重置統計"""
        self.click_count = 0
        self.start_time = None


class DesktopMonitor:
    """桌面狀態監控類別 (檢測視窗是否在當前活躍的桌面)"""

    def __init__(self, root_window):
        self.root = root_window
        self.macos_support = MACOS_DESKTOP_SUPPORT
        self._is_on_active_desktop = True  # 預設在當前桌面

    def is_on_active_desktop(self):
        """檢查視窗是否在當前活躍的桌面

        使用 CGWindowListCopyWindowInfo 檢查視窗是否在螢幕上可見。
        當視窗在其他桌面時，kCGWindowIsOnscreen 會是 False。
        """
        if not self.macos_support:
            return True  # 非 macOS 系統，總是返回 True

        try:
            # 基本檢查：視窗狀態
            state = self.root.state()
            if state == 'iconic':  # 最小化
                return False

            if not self.root.winfo_viewable():
                return False

            # 使用 CGWindowListCopyWindowInfo 檢查視窗是否在螢幕上
            # 這是最可靠的方法來判斷視窗是否在當前桌面
            try:
                current_pid = os.getpid()

                # 獲取所有視窗的資訊
                window_list = CGWindowListCopyWindowInfo(
                    kCGWindowListOptionAll,
                    kCGNullWindowID
                )

                # 檢查我們的進程是否有視窗在螢幕上可見
                # 注意：CGWindowListCopyWindowInfo 返回的是一個 CFArray
                # 在 PyObjC 中會自動轉換為 Python list

                has_onscreen_window = False

                for window in window_list:
                    # 檢查視窗所屬的進程 ID
                    window_pid = window.get('kCGWindowOwnerPID', 0)

                    if window_pid == current_pid:
                        # 檢查視窗是否在螢幕上可見
                        # kCGWindowIsOnscreen: 當視窗在當前桌面時為 True，在其他桌面時為 False
                        is_onscreen = window.get('kCGWindowIsOnscreen', False)

                        # 檢查視窗層級（排除特殊層級的視窗）
                        window_layer = window.get('kCGWindowLayer', 0)

                        # 只考慮正常層級的視窗（layer = 0）
                        if is_onscreen and window_layer == 0:
                            has_onscreen_window = True
                            break

                return has_onscreen_window

            except Exception as e:
                print(f"CGWindowListCopyWindowInfo 檢測失敗: {e}")
                # API 失敗時，使用備用方法
                # 如果視窗可見且未最小化，保守地認為在當前桌面
                return self.root.winfo_viewable() and state != 'iconic'

        except Exception as e:
            print(f"檢查桌面狀態時發生錯誤: {e}")
            return True  # 發生錯誤時預設為在當前桌面


class CoordinateCapture:
    """座標擷取類別"""

    def __init__(self, callback):
        self.callback = callback
        self.capturing = False

    def start_capture(self):
        """啟動座標擷取模式"""
        self.capturing = True
        # 在新執行緒中等待點擊
        thread = threading.Thread(target=self._wait_for_click, daemon=True)
        thread.start()

    def _wait_for_click(self):
        """等待使用者點擊並捕捉座標"""
        time.sleep(0.5)  # 短暫延遲,讓使用者可以移動滑鼠

        # 等待使用者按下滑鼠左鍵
        while self.capturing:
            # 檢查是否按下左鍵
            try:
                # 使用 pyautogui 取得當前滑鼠位置
                if keyboard.is_pressed('button'):  # 檢測點擊
                    x, y = pyautogui.position()
                    self.capturing = False
                    self.callback(x, y)
                    break
            except:
                pass
            time.sleep(0.01)

    def stop_capture(self):
        """停止擷取"""
        self.capturing = False


class ClickController:
    """點擊控制類別"""

    def __init__(self, statistics_tracker, desktop_monitor=None):
        self.statistics = statistics_tracker
        self.desktop_monitor = desktop_monitor
        self.running = False
        self.paused = False
        self.click_thread = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.auto_stop_callback = None  # 自動停止回調

    def start_clicking(self, x, y, interval, update_callback, max_clicks=0, auto_stop_callback=None):
        """啟動自動點擊"""
        if self.running:
            return False

        self.running = True
        self.paused = False
        self.stop_event.clear()
        self.pause_event.clear()
        self.statistics.start()
        self.auto_stop_callback = auto_stop_callback

        # 在新執行緒中執行點擊迴圈
        self.click_thread = threading.Thread(
            target=self._click_loop,
            args=(x, y, interval, update_callback, max_clicks),
            daemon=True
        )
        self.click_thread.start()
        return True

    def stop_clicking(self):
        """停止自動點擊"""
        self.running = False
        self.paused = False
        self.stop_event.set()
        self.pause_event.set()  # 確保暫停狀態被解除

    def toggle_pause(self):
        """切換暫停/恢復狀態"""
        if not self.running:
            return False

        self.paused = not self.paused
        if not self.paused:
            self.pause_event.set()  # 恢復執行
        else:
            self.pause_event.clear()  # 暫停
        return True

    def _click_loop(self, x, y, interval, update_callback, max_clicks=0):
        """點擊迴圈 (在獨立執行緒中運行)"""
        actual_click_count = 0  # 實際點擊次數計數

        while self.running and not self.stop_event.is_set():
            # 檢查是否手動暫停
            if self.paused:
                self.pause_event.wait()  # 等待恢復信號
                if not self.running:
                    break

            try:
                # 【桌面檢查】只在視窗位於當前桌面時執行點擊
                should_click = True
                if self.desktop_monitor:
                    should_click = self.desktop_monitor.is_on_active_desktop()

                # 只在當前桌面執行點擊
                if should_click:
                    # 執行點擊
                    pyautogui.click(x, y)
                    self.statistics.increment()
                    actual_click_count += 1

                    # 呼叫更新回調函數
                    if update_callback:
                        update_callback()

                    # 【自動停止】檢查是否達到點擊上限
                    if max_clicks > 0 and actual_click_count >= max_clicks:
                        print(f"已達到點擊上限 {max_clicks} 次,自動停止")
                        self.running = False
                        # 呼叫自動停止回調
                        if self.auto_stop_callback:
                            self.auto_stop_callback()
                        break

                # 等待間隔時間 (無論是否點擊都要等待)
                if not self.stop_event.wait(interval):
                    continue
                else:
                    break

            except Exception as e:
                print(f"點擊時發生錯誤: {e}")
                break

        self.running = False


class AutoClickerGUI:
    """自動點擊工具的 GUI 介面"""

    def __init__(self, root):
        self.root = root
        self.root.title("自動點擊工具 - Auto Clicker")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        # 【多桌面支援】不使用跨桌面置頂,改為單桌面浮動
        # 註: 在 macOS 上,-topmost True 會讓視窗出現在所有桌面
        # 我們改為使用 lift() 提升視窗層級,但不跨桌面
        if MACOS_DESKTOP_SUPPORT:
            # macOS: 不設定 topmost,使用 lift 保持在當前桌面最上層
            self.root.lift()
            # 定期更新視窗層級
            self._keep_window_raised()
        else:
            # 其他系統: 使用原本的置頂方式
            self.root.attributes('-topmost', True)

        # 初始化各模組
        self.config_manager = ConfigManager()
        self.statistics = StatisticsTracker()

        # 【多桌面支援】初始化桌面監控
        self.desktop_monitor = DesktopMonitor(self.root) if MACOS_DESKTOP_SUPPORT else None

        # 初始化點擊控制器,傳入桌面監控
        self.click_controller = ClickController(self.statistics, self.desktop_monitor)
        self.coordinate_capture = None

        # 建立 GUI 元件
        self._create_widgets()

        # 載入上次的設定
        self._load_last_config()

        # 註冊 F9 熱鍵
        self._register_hotkey()

        # 啟動統計更新定時器
        self._update_statistics()

    def _create_widgets(self):
        """建立 GUI 元件"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # ===== 座標設定區 =====
        coord_frame = ttk.LabelFrame(main_frame, text="座標設定", padding="10")
        coord_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(coord_frame, text="X 座標:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.x_var = tk.StringVar(value="100")
        ttk.Entry(coord_frame, textvariable=self.x_var, width=15).grid(row=0, column=1, padx=5)

        ttk.Label(coord_frame, text="Y 座標:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5))
        self.y_var = tk.StringVar(value="200")
        ttk.Entry(coord_frame, textvariable=self.y_var, width=15).grid(row=0, column=3, padx=5)

        self.capture_btn = ttk.Button(coord_frame, text="選取座標", command=self._start_coordinate_capture)
        self.capture_btn.grid(row=1, column=0, columnspan=4, pady=(10, 0))

        # ===== 參數設定區 =====
        param_frame = ttk.LabelFrame(main_frame, text="參數設定", padding="10")
        param_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(param_frame, text="點擊間隔:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.interval_var = tk.StringVar(value="0.5")
        ttk.Entry(param_frame, textvariable=self.interval_var, width=15).grid(row=0, column=1, padx=5)
        ttk.Label(param_frame, text="秒").grid(row=0, column=2, sticky=tk.W)

        # 【新增】點擊上限設定
        ttk.Label(param_frame, text="點擊上限:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.max_clicks_var = tk.StringVar(value="0")
        ttk.Entry(param_frame, textvariable=self.max_clicks_var, width=15).grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Label(param_frame, text="次 (0=無限)").grid(row=1, column=2, sticky=tk.W, pady=(5, 0))

        # ===== 控制區 =====
        control_frame = ttk.LabelFrame(main_frame, text="控制區", padding="10")
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        self.start_btn = ttk.Button(control_frame, text="開始點擊", command=self._start_clicking)
        self.start_btn.grid(row=0, column=0, padx=5, pady=5)

        self.stop_btn = ttk.Button(control_frame, text="停止", command=self._stop_clicking, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5, pady=5)

        # 緊急停止按鈕 (醒目的紅色按鈕)
        self.emergency_btn = tk.Button(
            control_frame,
            text="⚠️ 緊急停止",
            command=self._emergency_stop,
            bg='#ff4444',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief=tk.RAISED,
            bd=2,
            activebackground='#cc0000',
            activeforeground='white'
        )
        self.emergency_btn.grid(row=0, column=2, padx=5, pady=5)

        self.save_btn = ttk.Button(control_frame, text="儲存設定", command=self._save_config)
        self.save_btn.grid(row=0, column=3, padx=5, pady=5)

        self.load_btn = ttk.Button(control_frame, text="載入設定", command=self._load_config)
        self.load_btn.grid(row=0, column=4, padx=5, pady=5)

        # ===== 統計資訊區 =====
        stats_frame = ttk.LabelFrame(main_frame, text="統計資訊", padding="10")
        stats_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(stats_frame, text="已點擊次數:").grid(row=0, column=0, sticky=tk.W)
        self.click_count_label = ttk.Label(stats_frame, text="0 次", font=('Arial', 12, 'bold'))
        self.click_count_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))

        ttk.Label(stats_frame, text="運行時間:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.time_label = ttk.Label(stats_frame, text="00:00:00", font=('Arial', 12, 'bold'))
        self.time_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(5, 0))

        # ===== 提示區 =====
        hint_frame = ttk.Frame(main_frame)
        hint_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # 根據平台顯示對應的緊急停止熱鍵
        emergency_key = "Cmd+Shift+Q" if platform.system() == 'Darwin' else "Ctrl+Shift+Q"
        hint_label = ttk.Label(hint_frame, text=f"提示: 按 F9 暫停/恢復 | 按 {emergency_key} 緊急停止", foreground="gray")
        hint_label.grid(row=0, column=0)

    def _start_coordinate_capture(self):
        """開始座標擷取"""
        self.capture_btn.config(state=tk.DISABLED, text="請點擊目標位置...")
        messagebox.showinfo("座標擷取", "請在螢幕上點擊您想要自動點擊的位置")

        # 暫時最小化視窗
        self.root.iconify()

        # 等待一段時間後取得滑鼠位置
        def capture_after_delay():
            time.sleep(1)  # 等待1秒
            x, y = pyautogui.position()
            self._on_coordinate_captured(x, y)

        thread = threading.Thread(target=capture_after_delay, daemon=True)
        thread.start()

    def _on_coordinate_captured(self, x, y):
        """座標擷取完成的回調"""
        self.x_var.set(str(x))
        self.y_var.set(str(y))
        self.root.deiconify()  # 恢復視窗
        self.capture_btn.config(state=tk.NORMAL, text="選取座標")
        messagebox.showinfo("座標擷取成功", f"已設定座標為 ({x}, {y})")

    def _start_clicking(self):
        """開始自動點擊"""
        try:
            x = int(self.x_var.get())
            y = int(self.y_var.get())
            interval = float(self.interval_var.get())
            max_clicks = int(self.max_clicks_var.get())

            if interval <= 0:
                messagebox.showerror("錯誤", "點擊間隔必須大於 0")
                return

            if max_clicks < 0:
                messagebox.showerror("錯誤", "點擊上限不能小於 0")
                return

            # 啟動點擊,傳入自動停止回調
            if self.click_controller.start_clicking(
                x, y, interval,
                self._update_statistics,
                max_clicks=max_clicks,
                auto_stop_callback=self._on_auto_stop
            ):
                self.start_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.NORMAL)
                self.capture_btn.config(state=tk.DISABLED)
                # 更新視窗標題
                self.root.title("自動點擊工具 - 執行中")

        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數值")

    def _stop_clicking(self):
        """停止自動點擊"""
        self.click_controller.stop_clicking()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.capture_btn.config(state=tk.NORMAL)
        # 恢復視窗標題
        self.root.title("自動點擊工具 - Auto Clicker")

    def _on_auto_stop(self):
        """自動停止回調 (達到點擊上限時)"""
        # 在主執行緒中更新 GUI
        self.root.after(0, self._auto_stop_gui_update)

    def _auto_stop_gui_update(self):
        """自動停止的 GUI 更新"""
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.capture_btn.config(state=tk.NORMAL)
        self.root.title("自動點擊工具 - Auto Clicker")
        messagebox.showinfo("自動停止", f"已達到點擊上限,自動停止\n總點擊次數: {self.statistics.click_count}")

    def _emergency_stop(self):
        """緊急停止 (Cmd/Ctrl+Shift+Q 熱鍵)"""
        if self.click_controller.running:
            self.click_controller.stop_clicking()
            # 在主執行緒中更新 GUI
            self.root.after(0, self._emergency_stop_gui_update)

    def _emergency_stop_gui_update(self):
        """緊急停止的 GUI 更新"""
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.capture_btn.config(state=tk.NORMAL)
        self.root.title("自動點擊工具 - Auto Clicker")

        # 顯示對應平台的熱鍵
        hotkey = "Cmd+Shift+Q" if platform.system() == 'Darwin' else "Ctrl+Shift+Q"
        messagebox.showinfo("已停止", f"自動點擊已緊急停止 ({hotkey})")

    def _toggle_pause(self):
        """切換暫停/恢復 (由 F9 熱鍵觸發)"""
        if self.click_controller.toggle_pause():
            if self.click_controller.paused:
                self.stop_btn.config(text="恢復")
            else:
                self.stop_btn.config(text="停止")

    def _save_config(self):
        """儲存設定"""
        try:
            x = int(self.x_var.get())
            y = int(self.y_var.get())
            interval = float(self.interval_var.get())
            max_clicks = int(self.max_clicks_var.get())

            if self.config_manager.save_config(x, y, interval, max_clicks):
                messagebox.showinfo("成功", "設定已儲存")
            else:
                messagebox.showerror("錯誤", "儲存設定失敗")
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數值")

    def _load_config(self):
        """載入設定"""
        x, y, interval, max_clicks = self.config_manager.load_config()
        if x is not None:
            self.x_var.set(str(x))
            self.y_var.set(str(y))
            self.interval_var.set(str(interval))
            self.max_clicks_var.set(str(max_clicks if max_clicks is not None else 0))
            messagebox.showinfo("成功", "設定已載入")
        else:
            messagebox.showwarning("警告", "找不到設定檔")

    def _load_last_config(self):
        """自動載入上次的設定"""
        x, y, interval, max_clicks = self.config_manager.load_config()
        if x is not None:
            self.x_var.set(str(x))
            self.y_var.set(str(y))
            self.interval_var.set(str(interval))
            self.max_clicks_var.set(str(max_clicks if max_clicks is not None else 0))

    def _update_statistics(self):
        """更新統計顯示"""
        self.click_count_label.config(text=f"{self.statistics.click_count:,} 次")
        self.time_label.config(text=self.statistics.get_elapsed_time())

        # 每 100 毫秒更新一次
        self.root.after(100, self._update_statistics)

    def _keep_window_raised(self):
        """定期提升視窗層級 (macOS 單桌面置頂)"""
        try:
            self.root.lift()
        except:
            pass
        # 每 1 秒更新一次
        self.root.after(1000, self._keep_window_raised)

    def _register_hotkey(self):
        """註冊全域熱鍵"""
        success_count = 0
        failed_keys = []

        try:
            # F9: 暫停/恢復
            try:
                keyboard.add_hotkey('f9', self._toggle_pause)
                success_count += 1
            except Exception as e:
                failed_keys.append(f"F9 ({e})")

            # Cmd/Ctrl+Shift+Q: 緊急停止
            # macOS 使用 cmd，其他系統使用 ctrl
            emergency_key = 'cmd+shift+q' if platform.system() == 'Darwin' else 'ctrl+shift+q'
            try:
                keyboard.add_hotkey(emergency_key, self._emergency_stop)
                success_count += 1
            except Exception as e:
                failed_keys.append(f"{emergency_key} ({e})")
        except Exception as e:
            print(f"註冊熱鍵時發生錯誤: {e}")

        # 如果有熱鍵註冊失敗，顯示警告
        if failed_keys:
            warning_msg = "部分熱鍵註冊失敗:\n" + "\n".join(failed_keys)
            warning_msg += "\n\n可能需要授予「輔助使用」權限"
            warning_msg += "\n請到「系統偏好設定 > 安全性與隱私權 > 輔助使用」"
            print(warning_msg)
            # 延遲顯示對話框，避免在初始化時阻塞
            self.root.after(1000, lambda: messagebox.showwarning("熱鍵註冊警告", warning_msg))

    def _on_closing(self):
        """視窗關閉事件"""
        if self.click_controller.running:
            if messagebox.askokcancel("確認", "自動點擊正在運行,確定要關閉嗎?"):
                self.click_controller.stop_clicking()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """主程式進入點"""
    root = tk.Tk()
    app = AutoClickerGUI(root)
    root.protocol("WM_DELETE_WINDOW", app._on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
