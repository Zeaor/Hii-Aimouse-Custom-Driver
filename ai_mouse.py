import hid
import time
import threading
import webbrowser
import os
import sys
import ctypes 
import json
import socket
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# 嘗試匯入外部函式庫
try:
    import keyboard
except ImportError: keyboard = None
try:
    import speech_recognition as sr
except ImportError: sr = None

# ==============================================================================
# 核心常數
# ==============================================================================
VENDOR_ID = 0x95F1
PRODUCT_ID = 0xA1B6
CONFIG_FILE = "config.json"

# Windows API for DPI
SPI_SETMOUSESPEED = 0x0071
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02

# 預設設定結構
DEFAULT_PROFILE = {
    "mic": {"action": "voice_typing", "param": ""},
    "search": {"action": "open_url", "param": "https://www.google.com"},
    "side": {"action": "key_press", "param": "alt+left"},
    "dpi_fast": 18,
    "dpi_normal": 10
}

# 中英文對照表 (顯示名稱 -> 內部代碼)
ACTION_MAP_DISPLAY = {
    "🗣️ 按住說話 (語音輸入)": "voice_typing",
    "🔗 開啟網頁 (URL)": "open_url",
    "⌨️ 模擬按鍵 (快捷鍵)": "key_press",
    "🖱️ 切換 DPI (速度)": "toggle_dpi"
}
# 反向對照表 (內部代碼 -> 顯示名稱)
ACTION_MAP_INTERNAL = {v: k for k, v in ACTION_MAP_DISPLAY.items()}

# ==============================================================================
# 全域狀態
# ==============================================================================
GLOBAL_CONFIG = {
    "active_profile": "Mode A",
    "profiles": {"Mode A": DEFAULT_PROFILE.copy()}
}
ACTIVE_SETTINGS = GLOBAL_CONFIG["profiles"]["Mode A"]
current_speed_mode = "NORMAL"

# 語音錄製狀態旗標
MIC_IS_HELD = False 

# ==============================================================================
# 系統功能 (Backend)
# ==============================================================================

def set_mouse_speed(speed):
    speed = max(1, min(20, int(speed)))
    try:
        ctypes.windll.user32.SystemParametersInfoA(
            SPI_SETMOUSESPEED, 0, ctypes.c_void_p(speed), 
            SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
        )
    except: pass

def load_config():
    global GLOBAL_CONFIG, ACTIVE_SETTINGS
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "profiles" in data:
                    GLOBAL_CONFIG = data
    except Exception: pass
    
    active = GLOBAL_CONFIG.get("active_profile", "Mode A")
    if active not in GLOBAL_CONFIG["profiles"]:
        GLOBAL_CONFIG["profiles"][active] = DEFAULT_PROFILE.copy()
        GLOBAL_CONFIG["active_profile"] = active
    
    ACTIVE_SETTINGS = GLOBAL_CONFIG["profiles"][active]

def save_config():
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(GLOBAL_CONFIG, f, indent=4, ensure_ascii=False)
        return True
    except: return False

# ==============================================================================
# 動作執行邏輯
# ==============================================================================

def start_voice_recording():
    """按下按鍵時呼叫：開始錄音迴圈"""
    global MIC_IS_HELD
    if not sr: return print("❌ 缺 SpeechRecognition")
    
    MIC_IS_HELD = True # 標記為按下狀態
    
    def _record_task():
        r = sr.Recognizer()
        frames = []
        try:
            with sr.Microphone() as source:
                # 快速調整環境音 (縮短時間以提升反應)
                r.adjust_for_ambient_noise(source, duration=0.3)
                print("🎤 收音中... (放開按鍵結束)")
                
                # 持續錄音直到 MIC_IS_HELD 變為 False
                while MIC_IS_HELD:
                    # 讀取一小段音訊
                    buffer = source.stream.read(source.CHUNK)
                    frames.append(buffer)
                
                print("⏳ 處理中...")
                
                # 將錄下的片段組合成音訊資料
                audio_data = sr.AudioData(b''.join(frames), source.SAMPLE_RATE, source.SAMPLE_WIDTH)
                
                # 辨識
                text = r.recognize_google(audio_data, language="zh-TW")
                print(f"✅ 辨識結果: {text}")
                if keyboard: keyboard.write(text)
                
        except sr.UnknownValueError:
            print("❌ 無法辨識 (聲音太小或不清楚)")
        except Exception as e:
            print(f"❌ 錯誤: {e}")

    threading.Thread(target=_record_task).start()

def stop_voice_recording():
    """放開按鍵時呼叫：停止錄音迴圈"""
    global MIC_IS_HELD
    if MIC_IS_HELD:
        print("🛑 停止收音，開始轉譯...")
        MIC_IS_HELD = False # 這會讓 _record_task 跳出 while 迴圈

def toggle_dpi():
    global current_speed_mode
    fast = ACTIVE_SETTINGS.get("dpi_fast", 18)
    normal = ACTIVE_SETTINGS.get("dpi_normal", 10)
    
    if current_speed_mode == "NORMAL":
        set_mouse_speed(fast)
        current_speed_mode = "FAST"
    else:
        set_mouse_speed(normal)
        current_speed_mode = "NORMAL"
    print(f"👉 DPI: {current_speed_mode}")

def execute_button_press(btn_key):
    """按下按鍵時的動作"""
    cfg = ACTIVE_SETTINGS.get(btn_key, {})
    action = cfg.get("action")
    param = cfg.get("param")

    if action == "voice_typing":
        start_voice_recording() # 開始錄音
    elif action == "open_url":
        if param: threading.Thread(target=lambda: webbrowser.open(param)).start()
    elif action == "key_press":
        if keyboard and param: keyboard.send(param)
    elif action == "toggle_dpi":
        toggle_dpi()

def execute_button_release(btn_key):
    """放開按鍵時的動作"""
    cfg = ACTIVE_SETTINGS.get(btn_key, {})
    action = cfg.get("action")
    
    if action == "voice_typing":
        stop_voice_recording() # 停止錄音

# ==============================================================================
# HID 監聽執行緒
# ==============================================================================

def monitor_mouse(path):
    try:
        h = hid.device()
        h.open_path(path)
        h.set_nonblocking(1)
        # 狀態鎖 (防止重複觸發)
        pressed_flags = {33: False, 35: False, 37: False}
        
        while True:
            data = h.read(64)
            if data and len(data) > 5:
                key = data[5]
                
                # --- 按下事件 ---
                if key in [33, 35, 37]:
                    if not pressed_flags.get(key, False):
                        if key == 33: execute_button_press("mic")
                        elif key == 35: execute_button_press("search")
                        elif key == 37: execute_button_press("side")
                        pressed_flags[key] = True
                
                # --- 放開事件 ---
                elif key in [34, 36, 38]:
                    original_key = key - 1
                    if pressed_flags.get(original_key, False):
                        if original_key == 33: execute_button_release("mic")
                        elif original_key == 35: execute_button_release("search")
                        elif original_key == 37: execute_button_release("side")
                        pressed_flags[original_key] = False
            else:
                time.sleep(0.005)
    except: return

# ==============================================================================
# 🖥️ 動態 GUI 介面
# ==============================================================================

class ButtonConfigRow:
    """管理單一行按鍵設定的 UI 邏輯"""
    def __init__(self, parent_frame, button_key, button_label):
        self.button_key = button_key
        self.frame = ttk.LabelFrame(parent_frame, text=button_label, padding="10")
        self.frame.pack(fill="x", pady=5, padx=5)

        # 功能選擇 (顯示中文)
        ttk.Label(self.frame, text="功能:").pack(side="left")
        self.action_var = tk.StringVar()
        self.action_combo = ttk.Combobox(self.frame, textvariable=self.action_var, state="readonly", width=20)
        self.action_combo['values'] = list(ACTION_MAP_DISPLAY.keys()) # 使用中文列表
        self.action_combo.pack(side="left", padx=5)
        self.action_combo.bind("<<ComboboxSelected>>", self.on_action_change)

        # 動態區域
        self.param_frame = ttk.Frame(self.frame)
        self.param_frame.pack(side="left", fill="x", expand=True, padx=10)
        
        self.entry_var = tk.StringVar()
        self.dpi_fast_var = tk.IntVar()
        self.dpi_normal_var = tk.IntVar()

    def load_data(self, profile_data):
        btn_data = profile_data.get(self.button_key, {})
        internal_action = btn_data.get("action", "key_press")
        
        # 將內部代碼轉為中文顯示
        display_text = ACTION_MAP_INTERNAL.get(internal_action, "⌨️ 模擬按鍵 (快捷鍵)")
        self.action_var.set(display_text)
        
        self.entry_var.set(btn_data.get("param", ""))
        self.dpi_fast_var.set(profile_data.get("dpi_fast", 18))
        self.dpi_normal_var.set(profile_data.get("dpi_normal", 10))
        
        self.refresh_dynamic_ui(internal_action)

    def on_action_change(self, event):
        display_text = self.action_var.get()
        internal_action = ACTION_MAP_DISPLAY.get(display_text)
        self.refresh_dynamic_ui(internal_action)

    def refresh_dynamic_ui(self, internal_action):
        for widget in self.param_frame.winfo_children(): widget.destroy()

        if internal_action == "voice_typing":
            ttk.Label(self.param_frame, text="(按住按鍵說話，放開即輸入)", foreground="blue").pack(anchor="w")
        
        elif internal_action == "open_url":
            ttk.Label(self.param_frame, text="網址:").pack(side="left")
            ttk.Entry(self.param_frame, textvariable=self.entry_var, width=35).pack(side="left", padx=5)

        elif internal_action == "key_press":
            ttk.Label(self.param_frame, text="快捷鍵:").pack(side="left")
            ttk.Entry(self.param_frame, textvariable=self.entry_var, width=20).pack(side="left", padx=5)

        elif internal_action == "toggle_dpi":
            ttk.Label(self.param_frame, text="高速:").pack(side="left")
            ttk.Spinbox(self.param_frame, from_=1, to=20, textvariable=self.dpi_fast_var, width=3).pack(side="left")
            ttk.Label(self.param_frame, text="一般:").pack(side="left", padx=(10,0))
            ttk.Spinbox(self.param_frame, from_=1, to=20, textvariable=self.dpi_normal_var, width=3).pack(side="left")

    def get_ui_data(self):
        display_text = self.action_var.get()
        internal_action = ACTION_MAP_DISPLAY.get(display_text)
        
        data = {"action": internal_action, "param": ""}
        if internal_action in ["open_url", "key_press"]:
            data["param"] = self.entry_var.get()
        
        dpi_data = None
        if internal_action == "toggle_dpi":
            dpi_data = {
                "dpi_fast": self.dpi_fast_var.get(), 
                "dpi_normal": self.dpi_normal_var.get()
            }
        return data, dpi_data


class MouseApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Hii Aimouse 控制中心")
        self.root.geometry("700x550") # 加寬一點以容納中文
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        top_frame = ttk.Frame(root, padding="10")
        top_frame.pack(fill="x")
        
        ttk.Label(top_frame, text="當前模式:", font=("Microsoft JhengHei", 10, "bold")).pack(side="left")
        
        self.profile_var = tk.StringVar(value=GLOBAL_CONFIG["active_profile"])
        self.profile_combo = ttk.Combobox(top_frame, textvariable=self.profile_var, state="readonly")
        self.profile_combo['values'] = list(GLOBAL_CONFIG["profiles"].keys())
        self.profile_combo.pack(side="left", padx=10)
        self.profile_combo.bind("<<ComboboxSelected>>", self.change_profile)

        ttk.Button(top_frame, text="➕ 新增", command=self.add_profile).pack(side="left")
        ttk.Button(top_frame, text="🗑️ 刪除", command=self.del_profile).pack(side="left", padx=5)

        self.settings_frame = ttk.Frame(root)
        self.settings_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.rows = {}
        self.rows["mic"] = ButtonConfigRow(self.settings_frame, "mic", "🎤 麥克風鍵 (33)")
        self.rows["search"] = ButtonConfigRow(self.settings_frame, "search", "🔍 搜尋鍵 (35)")
        self.rows["side"] = ButtonConfigRow(self.settings_frame, "side", "👉 側邊鍵 (37)")

        btm_frame = ttk.Frame(root, padding="15")
        btm_frame.pack(fill="x", side="bottom")
        self.save_btn = ttk.Button(btm_frame, text="💾 儲存設定並套用", command=self.save_all, width=25)
        self.save_btn.pack(side="right")
        
        self.load_profile_to_gui(GLOBAL_CONFIG["active_profile"])

    def load_profile_to_gui(self, profile_name):
        data = GLOBAL_CONFIG["profiles"].get(profile_name, DEFAULT_PROFILE)
        for key, row in self.rows.items(): row.load_data(data)

    def change_profile(self, event=None):
        new_profile = self.profile_var.get()
        self.load_profile_to_gui(new_profile)
        GLOBAL_CONFIG["active_profile"] = new_profile

    def add_profile(self):
        name = simpledialog.askstring("新增", "輸入新模式名稱:")
        if name and name not in GLOBAL_CONFIG["profiles"]:
            GLOBAL_CONFIG["profiles"][name] = GLOBAL_CONFIG["profiles"][GLOBAL_CONFIG["active_profile"]].copy()
            self.profile_combo['values'] = list(GLOBAL_CONFIG["profiles"].keys())
            self.profile_var.set(name)
            self.change_profile()
            
    def del_profile(self):
        name = self.profile_var.get()
        if name == "Mode A": return messagebox.showerror("錯誤", "無法刪除預設模式")
        if messagebox.askyesno("確認", f"刪除模式 {name}?"):
            del GLOBAL_CONFIG["profiles"][name]
            self.profile_combo['values'] = list(GLOBAL_CONFIG["profiles"].keys())
            self.profile_var.set("Mode A")
            self.change_profile()

    def save_all(self):
        current = self.profile_var.get()
        new_data = {}
        current_data = GLOBAL_CONFIG["profiles"][current]
        dpi_settings = {"dpi_fast": current_data.get("dpi_fast", 18), "dpi_normal": current_data.get("dpi_normal", 10)}

        for key, row in self.rows.items():
            btn_data, row_dpi = row.get_ui_data()
            new_data[key] = btn_data
            if row_dpi: dpi_settings = row_dpi

        final_profile_data = {**new_data, **dpi_settings}
        GLOBAL_CONFIG["profiles"][current] = final_profile_data
        GLOBAL_CONFIG["active_profile"] = current
        
        global ACTIVE_SETTINGS
        ACTIVE_SETTINGS = final_profile_data
        
        if save_config():
            orig = self.save_btn['text']
            self.save_btn['text'] = "✅ 已儲存！"
            self.root.after(1000, lambda: self.save_btn.configure(text=orig))
            set_mouse_speed(ACTIVE_SETTINGS.get("dpi_normal", 10))

    def on_close(self):
        self.save_all()
        self.root.destroy()
        os._exit(0)

# ==============================================================================
# 主程式入口
# ==============================================================================
def main():
    load_config()
    
    devices = [d for d in hid.enumerate() if d['vendor_id'] == VENDOR_ID and d['product_id'] == PRODUCT_ID]
    if not devices:
        print("❌ 未偵測到滑鼠，GUI 僅供編輯模式。")
    else:
        print(f"🔥 滑鼠監聽中... ({len(devices)} 介面)")
        for dev in devices:
            t = threading.Thread(target=monitor_mouse, args=(dev['path'],))
            t.daemon = True
            t.start()

    root = tk.Tk()
    app = MouseApp(root)
    if keyboard:
        keyboard.add_hotkey('ctrl+alt+shift+q', lambda: app.on_close(), suppress=True)
    root.mainloop()

if __name__ == "__main__":
    main()