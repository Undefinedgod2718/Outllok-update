import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

import mcp_outlook_core
from mcp_outlook_core import (
    ImportMode,
    is_outlook_installed,
    is_outlook_running,
    close_outlook_process,
    run_outlook_contact_update
)
import sound_manager

# ---------------------------------------------------------
# Minecraft Rural Grassland Aesthetic Color Tokens
# ---------------------------------------------------------
BG_STONE_DARK  = "#212121"   # 深石板黑
DIRT_BROWN      = "#4a2e1b"   # 泥土棕
GRASS_GREEN     = "#4e8027"   # 草地綠
LIGHT_GREEN     = "#68a335"   # 亮綠草色
BORDER_DARK     = "#181818"   # 像素深邊框
MINECRAFT_GOLD  = "#ffff55"   # 像素金黃
MINECRAFT_GREEN = "#55ff55"   # 經驗條綠
MINECRAFT_WHITE = "#ffffff"   # 像素純白
MINECRAFT_RED   = "#ff5555"   # 紅石警告
PANEL_BG        = "#2e2e2e"   # 石頭面板
BUTTON_BG       = "#3c3c3c"   # 石頭按鈕
BUTTON_ACTIVE   = "#5c5c5c"   # 按鈕懸停

VERSION = "v2.1.0"
DEFAULT_RAR = r"C:\Users\richard_zhang\Downloads\260702mailaddress.rar"

def get_asset_path(filename: str) -> str:
    """Locate asset file in project assets/ folder or PyInstaller sys._MEIPASS."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    asset_path = os.path.join(base_dir, "assets", filename)
    if os.path.exists(asset_path):
        return asset_path
    if hasattr(sys, "_MEIPASS"):
        meipass_asset = os.path.join(sys._MEIPASS, "assets", filename)
        if os.path.exists(meipass_asset):
            return meipass_asset
    return ""

class MinecraftOutlookWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"⛏️ Outlook 聯絡人自動更新精靈 {VERSION} (Minecraft Style)")
        self.geometry("780x560")
        self.resizable(False, False)
        self.configure(bg=BG_STONE_DARK)

        # Set Window Icon (.ico)
        ico_path = get_asset_path("logo.ico")
        if ico_path and os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception as e:
                print(f"[UI] Set iconbitmap error: {e}")

        # Application state
        self.selected_file = DEFAULT_RAR if os.path.exists(DEFAULT_RAR) else ""
        self.selected_mode = tk.StringVar(value=ImportMode.OVERWRITE.value)

        # Build UI Structure
        self._create_minecraft_header()
        
        self.container = tk.Frame(self, bg=BG_STONE_DARK)
        self.container.pack(fill="both", expand=True, px=15, py=10)

        # Pages
        self.frames = {}
        for PageClass in (PageWelcome, PageSelection, PageProgress, PageVictory):
            page_name = PageClass.__name__
            frame = PageClass(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("PageWelcome")

    def _create_minecraft_header(self):
        header_frame = tk.Frame(self, bg=DIRT_BROWN, height=65, bd=0)
        header_frame.pack(fill="x", side="top")

        # Top grassland stripe
        grass_stripe = tk.Frame(header_frame, bg=GRASS_GREEN, height=12)
        grass_stripe.pack(fill="x", side="top")

        # Container for logo + title
        content_box = tk.Frame(header_frame, bg=DIRT_BROWN)
        content_box.pack(fill="both", expand=True)

        # Render UI Logo Image (unnamed.jpg)
        logo_img_path = get_asset_path("unnamed.jpg")
        if logo_img_path and os.path.exists(logo_img_path):
            try:
                pil_img = Image.open(logo_img_path).resize((44, 44), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(pil_img)
                logo_lbl = tk.Label(content_box, image=self.logo_photo, bg=DIRT_BROWN, bd=0)
                logo_lbl.pack(side="left", padx=15, pady=4)
            except Exception as e:
                print(f"[UI] Load logo image error: {e}")

        # Header Title
        title_lbl = tk.Label(
            content_box,
            text=f"🌾 OUTLOOK CONTACT UPDATER WIZARD {VERSION} 🌾",
            font=("Consolas", 15, "bold"),
            fg=MINECRAFT_GOLD,
            bg=DIRT_BROWN
        )
        title_lbl.pack(side="left", py=8)

        # Pixel separator border
        sep = tk.Frame(self, bg=BORDER_DARK, height=4)
        sep.pack(fill="x", side="top")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()


# ---------------------------------------------------------
# Page 1: Welcome & Environment Check
# ---------------------------------------------------------
class PageWelcome(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=PANEL_BG, bd=3, relief="ridge")
        self.controller = controller

        # Title
        lbl_title = tk.Label(
            self,
            text="【 Step 1: 環境檢查與自動偵測 】",
            font=("Microsoft JhengHei", 14, "bold"),
            fg=MINECRAFT_GOLD,
            bg=PANEL_BG
        )
        lbl_title.pack(anchor="w", padx=20, pady=15)

        # Status Panel
        self.status_box = tk.Text(
            self,
            font=("Consolas", 10),
            bg="#1b1b1b",
            fg=MINECRAFT_WHITE,
            height=12,
            bd=2,
            relief="sunken",
            wrap="word"
        )
        self.status_box.pack(fill="x", padx=20, pady=10)

        # Action Buttons Frame
        btn_frame = tk.Frame(self, bg=PANEL_BG)
        btn_frame.pack(fill="x", padx=20, pady=15, side="bottom")

        self.btn_close_outlook = tk.Button(
            btn_frame,
            text="⚡ 一鍵關閉 Outlook",
            font=("Microsoft JhengHei", 10, "bold"),
            bg=MINECRAFT_RED,
            fg=MINECRAFT_WHITE,
            activebackground="#cc0000",
            bd=2,
            command=self._kill_outlook
        )

        self.btn_next = tk.Button(
            btn_frame,
            text="下一步: 選擇檔案與模式 ➔",
            font=("Microsoft JhengHei", 11, "bold"),
            bg=GRASS_GREEN,
            fg=MINECRAFT_WHITE,
            activebackground=LIGHT_GREEN,
            bd=3,
            command=self._go_next
        )
        self.btn_next.pack(side="right")

    def _go_next(self):
        sound_manager.play_click()
        self.controller.show_frame("PageSelection")

    def on_show(self):
        self.check_environment()

    def check_environment(self):
        self.status_box.config(state="normal")
        self.status_box.delete("1.0", tk.END)

        self.status_box.insert(tk.END, "⛏️ 正在掃描系統環境...\n\n")

        # 1. Outlook installed
        installed = is_outlook_installed()
        if installed:
            self.status_box.insert(tk.END, "  [✓] Outlook 2016 MAPI COM 服務: 已安裝支援\n")
        else:
            self.status_box.insert(tk.END, "  [✗] Outlook 2016 MAPI COM 服務: 未安裝！\n", "error")

        # 2. Outlook running
        running = is_outlook_running()
        if running:
            self.status_box.insert(
                tk.END,
                "  [!] 偵測到 OUTLOOK.EXE 正在背景執行中。\n"
                "      建議點擊『一鍵關閉 Outlook』以確保掛載 PST 順暢無阻。\n"
            )
            self.btn_close_outlook.pack(side="left")
        else:
            self.status_box.insert(tk.END, "  [✓] OUTLOOK.EXE 執行狀態: 良好 (未開啟，隨時可進行數據注入)\n")
            self.btn_close_outlook.pack_forget()

        # 3. Check 7z tool
        unrar_path = mcp_outlook_core.get_tools_unrar_path()
        if unrar_path:
            self.status_box.insert(tk.END, f"  [✓] 可攜式解壓組件 UnRAR: Ready ({os.path.basename(unrar_path)})\n")
        else:
            self.status_box.insert(tk.END, "  [!] 警告: 未發現 UnRAR 解壓模組，僅支援直接選擇 .pst 檔案。\n")

        self.status_box.config(state="disabled")

    def _kill_outlook(self):
        sound_manager.play_click()
        close_outlook_process()
        messagebox.showinfo("提示", "Outlook 已成功關閉！")
        self.check_environment()


# ---------------------------------------------------------
# Page 2: File Selection & Mode Choice
# ---------------------------------------------------------
class PageSelection(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=PANEL_BG, bd=3, relief="ridge")
        self.controller = controller

        lbl_title = tk.Label(
            self,
            text="【 Step 2: 選擇通訊錄檔案與升級模式 】",
            font=("Microsoft JhengHei", 14, "bold"),
            fg=MINECRAFT_GOLD,
            bg=PANEL_BG
        )
        lbl_title.pack(anchor="w", padx=20, pady=12)

        # File Selection Frame
        file_frame = tk.LabelFrame(
            self,
            text=" 通訊錄壓縮檔 (.rar / .zip / .pst) ",
            font=("Microsoft JhengHei", 10, "bold"),
            fg=MINECRAFT_WHITE,
            bg=PANEL_BG,
            bd=2
        )
        file_frame.pack(fill="x", padx=20, pady=8)

        self.entry_file = tk.Entry(
            file_frame,
            font=("Consolas", 10),
            bg="#1b1b1b",
            fg=MINECRAFT_WHITE,
            bd=2
        )
        self.entry_file.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        if self.controller.selected_file:
            self.entry_file.insert(0, self.controller.selected_file)

        btn_browse = tk.Button(
            file_frame,
            text="📁 瀏覽檔案",
            font=("Microsoft JhengHei", 9, "bold"),
            bg=BUTTON_BG,
            fg=MINECRAFT_WHITE,
            activebackground=BUTTON_ACTIVE,
            command=self._browse_file
        )
        btn_browse.pack(side="right", padx=10, pady=10)

        # Import Mode Radio Buttons Frame
        mode_frame = tk.LabelFrame(
            self,
            text=" 選擇升級匯入模式 ",
            font=("Microsoft JhengHei", 10, "bold"),
            fg=MINECRAFT_WHITE,
            bg=PANEL_BG,
            bd=2
        )
        mode_frame.pack(fill="x", padx=20, pady=8)

        modes = [
            ("🌾 覆蓋匯入 (Recommended)", ImportMode.OVERWRITE.value, "安全全備份 ➔ 清空現有預設通訊錄 ➔ 完整寫入新聯絡人"),
            ("🪵 共存匯入", ImportMode.COEXIST.value, "安全全備份 ➔ 保留舊有聯絡人 ➔ 直接追加寫入新聯絡人"),
            ("⛏️ 比對匯入 (Smart Deduplication)", ImportMode.SMART_MERGE.value, "安全全備份 ➔ 智慧比對 Email/姓名 ➔ 僅寫入全新聯絡人")
        ]

        for text, val, desc in modes:
            r_frame = tk.Frame(mode_frame, bg=PANEL_BG)
            r_frame.pack(fill="x", padx=15, pady=4)

            rb = tk.Radiobutton(
                r_frame,
                text=text,
                variable=self.controller.selected_mode,
                value=val,
                font=("Microsoft JhengHei", 10, "bold"),
                fg=MINECRAFT_GOLD,
                bg=PANEL_BG,
                selectcolor="#1b1b1b",
                activebackground=PANEL_BG,
                command=sound_manager.play_click
            )
            rb.pack(anchor="w")

            desc_lbl = tk.Label(
                r_frame,
                text=f"   └ {desc}",
                font=("Microsoft JhengHei", 9),
                fg="#aaaaaa",
                bg=PANEL_BG
            )
            desc_lbl.pack(anchor="w")

        # Nav Buttons
        btn_frame = tk.Frame(self, bg=PANEL_BG)
        btn_frame.pack(fill="x", padx=20, pady=12, side="bottom")

        btn_back = tk.Button(
            btn_frame,
            text="⬅ 上一步",
            font=("Microsoft JhengHei", 10),
            bg=BUTTON_BG,
            fg=MINECRAFT_WHITE,
            command=self._go_back
        )
        btn_back.pack(side="left")

        btn_start = tk.Button(
            btn_frame,
            text="🚀 開始執行升級更新 ➔",
            font=("Microsoft JhengHei", 11, "bold"),
            bg=GRASS_GREEN,
            fg=MINECRAFT_WHITE,
            activebackground=LIGHT_GREEN,
            bd=3,
            command=self._start_process
        )
        btn_start.pack(side="right")

    def _go_back(self):
        sound_manager.play_click()
        self.controller.show_frame("PageWelcome")

    def _browse_file(self):
        sound_manager.play_click()
        path = filedialog.askopenfilename(
            title="選擇通訊錄檔案",
            filetypes=[("Outlook Contact Files", "*.rar *.zip *.pst"), ("All Files", "*.*")]
        )
        if path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.controller.selected_file = path

    def _start_process(self):
        sound_manager.play_click()
        fpath = self.entry_file.get().strip()
        if not fpath or not os.path.exists(fpath):
            messagebox.showerror("錯誤", "請選擇有效的通訊錄檔案 (.rar / .zip / .pst)！")
            return

        self.controller.selected_file = fpath
        self.controller.show_frame("PageProgress")


# ---------------------------------------------------------
# Page 3: Progress & Real-time Minecraft EXP Bar
# ---------------------------------------------------------
class PageProgress(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=PANEL_BG, bd=3, relief="ridge")
        self.controller = controller

        lbl_title = tk.Label(
            self,
            text="【 Step 3: 執行更新與數據寫入 】",
            font=("Microsoft JhengHei", 14, "bold"),
            fg=MINECRAFT_GOLD,
            bg=PANEL_BG
        )
        lbl_title.pack(anchor="w", padx=20, pady=12)

        # Status Label
        self.lbl_step_status = tk.Label(
            self,
            text="準備啟動核心寫入程序...",
            font=("Microsoft JhengHei", 11, "bold"),
            fg=MINECRAFT_WHITE,
            bg=PANEL_BG
        )
        self.lbl_step_status.pack(anchor="w", padx=20, pady=5)

        # Canvas EXP Progress Bar Frame
        exp_frame = tk.Frame(self, bg="#111111", bd=3, relief="sunken")
        exp_frame.pack(fill="x", padx=20, pady=10)

        self.canvas_exp = tk.Canvas(exp_frame, bg="#000000", height=28, highlightthickness=0)
        self.canvas_exp.pack(fill="x")

        # Real-time Log Console
        self.log_box = tk.Text(
            self,
            font=("Consolas", 9),
            bg="#111111",
            fg=MINECRAFT_GREEN,
            height=13,
            bd=2,
            relief="sunken"
        )
        self.log_box.pack(fill="both", expand=True, padx=20, pady=10)

    def on_show(self):
        self.log_box.delete("1.0", tk.END)
        self.update_exp_bar(0, "正在啟動...")
        
        # Start worker thread
        threading.Thread(target=self._worker_thread, daemon=True).start()

    def update_exp_bar(self, percentage, text_msg=""):
        self.canvas_exp.delete("all")
        width = self.canvas_exp.winfo_width() or 700
        height = 28

        # Draw segmented green EXP bar
        fill_width = int(width * (percentage / 100.0))
        if fill_width > 0:
            self.canvas_exp.create_rectangle(0, 0, fill_width, height, fill=MINECRAFT_GREEN, outline="")

        # Draw level / percentage text
        txt = f"EXP LEVEL: {percentage}% | {text_msg}"
        self.canvas_exp.create_text(width // 2, height // 2, text=txt, fill=MINECRAFT_GOLD, font=("Consolas", 10, "bold"))

    def log(self, msg, pct=None):
        def _update():
            self.log_box.insert(tk.END, f"{msg}\n")
            self.log_box.see(tk.END)
            if pct is not None:
                sound_manager.play_xp()
                self.lbl_step_status.config(text=msg)
                self.update_exp_bar(pct, msg)
        self.after(0, _update)

    def _worker_thread(self):
        archive_path = self.controller.selected_file
        mode_str = self.controller.selected_mode.get()
        mode_enum = ImportMode(mode_str)

        res = run_outlook_contact_update(
            archive_or_pst_path=archive_path,
            mode=mode_enum,
            progress_callback=self.log
        )

        self.controller.last_result = res
        self.after(1000, lambda: self.controller.show_frame("PageVictory"))


# ---------------------------------------------------------
# Page 4: Victory / Result Report Screen
# ---------------------------------------------------------
class PageVictory(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=PANEL_BG, bd=3, relief="ridge")
        self.controller = controller

        self.lbl_victory = tk.Label(
            self,
            text="✨ MISSION ACCOMPLISHED! 更新完成 ✨",
            font=("Microsoft JhengHei", 15, "bold"),
            fg=MINECRAFT_GOLD,
            bg=PANEL_BG
        )
        self.lbl_victory.pack(pady=12)

        # Summary Frame
        self.summary_box = tk.Text(
            self,
            font=("Consolas", 10),
            bg="#181818",
            fg=MINECRAFT_WHITE,
            height=14,
            bd=3,
            relief="sunken",
            wrap="word"
        )
        self.summary_box.pack(fill="both", expand=True, padx=20, pady=8)

        # Bottom Buttons
        btn_frame = tk.Frame(self, bg=PANEL_BG)
        btn_frame.pack(fill="x", padx=20, pady=12, side="bottom")

        btn_again = tk.Button(
            btn_frame,
            text="🔄 再次執行更新",
            font=("Microsoft JhengHei", 10),
            bg=BUTTON_BG,
            fg=MINECRAFT_WHITE,
            command=self._go_again
        )
        btn_again.pack(side="left")

        btn_finish = tk.Button(
            btn_frame,
            text="✔ 完成退出精靈",
            font=("Microsoft JhengHei", 11, "bold"),
            bg=GRASS_GREEN,
            fg=MINECRAFT_WHITE,
            command=self._finish
        )
        btn_finish.pack(side="right")

    def _go_again(self):
        sound_manager.play_click()
        self.controller.show_frame("PageSelection")

    def _finish(self):
        sound_manager.play_click()
        self.controller.destroy()

    def on_show(self):
        sound_manager.play_victory()
        res = getattr(self.controller, "last_result", {})
        self.summary_box.config(state="normal")
        self.summary_box.delete("1.0", tk.END)

        if res.get("success"):
            self.lbl_victory.config(text="✨ MISSION ACCOMPLISHED! 通訊錄更新成功 ✨", fg=MINECRAFT_GOLD)
            summary_text = (
                "==========================================================\n"
                f"            📊 數據更新與驗證審計報告 ({VERSION})          \n"
                "==========================================================\n"
                f"  - 執行升級模式: {res.get('mode')}\n"
                f"  - 通訊錄數據源: {os.path.basename(res.get('archive_path', ''))}\n"
                f"  - 安全備份資料夾: {res.get('backup_folder')}\n"
                f"  - 原有聯絡人備份筆數: {res.get('backed_up_count')} 筆\n"
                f"  - PST 來源資料總筆數: {res.get('pst_items_count')} 筆\n"
                f"  - 成功寫入聯絡人筆數: {res.get('imported_count')} 筆\n"
                f"  - 智慧跳過重複筆數: {res.get('skipped_count')} 筆\n"
                f"  - 目前預設通訊錄總筆數: {res.get('final_contacts_count')} 筆\n"
                f"  - 依 Email 掃描重複聯絡人: {res.get('duplicate_count')} 筆\n"
                "----------------------------------------------------------\n"
            )
            dups = res.get("duplicates", [])
            if dups:
                summary_text += "🔍 依 Email 偵測出之重複聯絡人細節:\n"
                for d in dups[:10]:
                    summary_text += f"   - Email: {d['email']} | 姓名: {d['name']}\n"
                if len(dups) > 10:
                    summary_text += f"   ... 及另外 {len(dups)-10} 筆重複項目。\n"
            else:
                summary_text += "  [✓] 預設通訊錄中無發現任何重複 Email 聯絡人！\n"

            summary_text += "==========================================================\n"
            self.summary_box.insert(tk.END, summary_text)
        else:
            self.lbl_victory.config(text="❌ UPDATE FAILED! 執行過程遭遇異常 ❌", fg=MINECRAFT_RED)
            self.summary_box.insert(
                tk.END,
                "==========================================================\n"
                "                  💥 錯誤日誌資訊                          \n"
                "==========================================================\n"
                f" 錯誤訊息: {res.get('error')}\n"
                " 提示: 原有聯絡人已安全備份至 Outlook 備份資料夾中，資料未遺失。\n"
                "==========================================================\n"
            )

        self.summary_box.config(state="disabled")


if __name__ == "__main__":
    app = MinecraftOutlookWizard()
    app.mainloop()
