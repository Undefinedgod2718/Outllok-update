import os
import sys
import subprocess

def build_exe():
    print("==================================================")
    print(" 🛠️ 開始打包 Outlook 聯絡人自動更新精靈 v2.1.0 (Single EXE)")
    print("==================================================")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    wizard_script = os.path.join(base_dir, "wizard_gui.py")
    tools_dir = os.path.join(base_dir, "tools", "UnRAR.exe")
    sounds_dir = os.path.join(base_dir, "sounds")
    assets_dir = os.path.join(base_dir, "assets")
    ico_file = os.path.join(assets_dir, "logo.ico")

    if not os.path.exists(tools_dir):
        print(f"❌ 錯誤: 找不到 UnRAR 工具: {tools_dir}")
        sys.exit(1)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconsole",
        f"--icon={ico_file}",
        "--name=Outlook_Contact_Updater_Wizard",
        f"--add-data={tools_dir};tools",
        f"--add-data={sounds_dir};sounds",
        f"--add-data={assets_dir};assets",
        wizard_script
    ]

    print(f"執行 PyInstaller 指令: {' '.join(cmd)}")
    res = subprocess.run(cmd, cwd=base_dir)

    if res.returncode == 0:
        exe_path = os.path.join(base_dir, "dist", "Outlook_Contact_Updater_Wizard.exe")
        print("\n==================================================")
        print(" 🎉 打包成功！獨立可執行檔已產出 (含 logo.ico 圖示與黑體 UI 圖章):")
        print(f" 📦 檔案路徑: {exe_path}")
        print("==================================================")
    else:
        print("\n❌ 打包失敗！請檢查錯誤日誌。")
        sys.exit(1)

if __name__ == "__main__":
    build_exe()
