import os
import sys
import time
import shutil
import subprocess
import tempfile
import datetime
import pythoncom
import win32com.client
from enum import Enum

class ImportMode(Enum):
    OVERWRITE = "OVERWRITE"       # 覆蓋匯入: 全備份 -> 清空 -> 匯入
    COEXIST = "COEXIST"           # 共存匯入: 全備份 -> 保留舊資料 -> 追加匯入
    SMART_MERGE = "SMART_MERGE"   # 比對匯入: 全備份 -> 比對 Email/姓名 -> 僅匯入全新資料

def is_outlook_installed() -> bool:
    """Check if Outlook is installed via COM registry dispatch."""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        return ns is not None
    except Exception:
        return False
    finally:
        pythoncom.CoUninitialize()

def is_outlook_running() -> bool:
    """Check if OUTLOOK.EXE process is currently running."""
    try:
        output = subprocess.check_output('tasklist /FI "IMAGENAME eq OUTLOOK.EXE"', shell=True, text=True)
        return "OUTLOOK.EXE" in output.upper()
    except Exception:
        return False

def close_outlook_process() -> bool:
    """Try to close OUTLOOK.EXE cleanly."""
    try:
        subprocess.call("taskkill /F /IM OUTLOOK.EXE", shell=True)
        return True
    except Exception:
        return False

def get_tools_unrar_path() -> str:
    """Locate UnRAR executable (either inside bundled tools/ or system WinRAR)."""
    # 1. Bundled tools folder
    base_dir = os.path.dirname(os.path.abspath(__file__))
    bundled_unrar = os.path.join(base_dir, "tools", "UnRAR.exe")
    if os.path.exists(bundled_unrar):
        return bundled_unrar
    
    # 2. PyInstaller sys._MEIPASS path
    if hasattr(sys, "_MEIPASS"):
        meipass_unrar = os.path.join(sys._MEIPASS, "tools", "UnRAR.exe")
        if os.path.exists(meipass_unrar):
            return meipass_unrar

    # 3. System WinRAR path
    sys_winrar = r"C:\Program Files\WinRAR\UnRAR.exe"
    if os.path.exists(sys_winrar):
        return sys_winrar

    sys_rar = r"C:\Program Files\WinRAR\rar.exe"
    if os.path.exists(sys_rar):
        return sys_rar
        
    return ""

def extract_archive(archive_path: str, extract_to: str) -> str:
    """Extract .rar, .zip, or directly return .pst file path."""
    if not os.path.exists(archive_path):
        raise FileNotFoundError(f"Archive file not found: {archive_path}")

    ext = os.path.splitext(archive_path)[1].lower()

    if ext == ".pst":
        return archive_path

    if ext == ".zip":
        shutil.unpack_archive(archive_path, extract_to, "zip")
    elif ext == ".rar":
        unrar_exe = get_tools_unrar_path()
        if not unrar_exe:
            raise RuntimeError("UnRAR.exe not found! Cannot extract .rar archive.")
        
        cmd = f'"{unrar_exe}" x -y "{archive_path}" "{extract_to}\\"'
        res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if res.returncode != 0:
            raise RuntimeError(f"UnRAR failed with code {res.returncode}: {res.stderr or res.stdout}")
    else:
        raise ValueError(f"Unsupported file format: {ext}. Only .rar, .zip, and .pst are supported.")

    # Search for .pst file in extracted directory
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.lower().endswith(".pst"):
                return os.path.join(root, f)

    raise FileNotFoundError(f"No .pst file found inside archive: {archive_path}")

def safe_copy_and_move(item, target_folder, max_retries=3):
    """Copy an Outlook COM item and move it to target folder with retries."""
    for attempt in range(max_retries):
        try:
            copied = item.Copy()
            copied.Move(target_folder)
            return True
        except Exception as ex:
            if attempt == max_retries - 1:
                raise ex
            time.sleep(0.1)

def run_outlook_contact_update(
    archive_or_pst_path: str,
    mode: ImportMode,
    progress_callback=None
) -> dict:
    """
    Core function to process contact update:
    1. Extract PST from archive if needed.
    2. Connect Outlook MAPI COM.
    3. Backup current contacts to a timestamped subfolder.
    4. Execute mode logic (OVERWRITE, COEXIST, SMART_MERGE).
    5. Perform duplicate check & verification report.
    """
    def log(msg, pct=None):
        print(f"[OutlookUpdater] {msg}")
        if progress_callback:
            progress_callback(msg, pct)

    log("初始化環境與檢測檔案...", 5)
    temp_dir = tempfile.mkdtemp(prefix="outlook_update_")
    
    pythoncom.CoInitialize()
    try:
        # Step 1: Extract PST
        log(f"解壓與載入通訊錄檔案: {os.path.basename(archive_or_pst_path)}...", 15)
        pst_path = extract_archive(archive_or_pst_path, temp_dir)
        log(f"成功取得 PST 檔案: {pst_path}", 25)

        # Step 2: Initialize COM MAPI
        log("連接 Outlook MAPI 服務...", 35)
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        default_folder = ns.GetDefaultFolder(10) # olFolderContacts
        old_items = default_folder.Items
        old_count = old_items.Count

        # Step 3: Backup
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder_name = f"連絡人_Backup_{timestamp}"
        log(f"建立安全備份資料夾 [{backup_folder_name}] (目前項目: {old_count})...", 45)
        backup_folder = default_folder.Folders.Add(backup_folder_name)

        copied_backup_count = 0
        for i in range(1, old_count + 1):
            item = old_items.Item(i)
            safe_copy_and_move(item, backup_folder)
            copied_backup_count += 1

        if backup_folder.Items.Count != old_count:
            raise RuntimeError(f"備份數量驗證失敗！預期 {old_count}，實際備份 {backup_folder.Items.Count}。操作終止。")
        log(f"安全備份完成！共備份 {backup_folder.Items.Count} 筆項目。", 55)

        # Step 4: Mount PST
        log("掛載 PST 通訊錄數據源...", 65)
        ns.AddStore(pst_path)

        pst_store = None
        for i in range(1, ns.Stores.Count + 1):
            s = ns.Stores.Item(i)
            if s.FilePath and os.path.normpath(s.FilePath) == os.path.normpath(pst_path):
                pst_store = s
                break
        if not pst_store:
            fn = os.path.basename(pst_path).lower()
            for i in range(1, ns.Stores.Count + 1):
                s = ns.Stores.Item(i)
                if s.FilePath and os.path.basename(s.FilePath).lower() == fn:
                    pst_store = s
                    break

        if not pst_store:
            raise RuntimeError(f"無法存取掛載之 PST 數據庫: {pst_path}")

        pst_root = pst_store.GetRootFolder()
        pst_contacts_folder = None
        for i in range(1, pst_root.Folders.Count + 1):
            f = pst_root.Folders.Item(i)
            if f.DefaultItemType == 2: # olContactItem
                pst_contacts_folder = f
                break

        if not pst_contacts_folder:
            ns.RemoveStore(pst_root)
            raise RuntimeError("PST 檔案中未發現聯絡人資料夾！")

        pst_items = pst_contacts_folder.Items
        pst_count = pst_items.Count
        log(f"PST 中共有 {pst_count} 筆聯絡人資料。", 70)

        imported_count = 0
        skipped_count = 0

        # Step 5: Execute selected mode
        if mode == ImportMode.OVERWRITE:
            log("【覆蓋模式】正在清空預設聯絡人資料夾...", 75)
            for i in range(default_folder.Items.Count, 0, -1):
                default_folder.Items.Item(i).Delete()

            log("【覆蓋模式】正在匯入全數新聯絡人...", 80)
            for i in range(1, pst_count + 1):
                item = pst_items.Item(i)
                safe_copy_and_move(item, default_folder)
                imported_count += 1

        elif mode == ImportMode.COEXIST:
            log("【共存模式】保持現有聯絡人，直接追加新聯絡人...", 80)
            for i in range(1, pst_count + 1):
                item = pst_items.Item(i)
                safe_copy_and_move(item, default_folder)
                imported_count += 1

        elif mode == ImportMode.SMART_MERGE:
            log("【比對模式】建立現有聯絡人索引並比對重複項...", 75)
            existing_emails = set()
            existing_names = set()

            for i in range(1, default_folder.Items.Count + 1):
                cur = default_folder.Items.Item(i)
                if cur.MessageClass == "IPM.Contact":
                    if cur.Email1Address:
                        existing_emails.add(cur.Email1Address.lower().strip())
                    if cur.FullName:
                        existing_names.add(cur.FullName.lower().strip())

            log("【比對模式】開始智慧篩選並匯入全新聯絡人...", 80)
            for i in range(1, pst_count + 1):
                item = pst_items.Item(i)
                if item.MessageClass == "IPM.Contact":
                    email = (item.Email1Address or "").lower().strip()
                    name = (item.FullName or "").lower().strip()

                    is_dup = (email and email in existing_emails) or (name and name in existing_names)
                    if is_dup:
                        skipped_count += 1
                    else:
                        safe_copy_and_move(item, default_folder)
                        imported_count += 1
                        if email:
                            existing_emails.add(email)
                        if name:
                            existing_names.add(name)
                else:
                    safe_copy_and_move(item, default_folder)
                    imported_count += 1

        # Unmount store
        ns.RemoveStore(pst_root)

        # Step 6: Verification & Duplicate audit
        log("執行最終數據驗證與重複項目掃描...", 90)
        final_items = default_folder.Items
        final_count = final_items.Count
        
        email_map = {}
        duplicates = []
        for i in range(1, final_count + 1):
            item = final_items.Item(i)
            if item.MessageClass == "IPM.Contact":
                email = item.Email1Address
                if email:
                    ekey = email.lower().strip()
                    if ekey in email_map:
                        duplicates.append({
                            "email": email,
                            "name": item.FullName or "",
                            "existing_name": email_map[ekey]
                        })
                    else:
                        email_map[ekey] = item.FullName or ""

        log("更新流程順利完成！", 100)

        result = {
            "success": True,
            "mode": mode.value,
            "archive_path": archive_or_pst_path,
            "backup_folder": backup_folder_name,
            "backed_up_count": old_count,
            "pst_items_count": pst_count,
            "imported_count": imported_count,
            "skipped_count": skipped_count,
            "final_contacts_count": final_count,
            "duplicate_count": len(duplicates),
            "duplicates": duplicates
        }
        return result

    except Exception as e:
        log(f" ERROR: {str(e)}", 100)
        return {
            "success": False,
            "error": str(e)
        }
    finally:
        pythoncom.CoUninitialize()
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    print("Outlook core module updated with safe_copy_and_move.")
