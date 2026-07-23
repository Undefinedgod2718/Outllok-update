import os
import sys
import time
from mcp_outlook_core import (
    run_outlook_contact_update,
    ImportMode,
    is_outlook_installed,
    is_outlook_running,
    close_outlook_process
)

EXAMPLE_RAR = r"C:\Users\richard_zhang\Downloads\260702mailaddress.rar"

def test_all_three_functions():
    print("==================================================")
    print(" 🚀 開始執行三大核心功能實機測試 (Real Machine Test)")
    print(f" 測試檔案: {EXAMPLE_RAR}")
    print("==================================================")

    if not os.path.exists(EXAMPLE_RAR):
        print(f"❌ 錯誤: 測試檔案不存在: {EXAMPLE_RAR}")
        sys.exit(1)

    if not is_outlook_installed():
        print("❌ 錯誤: 系統未安裝 Outlook COM 服務！")
        sys.exit(1)

    if is_outlook_running():
        print("⚠️ 偵測到 Outlook 正在執行中，自動進行程序處理...")
        close_outlook_process()
        time.sleep(2)

    # --------------------------------------------------
    # 測試 1: OVERWRITE (覆蓋匯入)
    # --------------------------------------------------
    print("\n--------------------------------------------------")
    print(" 測試 1/3: 覆蓋匯入 (OVERWRITE Mode)")
    print(" 預期: 全備份 -> 清空現有聯絡人 -> 匯入 367 筆新聯絡人")
    print("--------------------------------------------------")

    def progress_cb1(msg, pct):
        print(f"  [Progress {pct or 0}%] {msg}")

    res1 = run_outlook_contact_update(EXAMPLE_RAR, ImportMode.OVERWRITE, progress_cb1)
    print(f"  結果 Summary: {res1}")

    assert res1["success"] is True, f"OVERWRITE mode failed: {res1.get('error')}"
    assert res1["pst_items_count"] == 367, f"PST count mismatch: {res1['pst_items_count']}"
    assert res1["imported_count"] == 367, f"Imported count mismatch: {res1['imported_count']}"
    assert res1["final_contacts_count"] == 367, f"Final count mismatch: {res1['final_contacts_count']}"
    print("✅ 測試 1 (OVERWRITE 覆蓋匯入) 通過！")

    time.sleep(2)

    # --------------------------------------------------
    # 測試 2: COEXIST (共存匯入)
    # --------------------------------------------------
    print("\n--------------------------------------------------")
    print(" 測試 2/3: 共存匯入 (COEXIST Mode)")
    print(" 預期: 全備份 (367筆) -> 保留舊資料 -> 追加 367 筆 -> 總計 734 筆 (356 筆 Email 重複聯絡人)")
    print("--------------------------------------------------")

    def progress_cb2(msg, pct):
        print(f"  [Progress {pct or 0}%] {msg}")

    res2 = run_outlook_contact_update(EXAMPLE_RAR, ImportMode.COEXIST, progress_cb2)
    print(f"  結果 Summary: {res2}")

    assert res2["success"] is True, f"COEXIST mode failed: {res2.get('error')}"
    assert res2["backed_up_count"] == 367, f"Backup count mismatch: {res2['backed_up_count']}"
    assert res2["imported_count"] == 367, f"Imported count mismatch: {res2['imported_count']}"
    assert res2["final_contacts_count"] == 734, f"Final count mismatch: {res2['final_contacts_count']}"
    assert res2["duplicate_count"] == 356, f"Duplicate count mismatch: {res2['duplicate_count']}"
    print("✅ 測試 2 (COEXIST 共存匯入) 通過！")

    time.sleep(2)

    # --------------------------------------------------
    # 測試 3: SMART_MERGE (比對匯入)
    # --------------------------------------------------
    print("\n--------------------------------------------------")
    print(" 測試 3/3: 比對匯入 (SMART_MERGE Mode)")
    print(" 預期: 全備份 (734筆) -> 比對跳過重複項 -> 僅匯入非重複群組/聯絡人 -> 匯入完成")
    print("--------------------------------------------------")

    def progress_cb3(msg, pct):
        print(f"  [Progress {pct or 0}%] {msg}")

    res3 = run_outlook_contact_update(EXAMPLE_RAR, ImportMode.SMART_MERGE, progress_cb3)
    print(f"  結果 Summary: {res3}")

    assert res3["success"] is True, f"SMART_MERGE mode failed: {res3.get('error')}"
    assert res3["skipped_count"] >= 356, f"Skipped count mismatch: {res3['skipped_count']}"
    print("✅ 測試 3 (SMART_MERGE 比對匯入) 通過！")

    time.sleep(2)

    # --------------------------------------------------
    # 收尾復原: 使用 OVERWRITE 將聯絡人回復為標準 367 筆
    # --------------------------------------------------
    print("\n--------------------------------------------------")
    print(" 🧹 執行最終收尾: 將聯絡人重置為乾淨的 367 筆標準庫...")
    print("--------------------------------------------------")
    res_cleanup = run_outlook_contact_update(EXAMPLE_RAR, ImportMode.OVERWRITE, None)
    assert res_cleanup["success"] is True
    assert res_cleanup["final_contacts_count"] == 367
    print("✅ 最終復原完成！預設聯絡人資料庫狀態正常 (367 筆)。")

    print("\n==================================================")
    print(" 🎉 三大 Function 實機測試全部 100% 成功通過！")
    print("==================================================")

if __name__ == "__main__":
    test_all_three_functions()
